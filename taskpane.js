/* global Office, Excel, pdfjsLib, Tesseract, PDFLib, SnipLib, FolderLib, TextLib, ExportLib */
"use strict";

const PDFJS_VERSION = "3.11.174";
const LOG_SHEET = "_SNIP_LOG";
const NS_INDEX = "urn:snipper:index:v1";   // elenco dei documenti incorporati
const NS_DOC = "urn:snipper:doc:v1";       // contenuto dei documenti e dell'OCR (a blocchi)
const CHUNK = 2000000;                     // caratteri per blocco
const WARN_MB = 15;
const OCR_SCALE = 2;                       // risoluzione dell'OCR (2 = circa 150 dpi)
const EXT_OK = /\.(pdf|png|jpe?g)$/i;

const state = {
  mode: "snip",
  index: { docs: [], folders: [] },
  log: [],
  docId: null,
  pdf: null, image: null,
  page: 1, numPages: 1,
  scale: 1.6, zoom: 1, dpr: 1, nav: null, lastCell: null, lastNext: null,
  viewW: 0, viewH: 0,
  pageItems: [],            // testo della pagina aperta, in pixel dello schermo
  sumItems: [],
  focusId: null,            // snip evidenziato
  searchMarks: [],          // risultati di ricerca evidenziati
  native: {},               // docId -> {pagina: [pezzi di testo normalizzati]}
  ocr: {},                  // docId -> {pagina: [parole normalizzate]}
  ocrDismissed: new Set(),
  ui: { expanded: new Set(), creatingIn: null, renaming: null, selected: new Set(), currentFolder: "" },
  pendingScroll: false,
  busy: false
};

const $ = id => document.getElementById(id);
const canvas = $("canvas");
const overlay = $("overlay");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
const r4 = n => Math.round(n * 10000) / 10000;
const docById = id => state.index.docs.find(d => d.id === id);
const curDoc = () => docById(state.docId);
const fmtSize = b => b >= 1048576 ? (b / 1048576).toFixed(1).replace(".", ",") + " MB" : Math.max(1, Math.round(b / 1024)) + " KB";

function setStatus(msg, kind) {
  const el = $("status");
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}

function h(tag, props, ...kids) {
  const e = document.createElement(tag);
  Object.entries(props || {}).forEach(([k, v]) => {
    if (v == null || v === false) return;
    if (k === "class") e.className = v;
    else if (k === "text") e.textContent = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? "" : v);
  });
  kids.flat().forEach(c => { if (c != null && c !== false) e.append(c); });
  return e;
}

// =====================================================================
//  Avvio
// =====================================================================
Office.onReady(async info => {
  if (info.host !== Office.HostType.Excel) setStatus("Questo componente funziona solo in Excel.", "err");
  initUI();
  try {
    state.index = await readIndex();
    state.log = await loadLog();
    refreshAll();
    setStatus(state.index.docs.length
      ? `${state.index.docs.length} documento/i incorporato/i in questo file.`
      : "Nessun documento in questo file. Premi «➕ Importa» o trascina qui i file.");
  } catch (e) {
    console.error(e);
    setStatus("Errore all'avvio: " + (e.message || e), "err");
  }
  registerSelectionHandler();
  try { if (Office.addin && Office.addin.setStartupBehavior) await Office.addin.setStartupBehavior(Office.StartupBehavior.load); } catch (e) { /* non supportato */ }
});

async function initPdfWorker() {
  const url = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${PDFJS_VERSION}/pdf.worker.min.js`;
  try {
    const code = await (await fetch(url)).text();
    pdfjsLib.GlobalWorkerOptions.workerSrc = URL.createObjectURL(new Blob([code], { type: "text/javascript" }));
  } catch (e) {
    pdfjsLib.GlobalWorkerOptions.workerSrc = url;
  }
}

function initUI() {
  initPdfWorker();
  $("import").addEventListener("change", ev => { const f = Array.from(ev.target.files || []).map(file => ({ file, folder: state.ui.currentFolder })); ev.target.value = ""; importItems(f); });
  $("importdir").addEventListener("change", ev => {
    const f = Array.from(ev.target.files || []).map(file => {
      const rel = file.webkitRelativePath || file.name;
      const dir = rel.includes("/") ? rel.slice(0, rel.lastIndexOf("/")) : "";
      return { file, folder: FolderLib.cleanPath(FolderLib.join(state.ui.currentFolder, dir)) };
    });
    ev.target.value = "";
    importItems(f);
  });
  $("tab-snip").addEventListener("click", () => showTab("snip"));
  $("tab-org").addEventListener("click", () => { showTab("org"); renderOrg(); });
  $("docselect").addEventListener("change", e => e.target.value && openDoc(e.target.value));
  $("mode-snip").addEventListener("click", () => setMode("snip"));
  $("mode-sum").addEventListener("click", () => setMode("sum"));
  $("prev").addEventListener("click", () => gotoPage(state.page - 1));
  $("next").addEventListener("click", () => gotoPage(state.page + 1));
  $("zoomin").addEventListener("click", () => zoom(1.2));
  $("zoomout").addEventListener("click", () => zoom(1 / 1.2));
  $("findsnip").addEventListener("click", findSnipFromCell);
  $("sumwrite").addEventListener("click", writeSum);
  $("sumclear").addEventListener("click", () => { state.sumItems = []; renderSum(); drawMarks(); });
  $("docfilter").addEventListener("input", renderOrg);
  $("q").addEventListener("keydown", e => { if (e.key === "Enter") runSearch(); });
  $("qgo").addEventListener("click", runSearch);
  $("resultsclose").addEventListener("click", closeResults);
  $("ocrrun").addEventListener("click", runBannerOcr);
  $("ocrlater").addEventListener("click", () => { docsNeedingOcr().forEach(d => state.ocrDismissed.add(d.id)); updateOcrBanner(); });
  $("ocrall").addEventListener("click", runOcrAll);
  $("cmpall").addEventListener("click", () => compressDocs(state.index.docs.slice(), $("cmplevel").value));
  $("snipprev").addEventListener("click", () => stepNav(-1));
  $("snipnext").addEventListener("click", () => stepNav(1));
  setupZoomInput();
  $("selmovego").addEventListener("click", () => { FolderLib.moveDocs(state.index, [...state.ui.selected], $("selmove").value); state.ui.selected.clear(); saveIndex(); renderOrg(); });
  $("selclear").addEventListener("click", () => { state.ui.selected.clear(); renderOrg(); });
  armConfirm($("seldelete"), "🗑 Elimina", async () => { const ids = [...state.ui.selected]; for (const id of ids) { const d = docById(id); if (d) await removeDoc(d, true); } state.ui.selected.clear(); refreshAll(); });
  try { $("autoopen").checked = localStorage.getItem("snipper.autoopen") !== "0"; } catch (e) { /* ok */ }
  $("autoopen").addEventListener("change", () => { try { localStorage.setItem("snipper.autoopen", $("autoopen").checked ? "1" : "0"); } catch (e) { /* ok */ } });
  setupSelection();
  setupDragDrop();
}

// Pulsante "a doppio clic": il primo clic chiede conferma, il secondo esegue
function armConfirm(btn, label, fn) {
  btn.addEventListener("click", () => {
    if (!btn.classList.contains("armed")) {
      btn.classList.add("armed");
      btn.textContent = "Sicuro? Clicca di nuovo";
      setTimeout(() => { btn.classList.remove("armed"); btn.textContent = label; }, 4000);
    } else {
      btn.classList.remove("armed"); btn.textContent = label;
      fn(btn);
    }
  });
}

function showTab(t) {
  $("view-snip").hidden = t !== "snip";
  $("view-org").hidden = t !== "org";
  $("tab-snip").classList.toggle("active", t === "snip");
  $("tab-org").classList.toggle("active", t === "org");
  if (t === "snip" && state.docId) scheduleRender();
}

function setMode(m) {
  state.mode = m;
  $("mode-snip").classList.toggle("active", m === "snip");
  $("mode-sum").classList.toggle("active", m === "sum");
  $("sumpanel").hidden = m !== "sum";
  setStatus(m === "snip"
    ? "SNIP: trascina un rettangolo sul valore (finisce nella cella selezionata, poi si passa a quella a destra). Tieni premuto CTRL per aggiungere il testo alla cella precedente."
    : "SUM SNIP: seleziona i numeri uno alla volta, poi premi «Scrivi somma in cella».");
}

function refreshAll() {
  renderDocSelect();
  renderOrg();
  updateOcrBanner();
  updateOcrInfo();
}

// =====================================================================
//  Documenti incorporati nel file Excel (parti XML personalizzate)
// =====================================================================
async function readIndex() {
  return Excel.run(async ctx => {
    const col = ctx.workbook.customXmlParts.getByNamespace(NS_INDEX);
    col.load("items/id");
    await ctx.sync();
    let idx = { docs: [], folders: [] };
    if (col.items.length) {
      const xml = col.items[0].getXml();
      await ctx.sync();
      try {
        const d = new DOMParser().parseFromString(xml.value, "application/xml");
        idx = JSON.parse(d.documentElement.textContent);
      } catch (e) { /* indice illeggibile: ripartiamo vuoti */ }
    }
    idx.docs = idx.docs || [];
    idx.folders = idx.folders || [];
    return FolderLib.normalize(idx);
  });
}

async function writeIndex() {
  const xml = `<i xmlns="${NS_INDEX}">${SnipLib.escapeXml(JSON.stringify(state.index))}</i>`;
  return Excel.run(async ctx => {
    const col = ctx.workbook.customXmlParts.getByNamespace(NS_INDEX);
    col.load("items/id");
    await ctx.sync();
    if (col.items.length) {
      col.items[0].setXml(xml);
      for (let i = 1; i < col.items.length; i++) col.items[i].delete();
    } else {
      ctx.workbook.customXmlParts.add(xml);
    }
    await ctx.sync();
  });
}

let saveChain = Promise.resolve();
function saveIndex() {
  saveChain = saveChain.then(() => writeIndex()).catch(e => { console.error(e); setStatus("Errore nel salvataggio: " + (e.message || e), "err"); });
  return saveChain;
}

async function storeBytes(bytes) {
  const chunks = SnipLib.chunkString(SnipLib.bytesToB64(bytes), CHUNK);
  const ids = [];
  try {
    await Excel.run(async ctx => {
      for (const c of chunks) {
        const p = ctx.workbook.customXmlParts.add(`<c xmlns="${NS_DOC}">${c}</c>`);
        p.load("id");
        await ctx.sync();
        ids.push(p.id);
      }
    });
  } catch (e) {
    await deleteParts(ids).catch(() => {});
    throw e;
  }
  return ids;
}

async function loadPartsBytes(partIds) {
  let b64 = "";
  await Excel.run(async ctx => {
    for (const id of partIds) {
      const x = ctx.workbook.customXmlParts.getItem(id).getXml();
      await ctx.sync();
      const m = /<c[^>]*>([\s\S]*)<\/c>/.exec(x.value);
      b64 += m ? m[1] : "";
    }
  });
  return SnipLib.b64ToBytes(b64);
}
const loadBytes = doc => loadPartsBytes(doc.parts);

async function deleteParts(ids) {
  if (!ids || !ids.length) return;
  return Excel.run(async ctx => {
    ids.forEach(id => ctx.workbook.customXmlParts.getItem(id).delete());
    await ctx.sync();
  });
}

async function gzip(bytes) {
  if (typeof CompressionStream === "undefined") return { bytes, gz: false };
  const s = new Blob([bytes]).stream().pipeThrough(new CompressionStream("gzip"));
  return { bytes: new Uint8Array(await new Response(s).arrayBuffer()), gz: true };
}
async function gunzip(bytes) {
  const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("gzip"));
  return new Uint8Array(await new Response(s).arrayBuffer());
}

// =====================================================================
//  Testo delle pagine: nativo (PDF) oppure OCR
// =====================================================================
// pezzi di testo di una pagina PDF, in frazioni (0..1) della pagina vista
function itemsFromContent(tc, vp) {
  const out = [];
  tc.items.forEach(it => {
    if (!it.str || !it.str.trim()) return;
    const t = pdfjsLib.Util.transform(vp.transform, it.transform);
    const w = it.width * vp.scale;
    const sx = Math.hypot(t[0], t[1]) || 1, sy = Math.hypot(t[2], t[3]) || 1;
    const dx = t[0] / sx, dy = t[1] / sx, ux = -t[2] / sy, uy = -t[3] / sy;
    const hh = sy;
    const xs = [t[4], t[4] + dx * w, t[4] + ux * hh, t[4] + dx * w + ux * hh];
    const ys = [t[5], t[5] + dy * w, t[5] + uy * hh, t[5] + dy * w + uy * hh];
    out.push({
      s: it.str,
      u1: Math.min(...xs) / vp.width, v1: Math.min(...ys) / vp.height,
      u2: Math.max(...xs) / vp.width, v2: Math.max(...ys) / vp.height
    });
  });
  return out;
}

async function nativeItems(docId, pdf, n) {
  const c = (state.native[docId] = state.native[docId] || {});
  if (c[n]) return c[n];
  const page = await pdf.getPage(n);
  const vp = page.getViewport({ scale: 1 });
  c[n] = itemsFromContent(await page.getTextContent(), vp);
  return c[n];
}

function effItems(docId, n) {
  const o = state.ocr[docId];
  if (o && o[n] && o[n].length) return o[n];
  const t = state.native[docId];
  return t && t[n] && t[n].length ? t[n] : null;
}

async function loadOcr(doc) {
  if (state.ocr[doc.id]) return state.ocr[doc.id];
  const res = {};
  if (doc.ocr && doc.ocr.parts && doc.ocr.parts.length) {
    let bytes = await loadPartsBytes(doc.ocr.parts);
    if (doc.ocr.gz) bytes = await gunzip(bytes);
    const data = JSON.parse(new TextDecoder().decode(bytes));
    Object.entries(data.pages || {}).forEach(([n, arr]) => {
      res[n] = arr.map(a => ({ s: a[0], u1: a[1], v1: a[2], u2: a[3], v2: a[4] }));
    });
  }
  state.ocr[doc.id] = res;
  return res;
}

async function saveOcr(doc) {
  const data = { v: 1, pages: {} };
  Object.entries(state.ocr[doc.id] || {}).forEach(([n, items]) => {
    data.pages[n] = items.map(i => [i.s, r4(i.u1), r4(i.v1), r4(i.u2), r4(i.v2)]);
  });
  const raw = new TextEncoder().encode(JSON.stringify(data));
  const z = await gzip(raw);
  const old = doc.ocr && doc.ocr.parts ? doc.ocr.parts : [];
  const parts = await storeBytes(z.bytes);
  doc.ocr = { parts, gz: z.gz, pages: Object.keys(data.pages).map(Number).sort((a, b) => a - b), at: new Date().toISOString() };
  await saveIndex();
  await deleteParts(old).catch(() => {});
}

// Pagine senza testo selezionabile (candidate all'OCR)
async function analyzeDoc(doc, pdf) {
  const missing = [];
  for (let n = 1; n <= pdf.numPages; n++) {
    const items = await nativeItems(doc.id, pdf, n);
    const chars = items.reduce((a, i) => a + i.s.replace(/\s/g, "").length, 0);
    if (chars < 5) missing.push(n);
  }
  return missing;
}

const pagesMissing = doc => (doc.noText || []).filter(p => !(doc.ocr && doc.ocr.pages && doc.ocr.pages.includes(p)));
const docsNeedingOcr = () => state.index.docs.filter(d => pagesMissing(d).length > 0 && !state.ocrDismissed.has(d.id));

function updateOcrBanner() {
  const list = docsNeedingOcr();
  const b = $("ocrbanner");
  if (!list.length) { b.hidden = true; return; }
  b.hidden = false;
  if (list.length === 1) {
    const d = list[0], n = pagesMissing(d).length;
    $("ocrtext").textContent = `⚠ «${d.name}»: ${n === d.pages ? "nessuna pagina ha testo leggibile" : n + " pagine senza testo leggibile"} (probabile scansione). Senza OCR non posso leggere i valori né cercare nel testo.`;
  } else {
    $("ocrtext").textContent = `⚠ ${list.length} documenti sembrano scansioni (pagine senza testo leggibile). Senza OCR non posso leggere i valori né cercare nel testo.`;
  }
}

async function runBannerOcr() {
  const list = docsNeedingOcr();
  for (const d of list) await runOcr(d, "missing", 1);
}

function updateOcrInfo() {
  const d = curDoc();
  return d;
}

function pagesForScope(doc, scope, cur) {
  const all = Array.from({ length: doc.pages }, (_, i) => i + 1);
  if (scope === "current") return [cur];
  if (scope === "all") return all;
  const miss = pagesMissing(doc);
  return miss.length ? miss : [];
}

async function blobToImage(bytes, mime) {
  return new Promise((res, rej) => {
    const img = new Image();
    img.onload = () => res(img);
    img.onerror = () => rej(new Error("immagine non valida"));
    img.src = URL.createObjectURL(new Blob([bytes], { type: mime }));
  });
}

async function runOcr(doc, scope, curPage) {
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  state.busy = true;
  let worker = null, pdf = null;
  try {
    await loadOcr(doc);
    const pages = pagesForScope(doc, scope, curPage);
    if (!pages.length) { setStatus("Nessuna pagina senza testo: se vuoi comunque l'OCR scegli «tutto il documento» o «pagina corrente».", "ok"); return; }
    const bytes = await loadBytes(doc);
    let img = null;
    if (doc.type === "pdf") pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    else img = await blobToImage(bytes, doc.mime);
    setStatus("Preparo l'OCR… (la prima volta scarica i dati della lingua)");
    worker = await Tesseract.createWorker(["ita", "eng"]);
    for (let i = 0; i < pages.length; i++) {
      const n = pages[i];
      setStatus(`OCR di «${doc.name}»: pagina ${n} (${i + 1} di ${pages.length})…`);
      const cv = document.createElement("canvas");
      const ctx = cv.getContext("2d");
      if (pdf) {
        const page = await pdf.getPage(n);
        const vp = page.getViewport({ scale: OCR_SCALE });
        cv.width = Math.floor(vp.width); cv.height = Math.floor(vp.height);
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);
        await page.render({ canvasContext: ctx, viewport: vp }).promise;
      } else {
        cv.width = img.naturalWidth; cv.height = img.naturalHeight;
        ctx.fillStyle = "#fff"; ctx.fillRect(0, 0, cv.width, cv.height);
        ctx.drawImage(img, 0, 0);
      }
      const { data } = await worker.recognize(cv);
      const W = cv.width, H = cv.height;
      state.ocr[doc.id][n] = (data.words || [])
        .filter(w => w.text && w.text.trim() && w.bbox)
        .map(w => ({ s: w.text.trim(), u1: w.bbox.x0 / W, v1: w.bbox.y0 / H, u2: w.bbox.x1 / W, v2: w.bbox.y1 / H }));
    }
    setStatus("Salvo il testo letto nel file Excel…");
    await saveOcr(doc);
    state.ocrDismissed.delete(doc.id);
    refreshAll();
    if (state.docId === doc.id) await renderPage();
    setStatus(`OCR completato su ${pages.length} pagina/e di «${doc.name}». Ora puoi spuntare e cercare nel testo.`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("OCR non riuscito: " + (e.message || e), "err");
  } finally {
    if (worker) { try { await worker.terminate(); } catch (e) { /* ok */ } }
    if (pdf) { try { pdf.destroy(); } catch (e) { /* ok */ } }
    state.busy = false;
  }
}

// =====================================================================
//  Importazione (selezione manuale, cartelle, trascinamento)
// =====================================================================
async function importItems(items) {
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  const ok = items.filter(i => EXT_OK.test(i.file.name));
  const skipped = items.length - ok.length;
  if (!ok.length) { setStatus(skipped ? "Formati accettati: PDF, PNG, JPG." : "Nessun file.", "err"); return; }
  state.busy = true;
  const added = [];
  let dup = 0, last = null;
  try {
    for (let i = 0; i < ok.length; i++) {
      const { file: f, folder } = ok[i];
      setStatus(`Importo ${f.name} (${i + 1}/${ok.length})… non chiudere Excel.`);
      const same = state.index.docs.find(d => d.name === f.name && d.size === f.size);
      if (same) { dup++; last = same.id; continue; }
      const isPdf = /\.pdf$/i.test(f.name) || f.type === "application/pdf";
      const isPng = /\.png$/i.test(f.name) || f.type === "image/png";
      const bytes = new Uint8Array(await f.arrayBuffer());
      const doc = {
        id: uid(), name: f.name, type: isPdf ? "pdf" : "image",
        mime: isPdf ? "application/pdf" : (isPng ? "image/png" : "image/jpeg"),
        size: f.size, pages: 1, folder: FolderLib.cleanPath(folder), parts: [], added: new Date().toISOString(),
        noText: [], ocr: null
      };
      if (isPdf) {
        const test = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;   // controlla che sia valido
        doc.pages = test.numPages;
        doc.noText = await analyzeDoc(doc, test);
        test.destroy();
      } else {
        doc.noText = [1];                                                           // un'immagine non ha testo
      }
      doc.parts = await storeBytes(bytes);
      state.index.docs.push(doc);
      FolderLib.normalize(state.index);
      await writeIndex();
      added.push(doc);
      last = doc.id;
    }
    refreshAll();
    const big = ok.some(i => i.file.size > WARN_MB * 1048576);
    if (added.length === 1 && state.ui.currentFolder === added[0].folder) { /* resta dov'è */ }
    if (last && (added.length === 1 || (!added.length && dup === 1))) { showTab("snip"); await openDoc(last); }
    const scans = added.filter(d => pagesMissing(d).length).length;
    setStatus(`Importati ${added.length} documento/i` + (dup ? `, ${dup} già presente/i` : "") + (skipped ? `, ${skipped} ignorato/i (formato non valido)` : "") +
      (scans ? `. ${scans} sembra/no scansione/i: ti propongo l'OCR.` : ".") + " Ricorda di salvare il file Excel." +
      (big ? " Attenzione: file grandi appesantiscono molto la cartella di lavoro." : ""), big || scans ? "err" : "ok");
  } catch (e) {
    console.error(e);
    setStatus("Importazione non riuscita: " + (e.message || e), "err");
    refreshAll();
  } finally {
    state.busy = false;
  }
}

// ---- trascinamento di file e cartelle dentro il pannello ----
function setupDragDrop() {
  const hasFiles = e => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
  let depth = 0;
  window.addEventListener("dragenter", e => { if (!hasFiles(e)) return; e.preventDefault(); depth++; $("dropover").hidden = false; });
  window.addEventListener("dragover", e => { if (!hasFiles(e)) return; e.preventDefault(); e.dataTransfer.dropEffect = "copy"; });
  window.addEventListener("dragleave", e => { if (!hasFiles(e)) return; depth = Math.max(0, depth - 1); if (!depth) $("dropover").hidden = true; });
  window.addEventListener("drop", async e => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    depth = 0; $("dropover").hidden = true;
    // gli elementi vanno letti subito: dopo il gestore dell'evento non sono più disponibili
    const dt = e.dataTransfer;
    const entries = Array.from(dt.items || []).filter(i => i.kind === "file").map(i => (i.webkitGetAsEntry ? i.webkitGetAsEntry() : null));
    const plain = Array.from(dt.files || []);
    try {
      const out = [];
      if (entries.length && entries.every(Boolean)) {
        for (const en of entries) await walkEntry(en, state.ui.currentFolder, out, true);
      } else {
        plain.forEach(file => out.push({ file, folder: state.ui.currentFolder }));
      }
      await importItems(out);
    } catch (err) {
      console.error(err);
      setStatus("Impossibile leggere gli elementi trascinati: " + (err.message || err), "err");
    }
  });
}

async function walkEntry(entry, folder, out, top) {
  if (entry.isFile) {
    const file = await new Promise((res, rej) => entry.file(res, rej));
    out.push({ file, folder });
  } else if (entry.isDirectory) {
    const sub = FolderLib.join(folder, entry.name);
    const reader = entry.createReader();
    let batch;
    do {
      batch = await new Promise((res, rej) => reader.readEntries(res, rej));
      for (const child of batch) await walkEntry(child, sub, out, false);
    } while (batch.length);
  }
}

// =====================================================================
//  Schermata ORGANIZZA (cartelle, sottocartelle, spostamenti)
// =====================================================================
function renderDocSelect() {
  const sel = $("docselect");
  sel.innerHTML = "";
  if (!state.index.docs.length) {
    sel.appendChild(new Option("(nessun documento)", ""));
    return;
  }
  // elenco ordinato per cartella e nome
  [...state.index.docs].sort((a, b) => (a.folder + "/" + a.name).localeCompare(b.folder + "/" + b.name, "it"))
    .forEach(d => sel.appendChild(new Option((d.folder ? d.folder + " / " : "") + d.name, d.id)));
  if (state.docId) sel.value = state.docId;
}

function snipCount(docId) {
  return new Set(state.log.filter(r => r.docId === docId).map(r => r.id)).size;
}

function renderOrg() {
  const q = ($("docfilter").value || "").toLowerCase().trim();
  const total = state.index.docs.reduce((a, d) => a + d.size, 0);
  $("orgsummary").textContent = state.index.docs.length
    ? `${state.index.docs.length} documenti · ${state.index.folders.length} cartelle · ${fmtSize(total)} incorporati nel file`
    : "Nessun documento importato.";
  $("target").textContent = state.ui.currentFolder ? "📁 " + state.ui.currentFolder : "📂 cartella principale";

  // barra della selezione
  const sel = state.ui.selected;
  [...sel].forEach(id => { if (!docById(id)) sel.delete(id); });
  $("selbar").hidden = !sel.size;
  if (sel.size) {
    $("selcount").textContent = `${sel.size} selezionati`;
    const m = $("selmove");
    m.innerHTML = "";
    FolderLib.options(state.index).forEach(o => m.appendChild(new Option(o.label, o.path)));
  }

  const box = $("tree");
  box.innerHTML = "";
  if (q) {
    const found = state.index.docs.filter(d => d.name.toLowerCase().includes(q));
    if (!found.length) box.append(h("div", { class: "muted pad", text: "Nessun documento trovato." }));
    found.sort((a, b) => a.name.localeCompare(b.name, "it")).forEach(d => box.append(docCard(d, true)));
    return;
  }
  const tree = FolderLib.buildTree(state.index);
  box.append(folderRow(tree, 0, true));
  if (state.ui.creatingIn === "") box.append(newFolderInput("", 1));
  if (state.ui.expanded.has("") || !state.ui.expanded.size) renderChildren(tree, 1, box);
}

function renderChildren(node, depth, box) {
  node.children.forEach(c => {
    box.append(folderRow(c, depth, false));
    if (state.ui.creatingIn === c.path) box.append(newFolderInput(c.path, depth + 1));
    if (!state.ui.collapsed || !state.ui.collapsed.has(c.path)) renderChildren(c, depth + 1, box);
  });
  node.docs.forEach(d => { const w = h("div", { style: `margin-left:${(depth) * 14}px` }, docCard(d, false)); box.append(w); });
}

function folderRow(node, depth, isRoot) {
  state.ui.collapsed = state.ui.collapsed || new Set();
  const collapsed = state.ui.collapsed.has(node.path);
  const hasKids = node.children.length || node.docs.length;
  const row = h("div", { class: "frow" + (state.ui.currentFolder === node.path ? " cur" : ""), style: `margin-left:${depth * 14}px` });
  row.append(h("span", { class: "twisty", text: isRoot ? "" : (hasKids ? (collapsed ? "▸" : "▾") : ""), onclick: ev => { ev.stopPropagation(); if (collapsed) state.ui.collapsed.delete(node.path); else state.ui.collapsed.add(node.path); renderOrg(); } }));
  if (state.ui.renaming === node.path && !isRoot) {
    const inp = h("input", { value: node.name });
    const done = ok => {
      const v = inp.value;
      state.ui.renaming = null;
      if (ok && v.trim() && v.trim() !== node.name) {
        const np = FolderLib.renameFolder(state.index, node.path, v);
        if (np === null) setStatus("Esiste già una cartella con quel nome.", "err");
        else { if (state.ui.currentFolder === node.path || FolderLib.isInside(state.ui.currentFolder, node.path)) state.ui.currentFolder = np + state.ui.currentFolder.slice(node.path.length); saveIndex(); refreshAll(); return; }
      }
      renderOrg();
    };
    inp.addEventListener("keydown", e => { if (e.key === "Enter") done(true); else if (e.key === "Escape") done(false); });
    inp.addEventListener("blur", () => done(true));
    inp.addEventListener("click", e => e.stopPropagation());
    row.append(h("span", { text: "📁" }), inp);
    setTimeout(() => { inp.focus(); inp.select(); }, 0);
  } else {
    row.append(h("span", { text: isRoot ? "📂" : "📁" }),
      h("span", { class: "fname", text: isRoot ? "Tutti i documenti" : node.name }),
      h("span", { class: "fcount", text: String(FolderLib.countDocs(node)) }));
  }
  const icons = h("span", { class: "ficons" });
  icons.append(h("button", { title: "Nuova sottocartella", text: "＋", onclick: ev => { ev.stopPropagation(); state.ui.creatingIn = node.path; state.ui.collapsed.delete(node.path); renderOrg(); } }));
  if (!isRoot) {
    icons.append(h("button", { title: "Rinomina", text: "✎", onclick: ev => { ev.stopPropagation(); state.ui.renaming = node.path; renderOrg(); } }));
    const del = h("button", { title: "Elimina cartella (il contenuto passa alla cartella superiore)", text: "🗑" });
    armConfirm(del, "🗑", () => {
      FolderLib.deleteFolder(state.index, node.path);
      if (FolderLib.isInside(state.ui.currentFolder, node.path)) state.ui.currentFolder = FolderLib.parentOf(node.path);
      saveIndex(); refreshAll();
      setStatus(`Cartella «${node.name}» eliminata: i documenti sono passati alla cartella superiore.`, "ok");
    });
    del.addEventListener("click", ev => ev.stopPropagation());
    icons.append(del);
  }
  row.append(icons);
  row.addEventListener("click", () => { state.ui.currentFolder = node.path; renderOrg(); });

  // trascinare documenti o cartelle su una cartella per spostarli
  row.addEventListener("dragover", e => {
    const t = Array.from(e.dataTransfer.types || []);
    if (t.includes("text/x-snipper-doc") || t.includes("text/x-snipper-folder")) { e.preventDefault(); row.classList.add("dragover"); }
  });
  row.addEventListener("dragleave", () => row.classList.remove("dragover"));
  row.addEventListener("drop", e => {
    row.classList.remove("dragover");
    const ids = e.dataTransfer.getData("text/x-snipper-doc");
    const fol = e.dataTransfer.getData("text/x-snipper-folder");
    if (ids) { e.preventDefault(); FolderLib.moveDocs(state.index, ids.split(","), node.path); state.ui.selected.clear(); saveIndex(); refreshAll(); }
    else if (fol) {
      e.preventDefault();
      if (FolderLib.moveFolder(state.index, fol, node.path) === null) setStatus("Non posso spostare la cartella lì.", "err");
      else { saveIndex(); refreshAll(); }
    }
  });
  if (!isRoot) {
    row.setAttribute("draggable", "true");
    row.addEventListener("dragstart", e => { e.dataTransfer.setData("text/x-snipper-folder", node.path); e.dataTransfer.effectAllowed = "move"; });
  }
  return row;
}

function newFolderInput(parent, depth) {
  const inp = h("input", { placeholder: "Nome della nuova cartella" });
  const row = h("div", { class: "frow", style: `margin-left:${depth * 14}px` }, h("span", { class: "twisty" }), h("span", { text: "📁" }), inp);
  const done = ok => {
    const v = inp.value.trim();
    state.ui.creatingIn = null;
    if (ok && v) { const p = FolderLib.addFolder(state.index, parent, v); state.ui.currentFolder = p; saveIndex(); refreshAll(); }
    else renderOrg();
  };
  inp.addEventListener("keydown", e => { if (e.key === "Enter") done(true); else if (e.key === "Escape") done(false); });
  inp.addEventListener("blur", () => done(true));
  setTimeout(() => inp.focus(), 0);
  return row;
}

function docCard(d, showPath) {
  const n = snipCount(d.id);
  const miss = pagesMissing(d).length;
  const ocrN = d.ocr && d.ocr.pages ? d.ocr.pages.length : 0;
  const card = h("div", { class: "doc" + (state.ui.selected.has(d.id) ? " sel" : ""), draggable: "true" });
  card.addEventListener("dragstart", e => {
    const ids = state.ui.selected.has(d.id) ? [...state.ui.selected] : [d.id];
    e.dataTransfer.setData("text/x-snipper-doc", ids.join(","));
    e.dataTransfer.effectAllowed = "move";
  });
  const cb = h("input", { type: "checkbox" });
  cb.checked = state.ui.selected.has(d.id);
  cb.addEventListener("change", () => { if (cb.checked) state.ui.selected.add(d.id); else state.ui.selected.delete(d.id); renderOrg(); });
  const meta = h("div", { class: "meta" }, `${d.pages} pag. · ${fmtSize(d.size)} · ${n} snip `);
  if (miss) meta.append(h("span", { class: "badge warn", text: `${miss} pag. senza testo` }), " ");
  if (ocrN) meta.append(h("span", { class: "badge", text: `OCR ${ocrN} pag.` }));
  card.append(h("div", { class: "top" }, cb, h("div", {}, h("div", { class: "docname", text: (d.type === "pdf" ? "📄 " : "🖼️ ") + d.name }), meta,
    showPath && d.folder ? h("div", { class: "meta", text: "📁 " + d.folder }) : null)));

  const move = h("select", { title: "Sposta in…" });
  FolderLib.options(state.index).forEach(o => move.appendChild(new Option(o.label, o.path)));
  move.value = d.folder;
  move.addEventListener("change", () => { FolderLib.moveDocs(state.index, [d.id], move.value); saveIndex(); renderOrg(); renderDocSelect(); });
  card.append(h("div", { class: "moveto" }, "Sposta in:", move));

  const act = h("div", { class: "actions" });
  const mk = (txt, fn, cls) => { const b = h("button", { text: txt, class: cls || null }); b.addEventListener("click", () => fn(b)); act.append(b); return b; };
  mk("Apri", () => { showTab("snip"); openDoc(d.id); });
  mk("⬇ Originale", () => downloadOriginal(d));
  mk(`⬇ Con spunte${n ? " (" + n + ")" : ""}`, () => downloadAnnotated(d));
  mk("🗜 Comprimi", () => compressDocs([d], $("cmplevel").value));
  mk("🔎 OCR", () => runOcr(d, miss ? "missing" : "all", 1));
  const del = h("button", { text: "🗑 Elimina", class: "danger" });
  armConfirm(del, "🗑 Elimina", () => removeDoc(d));
  act.append(del);
  card.append(act);
  return card;
}

async function removeDoc(d, silent) {
  if (state.busy && !silent) return;
  const prev = state.busy;
  state.busy = true;
  try {
    if (!silent) setStatus("Elimino…");
    await deleteParts(d.parts);
    if (d.ocr && d.ocr.parts) await deleteParts(d.ocr.parts).catch(() => {});
    state.index.docs = state.index.docs.filter(x => x.id !== d.id);
    delete state.native[d.id]; delete state.ocr[d.id];
    await writeIndex();
    const had = state.log.some(r => r.docId === d.id);
    if (had) {
      state.log = state.log.filter(r => r.docId !== d.id);
      await rewriteLog();
    }
    if (state.docId === d.id) closeDoc();
    if (!silent) { refreshAll(); setStatus(`«${d.name}» eliminato.` + (had ? " Le celle già spuntate restano in Excel ma non sono più collegate al documento." : ""), "ok"); }
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = prev;
  }
}

// =====================================================================
//  Download
// =====================================================================
function saveBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 5000);
}

async function downloadOriginal(d) {
  try {
    setStatus("Preparo il file…");
    const bytes = await loadBytes(d);
    saveBlob(new Blob([bytes], { type: d.mime }), d.name);
    setStatus(`Scaricato «${d.name}».`, "ok");
  } catch (e) { setStatus("Errore: " + (e.message || e), "err"); }
}

async function downloadAnnotated(d) {
  try {
    setStatus("Disegno le spunte sul documento…");
    const bytes = await loadBytes(d);
    const rows = state.log.filter(r => r.docId === d.id);
    const out = await ExportLib.annotate(PDFLib, bytes, d, rows);
    saveBlob(new Blob([out], { type: "application/pdf" }), d.name.replace(/\.[^.]+$/, "") + "_con_spunte.pdf");
    setStatus(`Scaricato «${d.name.replace(/\.[^.]+$/, "")}_con_spunte.pdf» (${new Set(rows.map(r => r.id)).size} snip).`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  }
}

// =====================================================================
//  Visualizzazione del documento
// =====================================================================
function closeDoc() {
  state.docId = null; state.pdf = null; state.image = null; state.pageItems = [];
  canvas.width = canvas.height = 0;
  canvas.style.width = canvas.style.height = "0px";
  $("stage").style.display = "none";
  $("empty").style.display = "";
  $("pageinfo").textContent = "–";
  updateOcrInfo();
  drawMarks();
}

async function openDoc(id, opts) {
  opts = opts || {};
  const d = docById(id);
  if (!d) return;
  try {
    setStatus(`Apro «${d.name}»…`);
    const bytes = await loadBytes(d);
    if (state.pdf) { try { state.pdf.destroy(); } catch (e) { /* ok */ } }
    state.pdf = null; state.image = null;
    if (d.type === "pdf") {
      state.pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
      state.numPages = state.pdf.numPages;
      if (d.noText === undefined) {                       // file creato con una versione precedente
        d.noText = await analyzeDoc(d, state.pdf);
        await saveIndex();
        refreshAll();
      }
    } else {
      state.image = await blobToImage(bytes, d.mime);
      state.numPages = 1;
      if (d.noText === undefined) { d.noText = [1]; await saveIndex(); }
    }
    await loadOcr(d);
    state.docId = id;
    state.page = Math.min(Math.max(1, opts.page || 1), state.numPages);
    state.sumItems = []; renderSum();
    state.zoom = 1;
    $("docselect").value = id;
    $("stage").style.display = "";
    $("empty").style.display = "none";
    state.pendingScroll = true;
    await renderPage();
    updateOcrBanner();
    updateOcrInfo();
    setStatus(opts.message || "Documento aperto. Trascina un rettangolo sul valore da spuntare.", "ok");
  } catch (e) {
    console.error(e);
    setStatus("Impossibile aprire il documento: " + (e.message || e), "err");
  }
}

async function renderPage() {
  const dpr = state.dpr = window.devicePixelRatio || 1;
  const ctx = canvas.getContext("2d");
  const d = curDoc();
  state.pageItems = [];

  if (state.pdf) {
    const page = await state.pdf.getPage(state.page);
    const v1 = page.getViewport({ scale: 1 });
    state.scale = fitWidth() / v1.width * state.zoom;
    const viewport = page.getViewport({ scale: state.scale });
    state.viewW = viewport.width; state.viewH = viewport.height;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = viewport.width + "px";
    canvas.style.height = viewport.height + "px";
    await page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: state.scale * dpr }) }).promise;
    await nativeItems(d.id, state.pdf, state.page);
  } else if (state.image) {
    const base = fitWidth() * state.zoom;
    const w = Math.round(base);
    const hh = Math.round(base * state.image.naturalHeight / state.image.naturalWidth);
    state.viewW = w; state.viewH = hh;
    canvas.width = Math.floor(w * dpr);
    canvas.height = Math.floor(hh * dpr);
    canvas.style.width = w + "px";
    canvas.style.height = hh + "px";
    ctx.drawImage(state.image, 0, 0, canvas.width, canvas.height);
  }
  if (d) {
    const it = effItems(d.id, state.page);
    state.pageItems = it ? it.map(i => ({ str: i.s, x: i.u1 * state.viewW, y: i.v2 * state.viewH, w: (i.u2 - i.u1) * state.viewW, h: (i.v2 - i.v1) * state.viewH })) : [];
  }
  $("pageinfo").textContent = state.docId ? `${state.page} / ${state.numPages}` : "–";
  drawMarks();
  if (state.pendingScroll) { state.pendingScroll = false; scrollToMark(); }
}

async function gotoPage(n) {
  if (!state.docId || n < 1 || n > state.numPages) return;
  state.page = n;
  await renderPage();
}

const fitWidth = () => Math.max(120, $("viewer").clientWidth - 16);

let rendering = false, again = false;
async function scheduleRender() {
  if (!state.docId) return;
  if (rendering) { again = true; return; }
  rendering = true;
  try { await renderPage(); } catch (e) { console.warn(e); }
  rendering = false;
  if (again) { again = false; scheduleRender(); }
}

function zoom(f) {
  state.zoom = Math.min(4, Math.max(0.4, state.zoom * f));
  scheduleRender();
}

// zoom con CTRL+rotellina (o CTRL +/−) solo sul documento, non sul pannello
function setupZoomInput() {
  window.addEventListener("wheel", e => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    if ($("viewer").contains(e.target)) zoom(e.deltaY < 0 ? 1.1 : 1 / 1.1);
  }, { passive: false });
  window.addEventListener("keydown", e => {
    if (!e.ctrlKey) return;
    if (e.key === "+" || e.key === "=") { e.preventDefault(); zoom(1.2); }
    else if (e.key === "-") { e.preventDefault(); zoom(1 / 1.2); }
    else if (e.key === "0") { e.preventDefault(); state.zoom = 1; scheduleRender(); }
  });
  let t = null;
  window.addEventListener("resize", () => { clearTimeout(t); t = setTimeout(scheduleRender, 200); });
}

function drawMarks() {
  overlay.querySelectorAll(".mark").forEach(n => n.remove());
  if (!state.docId) return;
  const W = state.viewW, H = state.viewH;
  const add = (r, cls, label) => {
    const el = h("div", { class: "mark " + cls });
    Object.assign(el.style, {
      left: Math.min(r.u1, r.u2) * W + "px", top: Math.min(r.v1, r.v2) * H + "px",
      width: Math.abs(r.u2 - r.u1) * W + "px", height: Math.abs(r.v2 - r.v1) * H + "px"
    });
    if (label) el.append(h("b", { text: label }));
    overlay.append(el);
  };
  const focusing = state.focusId != null;
  state.log.filter(r => r.docId === state.docId && r.page === state.page).forEach(r => {
    const isF = r.id === state.focusId;
    add(r, (r.type === "SUM SNIP" ? "sum " : "") + (isF ? "focus" : (focusing ? "dim" : "")), "#" + r.id);
  });
  state.sumItems.filter(s => s.page === state.page).forEach(s => add(s, "pending", "+"));
  state.searchMarks.filter(m => m.docId === state.docId && m.page === state.page).forEach(m => add(m, "search", ""));
}

function scrollToMark() {
  const m = overlay.querySelector(".mark.focus") || overlay.querySelector(".mark.search");
  const v = $("viewer");
  if (!m || !v) return;
  const top = m.offsetTop - v.clientHeight / 2 + m.offsetHeight / 2;
  const left = m.offsetLeft - v.clientWidth / 2 + m.offsetWidth / 2;
  v.scrollTo({ top: Math.max(0, top), left: Math.max(0, left), behavior: "smooth" });
}

// =====================================================================
//  Selezione con il mouse
// =====================================================================
function setupSelection() {
  let start = null, box = null;
  const pos = e => {
    const r = overlay.getBoundingClientRect();
    return { x: Math.min(Math.max(e.clientX - r.left, 0), r.width), y: Math.min(Math.max(e.clientY - r.top, 0), r.height), w: r.width, h: r.height };
  };
  overlay.addEventListener("mousedown", e => {
    if (!state.docId) return;
    start = pos(e);
    box = h("div", { class: "sel" });
    overlay.append(box);
    e.preventDefault();
  });
  window.addEventListener("mousemove", e => {
    if (!start) return;
    const p = pos(e);
    Object.assign(box.style, {
      left: Math.min(start.x, p.x) + "px", top: Math.min(start.y, p.y) + "px",
      width: Math.abs(p.x - start.x) + "px", height: Math.abs(p.y - start.y) + "px"
    });
  });
  window.addEventListener("mouseup", async e => {
    if (!start) return;
    const p = pos(e);
    const withCtrl = e.ctrlKey;
    const f = { u1: Math.min(start.x, p.x) / p.w, v1: Math.min(start.y, p.y) / p.h, u2: Math.max(start.x, p.x) / p.w, v2: Math.max(start.y, p.y) / p.h };
    const tiny = (f.u2 - f.u1) * p.w < 4 || (f.v2 - f.v1) * p.h < 4;
    start = null;
    if (box) { box.remove(); box = null; }
    if (tiny) return;
    await onSnip(f, withCtrl);
  });
}

// =====================================================================
//  Lettura dell'area: testo (PDF o OCR salvato), altrimenti OCR al volo
// =====================================================================
async function readArea(f) {
  if (state.pageItems.length) {
    const t = SnipLib.textInRect(state.pageItems, { x1: f.u1 * state.viewW, y1: f.v1 * state.viewH, x2: f.u2 * state.viewW, y2: f.v2 * state.viewH });
    if (t) {
      const o = state.ocr[state.docId];
      return { text: t, source: o && o[state.page] && o[state.page].length ? "OCR salvato" : "testo PDF" };
    }
  }
  return { text: await ocrArea(f), source: "OCR al volo" };
}

async function ocrArea(f) {
  setStatus("Pagina senza testo: leggo l'area con l'OCR… (suggerimento: usa «Avvia OCR» per tutta la pagina o il documento)");
  const sx = f.u1 * canvas.width, sy = f.v1 * canvas.height;
  const sw = (f.u2 - f.u1) * canvas.width, sh = (f.v2 - f.v1) * canvas.height;
  const up = Math.max(1, Math.min(3, 300 / Math.max(sh, 1)));
  const tmp = document.createElement("canvas");
  tmp.width = Math.ceil(sw * up) + 20;
  tmp.height = Math.ceil(sh * up) + 20;
  const c = tmp.getContext("2d");
  c.fillStyle = "#fff"; c.fillRect(0, 0, tmp.width, tmp.height);
  c.drawImage(canvas, sx, sy, sw, sh, 10, 10, sw * up, sh * up);
  const { data } = await Tesseract.recognize(tmp, "ita+eng");
  return (data.text || "").replace(/\s+/g, " ").trim();
}

// =====================================================================
//  SNIP / SUM SNIP
// =====================================================================
async function onSnip(f, withCtrl) {
  if (state.busy) return;
  state.busy = true;
  try {
    const { text, source } = await readArea(f);
    if (!text) { setStatus("Nessun testo trovato nell'area selezionata.", "err"); return; }
    const fmt = $("numfmt").value;

    if (state.mode === "sum") {
      const v = SnipLib.parseNumber(text, fmt);
      if (v === null) { setStatus(`«${text}» non è un numero valido (formato: ${fmt === "it" ? "italiano" : "inglese"}).`, "err"); return; }
      state.sumItems.push({ text, value: v, page: state.page, ...f });
      drawMarks(); renderSum();
      setStatus(`Aggiunto ${text} (${source}).`, "ok");
      return;
    }

    let value = text, isNum = false;
    if ($("asnumber").checked) {
      const n = SnipLib.parseNumber(text, fmt);
      if (n !== null) { value = n; isNum = true; }
    }
    let addr, appended = false;
    if (withCtrl) {
      const sel = await currentCellAddress();
      const target = state.lastCell && state.lastNext === sel ? state.lastCell : sel;
      const r = await appendToCell(target, text, "#C6EFCE");
      addr = target; appended = r;
    } else {
      addr = await writeToCell(value, "#C6EFCE");
    }
    const id = nextSnipId();
    await addLog([{ id, cell: addr, type: "SNIP", page: state.page, ...f, text, value }]);
    state.focusId = null; state.searchMarks = [];
    drawMarks();
    if (withCtrl) setStatus(`#${id}: «${text}» aggiunto alla cella ${addr}${appended ? " (separato da « - »)" : ""} (${source}).`, "ok");
    else {
      state.lastCell = addr;
      state.lastNext = await moveRight();
      setStatus(`#${id}: scritto ${isNum ? value : "«" + text + "»"} in ${addr} (${source}). Prossimo snip in ${state.lastNext}.`, "ok");
    }
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

function renderSum() {
  const ul = $("sumlist");
  ul.innerHTML = "";
  state.sumItems.forEach((it, i) => {
    const li = h("li", { text: `${it.text}  →  ${it.value}` });
    li.append(h("button", { text: "✕", title: "Rimuovi", onclick: () => { state.sumItems.splice(i, 1); renderSum(); drawMarks(); } }));
    ul.append(li);
  });
  const total = state.sumItems.reduce((a, b) => a + b.value, 0);
  $("sumtotal").textContent = "Totale: " + (Math.round(total * 1e6) / 1e6).toLocaleString("it-IT", { maximumFractionDigits: 6 });
}

async function writeSum() {
  if (!state.sumItems.length) { setStatus("Nessun numero da sommare.", "err"); return; }
  if (state.busy) return;
  state.busy = true;
  try {
    const values = state.sumItems.map(i => i.value);
    const formula = SnipLib.sumFormula(values);
    const addr = await writeToCell(formula, "#BDD7EE", true);
    const id = nextSnipId();
    await addLog(state.sumItems.map(i => ({
      id, cell: addr, type: "SUM SNIP", page: i.page, u1: i.u1, v1: i.v1, u2: i.u2, v2: i.v2, text: i.text, value: i.value
    })));
    state.sumItems = [];
    renderSum(); drawMarks();
    state.lastCell = addr;
    state.lastNext = await moveRight();
    setStatus(`#${id}: somma scritta in ${addr}: ${formula}. Prossimo in ${state.lastNext}.`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

// =====================================================================
//  Clic su una cella con uno snip: apre il pannello, il documento e lo evidenzia
// =====================================================================
async function registerSelectionHandler() {
  try {
    await Excel.run(async ctx => {
      ctx.workbook.onSelectionChanged.add(onSelectionChanged);
      await ctx.sync();
    });
  } catch (e) { console.warn("Evento di selezione non disponibile", e); }
}

let selTimer = null;
function onSelectionChanged() {
  clearTimeout(selTimer);
  selTimer = setTimeout(() => { handleSelection().catch(e => console.warn(e)); }, 200);
}

async function currentCellAddress() {
  return Excel.run(async ctx => {
    const c = ctx.workbook.getSelectedRange().getCell(0, 0);
    c.load("address");
    await ctx.sync();
    return c.address;
  });
}

async function handleSelection() {
  if (state.busy || !$("autoopen").checked || !state.log.length) return;
  const addr = await currentCellAddress();
  const ids = [...new Set(state.log.filter(r => r.cell === addr).map(r => r.id))];
  if (!ids.length) {
    state.nav = null; updateNav();
    if (state.focusId != null) { state.focusId = null; drawMarks(); }
    return;
  }
  if (state.nav && state.nav.addr === addr && state.focusId != null) return;   // già mostrato
  try { if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane(); } catch (e) { /* già aperto */ }
  await focusCell(addr, ids, 0);
}

async function focusCell(addr, ids, idx) {
  state.nav = { addr, ids, idx };
  updateNav();
  const row = state.log.find(r => r.id === ids[idx]);
  if (row) await focusSnip(row, addr, ids.length > 1 ? ` (${idx + 1}/${ids.length})` : "");
}

function updateNav() {
  const n = state.nav && state.nav.ids.length > 1 ? state.nav : null;
  ["snipprev", "snipnav", "snipnext"].forEach(i => { $(i).hidden = !n; });
  if (n) $("snipnav").textContent = `${n.idx + 1}/${n.ids.length}`;
}

async function stepNav(d) {
  const n = state.nav;
  if (!n) return;
  const idx = (n.idx + d + n.ids.length) % n.ids.length;
  await focusCell(n.addr, n.ids, idx);
}

async function focusSnip(row, addr, extra) {
  if (!docById(row.docId)) { setStatus("Il documento di questo snip non è più nel file.", "err"); return; }
  showTab("snip");
  state.focusId = row.id;
  state.searchMarks = [];
  const msg = `Snip #${row.id}${extra || ""} · ${addr || row.cell} · «${row.text}» · pagina ${row.page}`;
  if (state.docId !== row.docId) {
    await openDoc(row.docId, { page: row.page, message: msg });
  } else {
    state.page = row.page;
    state.pendingScroll = true;
    await renderPage();
    setStatus(msg, "ok");
  }
}

async function findSnipFromCell() {
  try {
    const addr = await currentCellAddress();
    const ids = [...new Set(state.log.filter(r => r.cell === addr).map(r => r.id))];
    if (!ids.length) { setStatus(`Nessuno snip sulla cella ${addr}.`, "err"); return; }
    await focusCell(addr, ids, 0);
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  }
}

// =====================================================================
//  Ricerca di parole in tutti i documenti
// =====================================================================
async function ensureDocText(doc, label) {
  await loadOcr(doc);
  if (doc.type !== "pdf") return;
  const cached = state.native[doc.id];
  if (cached && Object.keys(cached).length >= doc.pages) return;
  let pdf = state.docId === doc.id ? state.pdf : null, own = false;
  if (!pdf) {
    setStatus(label);
    pdf = await pdfjsLib.getDocument({ data: (await loadBytes(doc)).slice() }).promise;
    own = true;
  }
  try { for (let n = 1; n <= doc.pages; n++) await nativeItems(doc.id, pdf, n); }
  finally { if (own) { try { pdf.destroy(); } catch (e) { /* ok */ } } }
}

async function runSearch() {
  const q = $("q").value.trim();
  if (!q) { closeResults(); return; }
  if (!state.index.docs.length) { setStatus("Non ci sono documenti da cercare.", "err"); return; }
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  state.busy = true;
  try {
    const results = [];
    const docs = [...state.index.docs].sort((a, b) => (a.folder + "/" + a.name).localeCompare(b.folder + "/" + b.name, "it"));
    for (let i = 0; i < docs.length; i++) {
      const d = docs[i];
      await ensureDocText(d, `Leggo il testo di «${d.name}» (${i + 1}/${docs.length})…`);
      for (let n = 1; n <= d.pages && results.length < 300; n++) {
        const items = effItems(d.id, n);
        if (!items) continue;
        TextLib.search(TextLib.buildPage(items), q, 20).forEach(hit => results.push({ doc: d, page: n, hit }));
      }
    }
    showResults(q, results);
    const noText = state.index.docs.filter(d => pagesMissing(d).length).length;
    setStatus(results.length ? `${results.length} risultati per «${q}».` + (noText ? ` ${noText} documento/i hanno pagine senza testo: fai l'OCR per includerle.` : "")
      : `Nessun risultato per «${q}».` + (noText ? ` ${noText} documento/i hanno pagine senza testo: fai l'OCR per includerle nella ricerca.` : ""), results.length ? "ok" : "err");
  } catch (e) {
    console.error(e);
    setStatus("Ricerca non riuscita: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

function showResults(q, results) {
  $("results").hidden = false;
  $("resultstitle").textContent = `${results.length} risultati per «${q}»`;
  const ul = $("resultslist");
  ul.innerHTML = "";
  results.forEach(r => {
    const s = r.hit.snippet;
    const li = h("li", {}, h("span", { class: "rdoc", text: r.doc.name }), h("span", { class: "rpage", text: "pag. " + r.page }),
      h("div", {}, s.before, h("mark", { text: s.match }), s.after));
    li.addEventListener("click", () => openResult(r));
    ul.append(li);
  });
}

async function openResult(r) {
  state.focusId = null;
  state.searchMarks = r.hit.boxes.map(b => ({ docId: r.doc.id, page: r.page, ...b }));
  state.pendingScroll = true;
  if (state.docId !== r.doc.id) await openDoc(r.doc.id, { page: r.page, message: `Trovato in «${r.doc.name}», pagina ${r.page}.` });
  else { state.page = r.page; await renderPage(); setStatus(`Trovato in «${r.doc.name}», pagina ${r.page}.`, "ok"); }
}

function closeResults() {
  $("results").hidden = true;
  state.searchMarks = [];
  drawMarks();
}

// =====================================================================
//  Excel: scrittura nelle celle e registro nascosto
// =====================================================================
async function writeToCell(value, color, isFormula) {
  return Excel.run(async ctx => {
    const cell = ctx.workbook.getSelectedRange().getCell(0, 0);
    cell.load("address");
    if (isFormula) cell.formulas = [[value]]; else cell.values = [[value]];
    cell.format.fill.color = color;
    await ctx.sync();
    return cell.address;
  });
}

async function moveRight() {
  return Excel.run(async ctx => {
    const next = ctx.workbook.getSelectedRange().getCell(0, 0).getOffsetRange(0, 1);
    next.select();
    next.load("address");
    await ctx.sync();
    return next.address;
  });
}

function rangeByAddress(ctx, addr) {
  const i = addr.lastIndexOf("!");
  if (i < 0) return ctx.workbook.worksheets.getActiveWorksheet().getRange(addr);
  const sheet = addr.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
  return ctx.workbook.worksheets.getItem(sheet).getRange(addr.slice(i + 1));
}

// aggiunge il testo al contenuto della cella, separato da " - ". Restituisce true se c'era già qualcosa.
async function appendToCell(addr, text, color) {
  return Excel.run(async ctx => {
    const c = rangeByAddress(ctx, addr);
    c.load("text");
    await ctx.sync();
    const old = String((c.text && c.text[0] && c.text[0][0]) || "");
    c.values = [[old ? old + " - " + text : text]];
    c.format.fill.color = color;
    await ctx.sync();
    return !!old;
  });
}

const LOG_HEADER = ["N.", "Data/ora", "Cella", "Tipo", "ID documento", "File", "Pagina", "x1", "y1", "x2", "y2", "Testo letto", "Valore", "Formato numeri"];
const LOG_FORMAT = ["General", "@", "@", "@", "@", "@", "General", "General", "General", "General", "General", "@", "@", "@"];

const nextSnipId = () => state.log.reduce((m, r) => Math.max(m, r.id), 0) + 1;

function rowToArray(r) {
  return [r.id, r.when, r.cell, r.type, r.docId, r.file, r.page, r.u1, r.v1, r.u2, r.v2, r.text, String(r.value), r.fmt];
}

async function getLogSheet(ctx) {
  let sh = ctx.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
  await ctx.sync();
  if (sh.isNullObject) {
    sh = ctx.workbook.worksheets.add(LOG_SHEET);
    const hd = sh.getRange("A1:N1");
    hd.numberFormat = [LOG_FORMAT.map(() => "@")];
    hd.values = [LOG_HEADER];
    hd.format.font.bold = true;
    sh.visibility = Excel.SheetVisibility.hidden;
    await ctx.sync();
  }
  return sh;
}

async function addLog(partial) {
  const doc = curDoc();
  const when = new Date().toLocaleString("it-IT");
  const rows = partial.map(p => ({ when, docId: state.docId, file: doc ? doc.name : "", fmt: $("numfmt").value, ...p }));
  await Excel.run(async ctx => {
    const sh = await getLogSheet(ctx);
    const used = sh.getUsedRange();
    used.load("rowCount");
    await ctx.sync();
    const start = used.rowCount + 1;
    const r = sh.getRangeByIndexes(start - 1, 0, rows.length, 14);
    r.numberFormat = rows.map(() => LOG_FORMAT);
    r.values = rows.map(rowToArray);
    await ctx.sync();
  });
  state.log.push(...rows);
}

async function loadLog() {
  return Excel.run(async ctx => {
    const sh = ctx.workbook.worksheets.getItemOrNullObject(LOG_SHEET);
    await ctx.sync();
    if (sh.isNullObject) return [];
    const used = sh.getUsedRangeOrNullObject();
    used.load("values");
    await ctx.sync();
    if (used.isNullObject) return [];
    return used.values.slice(1).filter(r => r[0] !== "" && r[0] != null).map(r => ({
      id: Number(r[0]), when: String(r[1]), cell: String(r[2]), type: String(r[3]), docId: String(r[4]), file: String(r[5]),
      page: Number(r[6]), u1: Number(r[7]), v1: Number(r[8]), u2: Number(r[9]), v2: Number(r[10]),
      text: String(r[11]), value: r[12], fmt: String(r[13])
    }));
  });
}

async function rewriteLog() {
  return Excel.run(async ctx => {
    const sh = await getLogSheet(ctx);
    const used = sh.getUsedRange();
    used.load("rowCount");
    await ctx.sync();
    if (used.rowCount > 1) sh.getRangeByIndexes(1, 0, used.rowCount - 1, 14).clear();
    if (state.log.length) {
      const r = sh.getRangeByIndexes(1, 0, state.log.length, 14);
      r.numberFormat = state.log.map(() => LOG_FORMAT);
      r.values = state.log.map(rowToArray);
    }
    await ctx.sync();
  });
}

// =====================================================================
//  OCR su tutti i documenti
// =====================================================================
async function runOcrAll() {
  const scope = $("ocrallscope").value;
  const list = state.index.docs.filter(d => scope === "all" || pagesMissing(d).length);
  if (!list.length) { setStatus("Nessun documento ha pagine senza testo da leggere.", "ok"); return; }
  for (let i = 0; i < list.length; i++) {
    setStatus(`OCR documento ${i + 1} di ${list.length}: «${list[i].name}»…`);
    await runOcr(list[i], scope, 1);
  }
  setStatus(`OCR eseguito su ${list.length} documento/i.`, "ok");
}

// =====================================================================
//  Compressione
// =====================================================================
const jpegBlob = (cv, q) => new Promise(res => cv.toBlob(res, "image/jpeg", q));

async function compressOne(doc, level) {
  const before = doc.size;
  const bytes = await loadBytes(doc);
  let out = null, mime = doc.mime, newName = doc.name;
  if (doc.type === "pdf" && level === "light") {
    const p = await PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
    out = await p.save({ useObjectStreams: true });
  } else if (doc.type === "pdf") {
    await ensureDocText(doc, "Leggo il testo prima di comprimere…");
    const pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    const dst = await PDFLib.PDFDocument.create();
    for (let n = 1; n <= pdf.numPages; n++) {
      setStatus(`Comprimo «${doc.name}»: pagina ${n} di ${pdf.numPages}…`);
      const page = await pdf.getPage(n);
      const v1 = page.getViewport({ scale: 1 });
      const vp = page.getViewport({ scale: Math.min(2.5, 1600 / v1.width) });
      const cv = document.createElement("canvas");
      cv.width = Math.floor(vp.width); cv.height = Math.floor(vp.height);
      const cx = cv.getContext("2d");
      cx.fillStyle = "#fff"; cx.fillRect(0, 0, cv.width, cv.height);
      await page.render({ canvasContext: cx, viewport: vp }).promise;
      const jb = new Uint8Array(await (await jpegBlob(cv, 0.72)).arrayBuffer());
      cv.width = cv.height = 0;
      const img = await dst.embedJpg(jb);
      dst.addPage([v1.width, v1.height]).drawImage(img, { x: 0, y: 0, width: v1.width, height: v1.height });
    }
    pdf.destroy();
    out = await dst.save({ useObjectStreams: true });
  } else if (level === "strong") {
    const img = await blobToImage(bytes, doc.mime);
    const k = Math.min(1, 2000 / img.naturalWidth);
    const cv = document.createElement("canvas");
    cv.width = Math.round(img.naturalWidth * k); cv.height = Math.round(img.naturalHeight * k);
    const cx = cv.getContext("2d");
    cx.fillStyle = "#fff"; cx.fillRect(0, 0, cv.width, cv.height);
    cx.drawImage(img, 0, 0, cv.width, cv.height);
    out = new Uint8Array(await (await jpegBlob(cv, 0.75)).arrayBuffer());
    mime = "image/jpeg";
    newName = doc.name.replace(/\.[^.]+$/, "") + ".jpg";
  }
  if (!out || out.length >= before * (level === "strong" ? 0.9 : 0.97)) return { before, after: before, replaced: false };

  if (doc.type === "pdf" && level === "strong") {
    // il testo del PDF originale diventa il "testo letto" del documento, così ricerca e snip continuano a funzionare
    const o = await loadOcr(doc);
    Object.entries(state.native[doc.id] || {}).forEach(([n, it]) => { if (it.length && !(o[n] && o[n].length)) o[n] = it; });
    await saveOcr(doc);
  }
  const old = doc.parts;
  doc.parts = await storeBytes(out);
  doc.size = out.length; doc.mime = mime; doc.name = newName;
  await saveIndex();
  await deleteParts(old).catch(() => {});
  delete state.native[doc.id];
  return { before, after: out.length, replaced: true };
}

async function compressDocs(list, level) {
  if (!list.length) { setStatus("Nessun documento da comprimere.", "err"); return; }
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  state.busy = true;
  let b = 0, a = 0, n = 0;
  try {
    for (const d of list) {
      const r = await compressOne(d, level);
      b += r.before; a += r.after; if (r.replaced) n++;
    }
    const reopen = state.docId && list.some(d => d.id === state.docId) ? state.docId : null;
    state.busy = false;
    refreshAll();
    if (reopen) { const pg = state.page; await openDoc(reopen, { page: pg }); }
    setStatus(n ? `Compressi ${n} documento/i: da ${fmtSize(b)} a ${fmtSize(a)} (risparmio ${fmtSize(b - a)}). Salva il file Excel per rendere effettivo il risparmio.` : "Niente da comprimere: i documenti sono già leggeri.", "ok");
  } catch (e) {
    console.error(e);
    setStatus("Compressione non riuscita: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}
