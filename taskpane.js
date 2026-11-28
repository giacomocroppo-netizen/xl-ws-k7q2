/* global Office, Excel, pdfjsLib, Tesseract, PDFLib, SnipLib, FolderLib, TextLib, ExportLib */
"use strict";

const PDFJS_VERSION = "3.11.174";
const LOG_SHEET = "_SNIP_LOG";
const NS_INDEX = "urn:snipper:index:v1";   // elenco dei documenti incorporati
const NS_DOC = "urn:snipper:doc:v1";       // contenuto dei documenti e dell'OCR (a blocchi)
const CHUNK = 2000000;                     // caratteri per blocco
const WARN_MB = 15;
const OCR_SCALE = 2;                       // risoluzione dell'OCR (2 = circa 150 dpi)
const COL = { snip: "#DDEBF7", sum: "#FFF2CC", val: "#C6EFCE", exc: "#FFC7CE" };
const EXT_OK = /\.(pdf|png|jpe?g)$/i;

const state = {
  mode: "snip",
  index: { docs: [], folders: [] },
  log: [],
  docId: null,
  pdf: null, image: null,
  page: 1, numPages: 1,
  scale: 1.6, zoom: 1, dpr: 1, nav: null, lastCell: null,
  viewW: 0, viewH: 0,
  pageItems: [],            // testo della pagina aperta, in pixel dello schermo
  sumItems: [],
  focusId: null,            // snip evidenziato
  searchMarks: [],          // risultati di ricerca evidenziati
  native: {},               // docId -> {pagina: [pezzi di testo normalizzati]}
  ocr: {},                  // docId -> {pagina: [parole normalizzate]}
  ocrDismissed: new Set(),
  ui: { folderFilter: new Set(), expanded: new Set(), creatingIn: null, renaming: null, selected: new Set(), currentFolder: "" },
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
let readyDone;
const readyP = new Promise(r => { readyDone = r; });

async function showPane() {
  try { if (Office.addin && Office.addin.showAsTaskpane) await Office.addin.showAsTaskpane(); } catch (e) { /* già aperto */ }
}
const ribbon = fn => async ev => {
  try { await readyP; await fn(); } catch (e) { console.warn(e); }
  ev.completed();
};

Office.onReady(async info => {
  try {
    Office.actions.associate("setModeSnip", ribbon(async () => { showTab("snip"); setMode("snip"); await showPane(); }));
    Office.actions.associate("setModeSum", ribbon(async () => { showTab("snip"); setMode("sum"); await showPane(); }));
    Office.actions.associate("setModeVal", ribbon(async () => { showTab("snip"); setMode("val"); await showPane(); }));
    Office.actions.associate("setModeExc", ribbon(async () => { showTab("snip"); setMode("exc"); await showPane(); }));
    Office.actions.associate("openOrganizer", ribbon(async () => { showTab("org"); setOrgTab("docs"); await showPane(); }));
    Office.actions.associate("exportAll", ribbon(async () => { showTab("set"); await showPane(); await exportAll(); }));
    Office.actions.associate("openSettings", ribbon(async () => { showTab("set"); await showPane(); }));
    Office.actions.associate("importDocs", ribbon(async () => { showTab("org"); setOrgTab("docs"); await showPane(); flashDropzone(); }));
    Office.actions.associate("deleteSnip", ribbon(async () => { await showPane(); await deleteOnSelection(); }));
  } catch (e) { console.warn("Comandi della barra non disponibili", e); }
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
  readyDone();
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
  $("orgclose").addEventListener("click", () => showTab("snip"));
  ["docs", "list"].forEach(k => $("stab-" + k).addEventListener("click", () => setOrgTab(k)));
  $("foldersbtn").addEventListener("click", () => togglePop("pop-folders"));
  $("docbtn").addEventListener("click", () => togglePop("pop-docs"));
  $("morebtn").addEventListener("click", () => togglePop("pop-more"));
  $("popimport").addEventListener("click", () => { showTab("org"); setOrgTab("docs"); flashDropzone(); });
  $("folderall").addEventListener("click", () => { state.ui.folderFilter.clear(); renderFolderList(); renderDocList(); updateDocBtn(); });
  $("docfilter2").addEventListener("input", renderDocList);
  document.querySelectorAll(".popclose").forEach(b => b.addEventListener("click", () => closePops()));
  document.addEventListener("click", e => { if (!e.target.closest(".pop, #foldersbtn, #docbtn, #morebtn, #popimport")) closePops(); });
  $("pageinput").addEventListener("change", () => gotoPage(parseInt($("pageinput").value, 10) || 1));
  $("zoomfit").addEventListener("click", () => { state.zoom = 1; scheduleRender(); });
  $("ocrcur").addEventListener("click", () => { const d = curDoc(); if (!d) return setStatus("Apri prima un documento.", "err"); closePops(); runOcr(d, pagesMissing(d).length ? "missing" : "all", state.page); });
  $("exportall").addEventListener("click", () => exportAll());
  $("dlorig").addEventListener("click", () => { const d = curDoc(); if (d) downloadOriginal(d); });
  $("dlann").addEventListener("click", () => { const d = curDoc(); if (d) downloadAnnotated(d); });
  ["lfilter", "ltype"].forEach(i => $(i).addEventListener("input", renderList));
  $("zoomin").addEventListener("click", () => zoom(1.2));
  $("zoomout").addEventListener("click", () => zoom(1 / 1.2));
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
  $("selclear").addEventListener("click", () => { state.ui.selected.clear(); selFolders().clear(); renderOrg(); });
  $("newfolder").addEventListener("click", () => { state.ui.creatingIn = state.ui.currentFolder; $("docfilter").value = ""; renderOrg(); });
  $("selcompress").addEventListener("click", () => {
    const set = new Map();
    [...state.ui.selected].forEach(id => { const d = docById(id); if (d) set.set(d.id, d); });
    [...selFolders()].forEach(p => docsUnder(p).forEach(d => set.set(d.id, d)));
    compressDocs([...set.values()], $("cmplevel").value);
  });
  armConfirm($("seldelete"), "🗑 Elimina", async () => {
    const ids = [...state.ui.selected];
    for (const id of ids) { const d = docById(id); if (d) await removeDoc(d, true); }
    const fl = [...selFolders()].sort((x, y) => y.length - x.length);
    fl.forEach(p => {
      if (!state.index.folders.includes(p)) return;
      FolderLib.deleteFolder(state.index, p);
      if (FolderLib.isInside(state.ui.currentFolder, p)) state.ui.currentFolder = FolderLib.parentOf(p);
    });
    state.ui.selected.clear(); selFolders().clear();
    if (fl.length) await saveIndex();
    refreshAll();
  });
  try { $("autoopen").checked = localStorage.getItem("snipper.autoopen") !== "0"; } catch (e) { /* ok */ }
  $("autoopen").addEventListener("change", () => { try { localStorage.setItem("snipper.autoopen", $("autoopen").checked ? "1" : "0"); } catch (e) { /* ok */ } });
  setMode("snip");
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
  const org = t === "org" || t === "list" || t === "set";
  $("orgtabs").hidden = t === "set";
  $("orgtitle").textContent = t === "set" ? "Settings" : "";
  $("view-snip").hidden = org;
  $("view-org").hidden = !org;
  closePops();
  if (t === "list") setOrgTab("list");
  if (t === "set") setOrgTab("tools");
  if (t === "org") renderOrg();
  if (t === "snip" && state.docId) scheduleRender();
  if (t === "snip") ensureDocOpen();
}

function setOrgTab(k) {
  ["docs", "list", "tools"].forEach(x => { $("org-" + x).hidden = x !== k; const b = $("stab-" + x); if (b) b.classList.toggle("active", x === k); });
  if (k === "docs") renderOrg();
  if (k === "list") renderList();
}

function flashDropzone() {
  const z = $("dropzone");
  z.classList.remove("flash");
  void z.offsetWidth;
  z.classList.add("flash");
  z.scrollIntoView({ block: "nearest" });
}

function closePops() { ["pop-folders", "pop-docs", "pop-more"].forEach(i => { $(i).hidden = true; }); }
function togglePop(id) {
  const open = $(id).hidden;
  closePops();
  $(id).hidden = !open;
  if (open && id === "pop-folders") renderFolderList();
  if (open && id === "pop-docs") { $("docfilter2").value = ""; renderDocList(); $("docfilter2").focus(); }
}

function folderMatches(d) {
  const f = state.ui.folderFilter;
  if (!f.size) return true;
  return [...f].some(p => p === "" ? d.folder === "" : (d.folder === p || FolderLib.isInside(d.folder, p)));
}
const visibleDocs = () => state.index.docs.filter(folderMatches).sort((a, b) => (a.folder + "/" + a.name).localeCompare(b.folder + "/" + b.name, "it"));

function renderFolderList() {
  const box = $("folderlist");
  box.innerHTML = "";
  const opts = FolderLib.options(state.index).filter(o => o.path !== "");
  if (!opts.length) box.append(h("div", { class: "muted wrap pad", text: "Non ci sono cartelle: creale in Organizer." }));
  opts.forEach(o => {
    const cb = h("input", { type: "checkbox" });
    cb.checked = state.ui.folderFilter.has(o.path);
    const it = h("div", { class: "it" }, cb, h("span", { text: o.label }));
    const toggle = () => { if (state.ui.folderFilter.has(o.path)) state.ui.folderFilter.delete(o.path); else state.ui.folderFilter.add(o.path); renderFolderList(); renderDocList(); updateDocBtn(); };
    it.addEventListener("click", e => { if (e.target !== cb) toggle(); else toggle(); });
    box.append(it);
  });
}

function renderDocList() {
  const q = ($("docfilter2").value || "").toLowerCase().trim();
  const box = $("doclist");
  box.innerHTML = "";
  const list = visibleDocs().filter(d => !q || d.name.toLowerCase().includes(q));
  if (!list.length) box.append(h("div", { class: "muted wrap pad", text: state.index.docs.length ? "Nessun documento corrisponde." : "Nessun documento importato." }));
  list.forEach(d => {
    const it = h("div", { class: "it" + (d.id === state.docId ? " cur" : "") }, h("span", { text: (d.type === "pdf" ? "📄 " : "🖼️ ") + d.name }), d.folder ? h("span", { class: "pth", text: d.folder }) : null);
    it.addEventListener("click", () => { closePops(); openDoc(d.id); });
    box.append(it);
  });
}

function updateDocBtn() {
  const d = curDoc();
  const f = state.ui.folderFilter.size;
  $("foldersbtn").textContent = "Fold" + (f ? ` (${f})` : "…") + " ▾";
  $("docbtn").textContent = (d ? d.name : (state.index.docs.length ? "Scegli documento" : "(nessun documento)")) + " ▾";
}

function updatePageUI() {
  $("pagebadge").hidden = !state.docId;
  $("pageinput").value = state.docId ? state.page : 1;
  $("pagetotal").textContent = state.docId ? state.numPages : "–";
}

const MODE_HELP = {
  snip: "SNIP: trascina un rettangolo sul valore → cella selezionata. Lo snip successivo va nella cella a destra. CTRL = aggiungi alla stessa cella.",
  sum: "SUM SNIP: trascina sul numero (o su una colonna di numeri) → somma nella cella. Lo snip successivo va a destra. CTRL = aggiungi alla somma della cella.",
  val: "VALIDATION: trascina sull'area verificata → cella verde con ✓. Con CTRL altri riquadri sulla stessa cella (un solo ✓).",
  exc: "EXCEPTION: trascina sull'area con anomalia → cella rossa con ✗. Con CTRL altri riquadri sulla stessa cella (un solo ✗)."
};
function setMode(m) {
  state.mode = m;
  $("modestrip").className = "modestrip " + m;
  setStatus(MODE_HELP[m]);
}

function refreshAll() {
  renderDocSelect();
  renderOrg();
  updateOcrBanner();
  updateOcrInfo();
  ensureDocOpen();
}

// Se nel file c'è almeno un documento, il visualizzatore ne mostra sempre uno (niente banner di importazione).
function ensureDocOpen() {
  if (state.docId || state.opening || state.busy || !state.index.docs.length || $("view-snip").hidden) return;
  const d = (state.lastDocId && docById(state.lastDocId)) || state.index.docs[0];
  if (d) openDoc(d.id);
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
    const dx = t[0] / sx, dy = t[1] / sx, ux = t[2] / sy, uy = t[3] / sy;
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
    if (last && !$("view-snip").hidden && (added.length === 1 || (!added.length && dup === 1))) await openDoc(last);
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
    ensureDocOpen();
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
  updateDocBtn();
  if (!$("pop-docs").hidden) renderDocList();
  if (!$("pop-folders").hidden) renderFolderList();
}

function snipCount(docId) {
  return new Set(state.log.filter(r => r.docId === docId).map(r => r.id)).size;
}

function findNode(node, path) {
  if (node.path === path) return node;
  for (const c of node.children) { const r = findNode(c, path); if (r) return r; }
  return null;
}

function docsUnder(path) { return state.index.docs.filter(d => FolderLib.isInside(d.folder, path)); }

function selFolders() { return state.ui.selectedFolders || (state.ui.selectedFolders = new Set()); }

function renderOrg() {
  const q = ($("docfilter").value || "").toLowerCase().trim();
  const total = state.index.docs.reduce((a, d) => a + d.size, 0);
  $("orgsummary").textContent = state.index.docs.length
    ? `${state.index.docs.length} documenti · ${state.index.folders.length} cartelle · ${fmtSize(total)} incorporati nel file`
    : "Nessun documento importato.";

  const tree = FolderLib.buildTree(state.index);
  let cur = findNode(tree, state.ui.currentFolder);
  if (!cur) { state.ui.currentFolder = ""; cur = tree; }

  // percorso (come la barra degli indirizzi di Esplora risorse)
  const cr = $("crumb");
  cr.innerHTML = "";
  const segs = [{ path: "", label: "🏠 Principale" }];
  FolderLib.cleanPath(state.ui.currentFolder).split("/").filter(Boolean).forEach((p, i, arr) => segs.push({ path: arr.slice(0, i + 1).join("/"), label: p }));
  segs.forEach((s, i) => {
    if (i) cr.append(h("span", { class: "csep", text: "›" }));
    const el = h("span", { class: "cseg" + (i === segs.length - 1 ? " last" : ""), text: s.label, onclick: () => { state.ui.currentFolder = s.path; renderOrg(); } });
    attachDrop(el, s.path);
    cr.append(el);
  });

  // barra della selezione
  const sel = state.ui.selected, selF = selFolders();
  [...sel].forEach(id => { if (!docById(id)) sel.delete(id); });
  [...selF].forEach(p => { if (!state.index.folders.includes(p)) selF.delete(p); });
  $("selbar").hidden = !(sel.size || selF.size);
  if (sel.size || selF.size) {
    const parts = [];
    if (selF.size) parts.push(`${selF.size} ${selF.size === 1 ? "cartella" : "cartelle"}`);
    if (sel.size) parts.push(`${sel.size} ${sel.size === 1 ? "documento" : "documenti"}`);
    $("selcount").textContent = parts.join(" + ");
  }

  const box = $("tree");
  box.innerHTML = "";
  if (q) {
    const found = state.index.docs.filter(d => d.name.toLowerCase().includes(q));
    if (!found.length) box.append(h("div", { class: "muted pad", text: "Nessun documento trovato." }));
    found.sort((a, b) => a.name.localeCompare(b.name, "it")).forEach(d => box.append(docCard(d, true)));
    return;
  }
  if (state.ui.creatingIn !== null && state.ui.creatingIn !== undefined) box.append(newFolderInput(state.ui.creatingIn));
  cur.children.forEach(c => box.append(folderItem(c)));
  cur.docs.forEach(d => box.append(docCard(d, false)));
  if (!cur.children.length && !cur.docs.length && state.ui.creatingIn === null) box.append(h("div", { class: "muted pad", text: "Cartella vuota. Trascina qui documenti o cartelle, oppure usa Importa." }));
}

// trascinare documenti o cartelle su una cartella (o sul percorso) per spostarli
function attachDrop(el, path) {
  el.addEventListener("dragover", e => {
    const t = Array.from(e.dataTransfer.types || []);
    if (t.includes("text/x-snipper-doc") || t.includes("text/x-snipper-folder")) { e.preventDefault(); el.classList.add("dragover"); }
  });
  el.addEventListener("dragleave", () => el.classList.remove("dragover"));
  el.addEventListener("drop", e => {
    el.classList.remove("dragover");
    const ids = e.dataTransfer.getData("text/x-snipper-doc");
    const fol = e.dataTransfer.getData("text/x-snipper-folder");
    if (ids) { e.preventDefault(); FolderLib.moveDocs(state.index, ids.split(","), path); state.ui.selected.clear(); saveIndex(); refreshAll(); }
    else if (fol) {
      e.preventDefault();
      if (FolderLib.moveFolder(state.index, fol, path) === null) setStatus("Non posso spostare la cartella lì.", "err");
      else { selFolders().clear(); saveIndex(); refreshAll(); }
    }
  });
}

function folderItem(node) {
  const selF = selFolders();
  const row = h("div", { class: "frow" + (selF.has(node.path) ? " sel" : "") });
  const cb = h("input", { type: "checkbox", title: "Seleziona la cartella" });
  cb.checked = selF.has(node.path);
  cb.addEventListener("click", ev => ev.stopPropagation());
  cb.addEventListener("change", () => { if (cb.checked) selF.add(node.path); else selF.delete(node.path); renderOrg(); });
  row.append(cb);
  if (state.ui.renaming === node.path) {
    const inp = h("input", { value: node.name });
    const done = ok => {
      const v = inp.value;
      state.ui.renaming = null;
      if (ok && v.trim() && v.trim() !== node.name) {
        const np = FolderLib.renameFolder(state.index, node.path, v);
        if (np === null) setStatus("Esiste già una cartella con quel nome.", "err");
        else { if (state.ui.currentFolder === node.path || FolderLib.isInside(state.ui.currentFolder, node.path)) state.ui.currentFolder = np + state.ui.currentFolder.slice(node.path.length); selF.clear(); saveIndex(); refreshAll(); return; }
      }
      renderOrg();
    };
    inp.addEventListener("keydown", e => { if (e.key === "Enter") done(true); else if (e.key === "Escape") done(false); });
    inp.addEventListener("blur", () => done(true));
    inp.addEventListener("click", e => e.stopPropagation());
    row.append(h("span", { text: "📁" }), inp);
    setTimeout(() => { inp.focus(); inp.select(); }, 0);
  } else {
    row.append(h("span", { text: "📁" }),
      h("span", { class: "fname", text: node.name }),
      h("span", { class: "fcount", text: String(FolderLib.countDocs(node)) }));
  }
  const icons = h("span", { class: "ficons" });
  icons.append(h("button", { title: "Rinomina", text: "✎", onclick: ev => { ev.stopPropagation(); state.ui.renaming = node.path; renderOrg(); } }));
  const del = h("button", { title: "Elimina cartella (il contenuto passa alla cartella superiore)", text: "🗑" });
  armConfirm(del, "🗑", () => {
    FolderLib.deleteFolder(state.index, node.path);
    if (FolderLib.isInside(state.ui.currentFolder, node.path)) state.ui.currentFolder = FolderLib.parentOf(node.path);
    selF.clear();
    saveIndex(); refreshAll();
    setStatus(`Cartella «${node.name}» eliminata: i documenti sono passati alla cartella superiore.`, "ok");
  });
  del.addEventListener("click", ev => ev.stopPropagation());
  icons.append(del);
  row.append(icons);
  row.addEventListener("click", () => { state.ui.currentFolder = node.path; renderOrg(); });
  attachDrop(row, node.path);
  row.setAttribute("draggable", "true");
  row.addEventListener("dragstart", e => { e.dataTransfer.setData("text/x-snipper-folder", node.path); e.dataTransfer.effectAllowed = "move"; });
  return row;
}

function newFolderInput(parent) {
  const inp = h("input", { placeholder: "Nome della nuova cartella" });
  const row = h("div", { class: "frow" }, h("span", { text: "📁" }), inp);
  const done = ok => {
    const v = inp.value.trim();
    state.ui.creatingIn = null;
    if (ok && v) { FolderLib.addFolder(state.index, parent, v); saveIndex(); refreshAll(); }
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
  const dots = h("button", { class: "dots", title: "Azioni sul documento", text: "⋯" });
  dots.addEventListener("click", ev => { ev.stopPropagation(); openDocMenu(d, dots); });
  card.append(h("div", { class: "top" }, cb, h("div", { class: "dbody" }, h("div", { class: "docname", text: (d.type === "pdf" ? "📄 " : "🖼️ ") + d.name }), meta,
    showPath && d.folder ? h("div", { class: "meta", text: "📁 " + d.folder }) : null), dots));

  return card;
}

let docMenuEl = null;
function closeDocMenu() { if (docMenuEl) { docMenuEl.remove(); docMenuEl = null; } }
document.addEventListener("click", e => { if (docMenuEl && !e.target.closest(".docmenu")) closeDocMenu(); });

function openDocMenu(d, anchor) {
  closeDocMenu();
  const n = snipCount(d.id);
  const miss = pagesMissing(d).length;
  const m = h("div", { class: "docmenu" });
  const item = (txt, fn) => {
    const b = h("button", { text: txt });
    b.addEventListener("click", ev => { ev.stopPropagation(); closeDocMenu(); fn(); });
    m.append(b);
    return b;
  };
  item("📂 Apri", () => { showTab("snip"); openDoc(d.id); });
  const row = h("div", { class: "dmrow" }, h("span", { text: "Sposta in:" }));
  const move = h("select", {});
  FolderLib.options(state.index).forEach(o => move.appendChild(new Option(o.label, o.path)));
  move.value = d.folder;
  move.addEventListener("click", ev => ev.stopPropagation());
  move.addEventListener("change", () => { FolderLib.moveDocs(state.index, [d.id], move.value); closeDocMenu(); saveIndex(); renderOrg(); renderDocSelect(); });
  row.append(move);
  m.append(row);
  item("⬇ Scarica originale", () => downloadOriginal(d));
  item(`⬇ Scarica con spunte${n ? " (" + n + ")" : ""}`, () => downloadAnnotated(d));
  item("🗜 Comprimi", () => compressDocs([d], $("cmplevel").value));
  item("🔎 OCR di questo documento", () => runOcr(d, miss ? "missing" : "all", 1));
  const del = h("button", { text: "🗑 Elimina documento", class: "danger" });
  let armed = false;
  del.addEventListener("click", ev => {
    ev.stopPropagation();
    if (!armed) { armed = true; del.textContent = "Sicuro? Clicca di nuovo"; return; }
    closeDocMenu();
    removeDoc(d);
  });
  m.append(del);
  document.body.append(m);
  docMenuEl = m;
  const r = anchor.getBoundingClientRect(), mh = m.offsetHeight, mw = m.offsetWidth;
  let top = r.bottom + 4;
  if (top + mh > window.innerHeight - 4) top = Math.max(4, r.top - mh - 4);
  m.style.top = top + "px";
  m.style.left = Math.max(4, Math.min(window.innerWidth - mw - 4, r.right - mw)) + "px";
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
    ensureDocOpen();
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
//  Esporta tutto: zip con originali, documenti con gli snip e tabella
// =====================================================================
const safeName = n => String(n).replace(/[\\/:*?"<>|\u0000-\u001F]/g, "_");
function splitCell(addr) {
  const i = addr.lastIndexOf("!");
  if (i < 0) return ["", addr];
  return [addr.slice(0, i).replace(/^'(.*)'$/, "$1").replace(/''/g, "'"), addr.slice(i + 1)];
}
// righe della tabella: una per ogni snip; i documenti senza snip compaiono con tipo vuoto
function exportRows(docs, log) {
  const out = [["Nome documento", "Cartella", "Tipologia snip", "Foglio", "Cella", "Pagina"]];
  for (const d of docs) {
    const mine = log.filter(r => r.docId === d.id);
    const ids = [...new Set(mine.map(r => r.id))].sort((a, b) => a - b);
    if (!ids.length) { out.push([d.name, d.folder || "", "", "", "", ""]); continue; }
    for (const id of ids) {
      const rs = mine.filter(r => r.id === id);
      const [sheet, cell] = splitCell(rs[0].cell || "");
      const pages = [...new Set(rs.map(r => r.page))].sort((a, b) => a - b).join(", ");
      out.push([d.name, d.folder || "", rs[0].type, sheet, cell, pages]);
    }
  }
  return out;
}
async function exportAll() {
  if (state.busy) return;
  if (!state.index.docs.length) { setStatus("Non ci sono documenti da esportare in questo file.", "err"); return; }
  state.busy = true;
  try {
    const docs = [...state.index.docs].sort((a, b) => (a.folder + "/" + a.name).localeCompare(b.folder + "/" + b.name, "it"));
    const entries = [], used = new Set(), problems = [];
    const uniq = path => { let p = path, n = 2; while (used.has(p.toLowerCase())) { p = path.replace(/(\.[^./]*)?$/, m => ` (${n++})${m}`); } used.add(p.toLowerCase()); return p; };
    let i = 0;
    for (const d of docs) {
      i++;
      setStatus(`Esporto ${i}/${docs.length}: «${d.name}»…`);
      const dir = d.folder ? d.folder.split("/").map(safeName).join("/") + "/" : "";
      const bytes = await loadBytes(d);
      entries.push({ name: uniq("1 - Originali/" + dir + safeName(d.name)), data: new Uint8Array(bytes) });
      const rows = state.log.filter(r => r.docId === d.id);
      if (rows.length) {
        try {
          const out = await ExportLib.annotate(PDFLib, bytes, d, rows);
          entries.push({ name: uniq("2 - Con snip/" + dir + safeName(d.name.replace(/\.[^.]+$/, "")) + "_con_snip.pdf"), data: new Uint8Array(out) });
        } catch (e) { console.error(e); problems.push(d.name); }
      }
    }
    entries.push({ name: "3 - Tabella snip.xlsx", data: ZipLib.xlsx(exportRows(docs, state.log), "Snip", [40, 25, 16, 20, 10, 10]) });
    setStatus("Creo lo zip…");
    const zip = ZipLib.zip(entries);
    const d = new Date(), p2 = n => String(n).padStart(2, "0");
    const fname = `Snipper_export_${d.getFullYear()}-${p2(d.getMonth() + 1)}-${p2(d.getDate())}.zip`;
    saveBlob(new Blob([zip], { type: "application/zip" }), fname);
    setStatus(`Creato «${fname}»: ${docs.length} documenti, ${new Set(state.log.map(r => r.id)).size} snip.` +
      (problems.length ? ` Non sono riuscito a disegnare gli snip su: ${problems.join(", ")} (gli originali ci sono).` : "") +
      " Se il download non parte, premi di nuovo «Esporta tutto» nelle Impostazioni.", problems.length ? "err" : "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
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
  updatePageUI();
  updateOcrInfo();
  updateDocBtn();
  drawMarks();
}

async function openDoc(id, opts) {
  opts = opts || {};
  const d = docById(id);
  if (!d) return;
  state.opening = true;
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
    state.lastDocId = id;
    state.page = Math.min(Math.max(1, opts.page || 1), state.numPages);
    state.sumItems = []; renderSum();
    state.zoom = 1;
    updateDocBtn();
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
  } finally {
    state.opening = false;
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
  updatePageUI();
  drawMarks();
  if (state.pendingScroll) { state.pendingScroll = false; scrollToMark(); }
}

async function gotoPage(n, where) {
  if (!state.docId || n < 1 || n > state.numPages) { updatePageUI(); return; }
  state.page = n;
  await renderPage();
  const v = $("viewer");
  if (where === "top") v.scrollTop = 0;
  else if (where === "bottom") v.scrollTop = v.scrollHeight;
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
  let lastFlip = 0;
  window.addEventListener("wheel", e => {
    if (!e.ctrlKey) {
      const v = $("viewer");
      // con la rotellina/touchpad ci si muove solo in alto e in basso; a destra e sinistra solo con le barre
      if (v.contains(e.target) && (e.deltaX !== 0 || e.shiftKey)) {
        e.preventDefault();
        if (!e.shiftKey && e.deltaY) v.scrollTop += e.deltaY;
        return;
      }
      if (!state.docId || !v.contains(e.target) || Math.abs(e.deltaY) < 1) return;
      const atBottom = v.scrollTop + v.clientHeight >= v.scrollHeight - 3;
      const atTop = v.scrollTop <= 0;
      // scorrimento verticale più lento (circa un terzo del normale)
      e.preventDefault();
      v.scrollTop += e.deltaY * (e.deltaMode === 1 ? 16 : 1) * 0.3;
      if (Math.abs(e.deltaY) < 4) return;
      const now = Date.now();
      if (now - lastFlip < 600) return;
      if (e.deltaY > 0 && atBottom && state.page < state.numPages) { lastFlip = now; gotoPage(state.page + 1, "top"); }
      else if (e.deltaY < 0 && atTop && state.page > 1) { lastFlip = now; gotoPage(state.page - 1, "bottom"); }
      return;
    }
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

function hideHoverX() { const x = overlay.querySelector(".hoverx"); if (x) x.remove(); }

function drawMarks() {
  overlay.querySelectorAll(".mark").forEach(n => n.remove());
  hideHoverX();
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
    return el;
  };
  const focusing = state.focusId != null;
  let xDone = false;
  state.log.filter(r => r.docId === state.docId && r.page === state.page).forEach(r => {
    const isF = r.id === state.focusId;
    const tc = r.type === "SUM SNIP" ? "sum " : r.type === "VALIDATION" ? "val " : r.type === "EXCEPTION" ? "exc " : "";
    const el = add(r, tc + (isF ? "focus" : (focusing ? "dim" : "")), "#" + r.id);
    if (isF && !xDone) {
      // X per eliminare lo snip evidenziato (come il tasto Delete Snip)
      xDone = true;
      const x = h("button", { class: "markx", title: "Elimina questo snip (definitivo, come Delete Snip)", text: "✕" });
      x.addEventListener("mousedown", ev => ev.stopPropagation());
      x.addEventListener("click", ev => { ev.stopPropagation(); ev.preventDefault(); deleteSnipPart(r); });
      el.append(x);
    }
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
  // passando col cursore sul riquadro di uno snip compare una X per eliminarlo
  overlay.addEventListener("mousemove", e => {
    if (start || !state.docId) { hideHoverX(); return; }
    const r = overlay.getBoundingClientRect();
    const px = e.clientX - r.left, py = e.clientY - r.top, W = state.viewW, H = state.viewH, M = 14;
    let best = null, area = Infinity;
    state.log.forEach(m => {
      if (m.docId !== state.docId || m.page !== state.page || m.id === state.focusId) return;
      const x1 = Math.min(m.u1, m.u2) * W, x2 = Math.max(m.u1, m.u2) * W, y1 = Math.min(m.v1, m.v2) * H, y2 = Math.max(m.v1, m.v2) * H;
      if (px < x1 - M || px > x2 + M || py < y1 - M || py > y2 + M) return;
      const a = (x2 - x1) * (y2 - y1);
      if (a < area) { area = a; best = { row: m, x2, y1 }; }
    });
    let x = overlay.querySelector(".hoverx");
    if (!best) { if (x) x.remove(); return; }
    if (!x) {
      x = h("button", { class: "markx hoverx", title: "Elimina questo snip (definitivo, come Delete Snip)", text: "✕" });
      x.addEventListener("mousedown", ev => ev.stopPropagation());
      x.addEventListener("click", ev => { ev.stopPropagation(); ev.preventDefault(); const row = x._row; x.remove(); if (row) deleteSnipPart(row); });
      overlay.append(x);
    }
    x._row = best.row;
    x.style.left = (best.x2 - 11) + "px";
    x.style.top = (best.y1 - 11) + "px";
  });
  overlay.addEventListener("mouseleave", () => hideHoverX());
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
    const down = start;
    start = null;
    if (box) { box.remove(); box = null; }
    if (tiny) {
      // clic (senza trascinare) su un riquadro già fatto: Excel va alla sua cella
      let hit = null, area = Infinity;
      state.log.forEach(m => {
        if (m.docId !== state.docId || m.page !== state.page) return;
        const x1 = Math.min(m.u1, m.u2) * p.w, x2 = Math.max(m.u1, m.u2) * p.w, y1 = Math.min(m.v1, m.v2) * p.h, y2 = Math.max(m.v1, m.v2) * p.h;
        if (down.x < x1 || down.x > x2 || down.y < y1 || down.y > y2) return;
        const a = (x2 - x1) * (y2 - y1);
        if (a < area) { area = a; hit = m; }
      });
      if (hit && hit.cell) {
        try {
          await selectCell(hit.cell);
          state.focusId = hit.id; state.searchMarks = [];
          drawMarks();
          setStatus(`Snip #${hit.id} · cella ${hit.cell}`, "ok");
        } catch (err) { setStatus("Errore: " + (err.message || err), "err"); }
      }
      return;
    }
    await onSnip(f, withCtrl);
  });
}

// =====================================================================
//  Lettura dell'area: testo (PDF o OCR salvato), altrimenti OCR al volo
// =====================================================================
async function readArea(f) {
  const mode = $("readmode").value;
  let layer = "", layerSrc = "testo PDF";
  if (mode !== "ocr" && state.pageItems.length) {
    layer = SnipLib.textInRect(state.pageItems, { x1: f.u1 * state.viewW, y1: f.v1 * state.viewH, x2: f.u2 * state.viewW, y2: f.v2 * state.viewH });
    const o = state.ocr[state.docId];
    if (o && o[state.page] && o[state.page].length) layerSrc = "OCR salvato";
  }
  if (mode === "text") return { text: layer, source: layerSrc, layer };

  // verifica con l'OCR su ciò che si vede davvero nel rettangolo
  let r = null;
  try { r = await ocrArea(f); } catch (e) { console.warn("OCR sull'area non riuscito", e); }
  if (!r) return { text: layer, source: layerSrc + " (OCR di verifica non disponibile)", layer };
  if (r.ink < 0.003) return { text: "", source: "", layer: "" };                          // area vuota
  if (!r.text) return { text: layer, source: layerSrc, layer };
  if (!layer) return { text: r.text, source: "OCR sull'area", layer };
  const norm = x => x.replace(/\s+/g, "").toLowerCase();
  return norm(layer) === norm(r.text)
    ? { text: layer, source: layerSrc + ", verificato con OCR", layer }
    : { text: r.text, source: "OCR sull'area: il testo nascosto del PDF era diverso («" + layer + "»)", layer };
}

let areaWorker = null;
async function ocrArea(f) {
  setStatus("Leggo l'area… (la prima volta scarica i dati della lingua)");
  const sx = f.u1 * canvas.width, sy = f.v1 * canvas.height;
  const sw = Math.max(1, (f.u2 - f.u1) * canvas.width), sh = Math.max(1, (f.v2 - f.v1) * canvas.height);
  // quanto "inchiostro" c'è nel rettangolo: se è bianco, non c'è nulla da leggere
  const probe = canvas.getContext("2d").getImageData(Math.floor(sx), Math.floor(sy), Math.ceil(sw), Math.ceil(sh));
  let dark = 0;
  for (let i = 0; i < probe.data.length; i += 4) if (probe.data[i] * 0.3 + probe.data[i + 1] * 0.59 + probe.data[i + 2] * 0.11 < 170) dark++;
  const ink = dark / (probe.data.length / 4);
  if (ink < 0.003) return { text: "", ink };
  const up = Math.max(1, Math.min(4, 120 / Math.max(sh, 1)));
  const tmp = document.createElement("canvas");
  tmp.width = Math.ceil(sw * up) + 40;
  tmp.height = Math.ceil(sh * up) + 40;
  const c = tmp.getContext("2d");
  c.fillStyle = "#fff"; c.fillRect(0, 0, tmp.width, tmp.height);
  c.imageSmoothingQuality = "high";
  c.drawImage(canvas, sx, sy, sw, sh, 20, 20, sw * up, sh * up);
  if (!areaWorker) areaWorker = await Tesseract.createWorker(["ita", "eng"]);
  const { data } = await areaWorker.recognize(tmp);
  return { text: SnipLib.cleanRead(data.text), ink };
}

// chiede all'utente il testo da inserire quando nell'area non c'è nulla di leggibile
function askManualText(msg) {
  return new Promise(resolve => {
    const p = $("manualpanel"), i = $("manualtext");
    $("manualmsg").textContent = msg;
    i.value = "";
    p.hidden = false;
    i.focus();
    const done = v => { p.hidden = true; $("manualok").onclick = $("manualcancel").onclick = i.onkeydown = null; resolve(v); };
    $("manualok").onclick = () => done(i.value.trim() || null);
    $("manualcancel").onclick = () => done(null);
    i.onkeydown = e => { if (e.key === "Enter") done(i.value.trim() || null); else if (e.key === "Escape") done(null); };
  });
}

// =====================================================================
//  SNIP / SUM SNIP
// =====================================================================
async function onSnip(f, withCtrl) {
  if (state.busy) return;
  state.busy = true;
  try {
    if (state.mode === "val" || state.mode === "exc") {
      const isV = state.mode === "val";
      const sym = isV ? "✓" : "✗";
      const type = isV ? "VALIDATION" : "EXCEPTION";
      let addr, id, extra = false;
      if (withCtrl) {
        addr = await currentCellAddress();
        const ex = state.log.find(r => r.cell === addr && r.type === type);
        if (ex) { id = ex.id; extra = true; }
        else { addr = await writeToCell(sym, isV ? COL.val : COL.exc, false, isV ? "#006100" : "#9C0006"); id = nextSnipId(); }
      } else {
        await advanceIfNeeded();
        addr = await writeToCell(sym, isV ? COL.val : COL.exc, false, isV ? "#006100" : "#9C0006");
        id = nextSnipId();
      }
      await addLog([{ id, cell: addr, type, page: state.page, ...f, text: sym, value: sym }]);
      state.focusId = null; state.searchMarks = [];
      drawMarks();
      state.lastCell = addr;
      setStatus(`#${id}: ${isV ? "validazione" : "eccezione"} ${extra ? "aggiunta (altro riquadro) alla cella" : "in"} ${addr}.`, "ok");
      return;
    }
    let { text, source, layer } = await readArea(f);
    if (!text) {
      setStatus("Nessun testo trovato nell'area selezionata.", "err");
      text = await askManualText("Nell'area non c'è testo leggibile. Scrivi il testo da inserire nella cella:");
      if (!text) { setStatus("Snip annullato.", "err"); return; }
      source = "testo inserito a mano";
    }
    const fmt = $("numfmt").value;

    if (state.mode === "sum") {
      // dall'area prende SOLO i numeri (anche tanti, in colonna o in riga) e ignora le parole
      const ex = SnipLib.extractNumbers(text, fmt);
      if (!ex.nums.length) {
        const one = SnipLib.parseNumber(text, fmt);
        if (one === null) { setStatus(`Nell'area non ho trovato numeri validi: «${text.length > 50 ? text.slice(0, 50) + "…" : text}» (formato: ${fmt === "it" ? "italiano" : "inglese"}).`, "err"); return; }
        ex.nums.push({ tok: text, v: one });
      }
      const values = ex.nums.map(n => n.v);
      const sum = Math.round(values.reduce((a, b) => a + b, 0) * 1e6) / 1e6;
      const stext = ex.nums.map(n => n.tok).join(" + ");
      let msg = values.length === 1 ? `${ex.nums[0].tok}` : `${values.length} numeri (somma ${fmtNum(sum)})`;
      if (ex.ignored.length) msg += `. Ignorat${ex.ignored.length === 1 ? "o" : "i"} ${ex.ignored.length} element${ex.ignored.length === 1 ? "o" : "i"} non numeric${ex.ignored.length === 1 ? "o" : "i"}: «${ex.ignored.slice(0, 4).join("», «")}${ex.ignored.length > 4 ? "»…" : "»"}`;
      if (layer && layer !== text) {
        const lx = SnipLib.extractNumbers(layer, fmt).nums.map(n => n.v);
        const ls = Math.round(lx.reduce((a, b) => a + b, 0) * 1e6) / 1e6;
        if (lx.length !== values.length || ls !== sum) msg += `. ATTENZIONE: il testo nascosto del PDF darebbe ${lx.length} numeri, somma ${fmtNum(ls)}: controlla`;
      }
      const row = { page: state.page, ...f, text: stext, value: sum, type: "SUM SNIP" };
      if (withCtrl) {
        // CTRL: il numero si aggiunge alla somma della cella (quella appena scritta, oppure quella su cui sei posizionato)
        const sel = await currentCellAddress();
        const target = sel;
        const existing = state.log.filter(r => r.cell === target && r.type === "SUM SNIP");
        const id = existing.length ? existing[0].id : nextSnipId();
        const all = existing.flatMap(sumRowValues).concat(values);
        const formula = SnipLib.sumFormula(all);
        await Excel.run(async ctx => {
          const c = rangeByAddress(ctx, target);
          c.formulas = [[formula]];
          c.format.fill.color = COL.sum;
          await ctx.sync();
        });
        await addLog([{ id, cell: target, ...row }]);
        state.focusId = null; state.searchMarks = [];
        drawMarks();
        state.lastCell = target;
        setStatus(`#${id}: aggiunto ${msg} alla somma di ${target}: ${formula} (${source}).`, "ok");
      } else {
        const formula = SnipLib.sumFormula(values);
        await advanceIfNeeded();
        const addr = await writeToCell(formula, COL.sum, true);
        const id = nextSnipId();
        await addLog([{ id, cell: addr, ...row }]);
        state.focusId = null; state.searchMarks = [];
        drawMarks();
        state.lastCell = addr;
        setStatus(`#${id}: somma in ${addr}: ${formula} (${source}). Per aggiungere altri numeri tieni premuto CTRL mentre trascini.`, "ok");
      }
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
      const target = sel;
      const r = await appendToCell(target, text, COL.snip);
      addr = target; appended = r;
    } else {
      await advanceIfNeeded();
      addr = await writeToCell(value, COL.snip);
    }
    const id = nextSnipId();
    await addLog([{ id, cell: addr, type: "SNIP", page: state.page, ...f, text, value }]);
    state.focusId = null; state.searchMarks = [];
    drawMarks();
    state.lastCell = addr;
    if (withCtrl) setStatus(`#${id}: «${text}» aggiunto alla cella ${addr}${appended ? " (separato da « - »)" : ""} (${source}).`, "ok");
    else setStatus(`#${id}: scritto ${isNum ? value : "«" + text + "»"} in ${addr} (${source}).`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

const fmtNum = n => (Math.round(n * 1e6) / 1e6).toLocaleString("it-IT", { maximumFractionDigits: 6 });

function renderSum() {
  const ul = $("sumlist");
  ul.innerHTML = "";
  state.sumItems.forEach((it, i) => {
    const li = h("li", { text: it.count > 1 ? `${it.count} numeri: ${it.text.length > 45 ? it.text.slice(0, 45) + "…" : it.text}  →  ${fmtNum(it.value)}` : `${it.text}  →  ${it.value}` });
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
    const values = state.sumItems.flatMap(i => i.values || [i.value]);
    const formula = SnipLib.sumFormula(values);
    const addr = await writeToCell(formula, COL.sum, true);
    const id = nextSnipId();
    await addLog(state.sumItems.map(i => ({
      id, cell: addr, type: "SUM SNIP", page: i.page, u1: i.u1, v1: i.v1, u2: i.u2, v2: i.v2, text: i.text, value: i.value
    })));
    state.sumItems = [];
    renderSum(); drawMarks();
    state.lastCell = addr;
    setStatus(`#${id}: somma scritta in ${addr}: ${formula}.`, "ok");
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

async function currentCell() {
  return Excel.run(async ctx => {
    const c = ctx.workbook.getSelectedRange().getCell(0, 0);
    c.load("address,text");
    await ctx.sync();
    return { address: c.address, text: String((c.text && c.text[0] && c.text[0][0]) || "") };
  });
}

async function handleSelection() {
  if (state.busy) return;
  if (!$("cellsearch").checked && (!$("autoopen").checked || !state.log.length)) return;
  const cell = await currentCell();
  const addr = cell.address;
  // il contenuto della cella finisce nella casella di ricerca (basta premere Invio o ↵)
  const t = cell.text.trim();
  if ($("cellsearch").checked && t && t.length <= 120 && document.activeElement !== $("q")) $("q").value = t;
  if (!$("autoopen").checked || !state.log.length) return;
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
async function writeToCell(value, color, isFormula, fontColor) {
  return Excel.run(async ctx => {
    const cell = ctx.workbook.getSelectedRange().getCell(0, 0);
    cell.load("address");
    if (isFormula) cell.formulas = [[value]]; else cell.values = [[value]];
    cell.format.fill.color = color;
    cell.format.font.color = fontColor || "#000000";
    await ctx.sync();
    return cell.address;
  });
}

// prima di un nuovo snip SENZA CTRL: se sei ancora sulla cella appena scritta, passa alla cella a destra
async function advanceIfNeeded() {
  const sel = await currentCellAddress();
  if (state.lastCell && sel === state.lastCell) return moveRight();
  return sel;
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

// =====================================================================
//  Elenco degli snip, selezione ed eliminazione
// =====================================================================
function snipGroups() {
  const m = new Map();
  state.log.forEach(r => {
    let g = m.get(r.id);
    if (!g) { g = { id: r.id, type: r.type, cell: r.cell, docId: r.docId, file: r.file, page: r.page, texts: [] }; m.set(r.id, g); }
    g.texts.push(r.text);
  });
  return [...m.values()].sort((a, b) => a.id - b.id);
}

const TAGS = { "SNIP": ["snip", "Snip"], "SUM SNIP": ["sum", "Sum"], "VALIDATION": ["val", "✓"], "EXCEPTION": ["exc", "✗"] };

function renderList() {
  const all = snipGroups();
  const q = ($("lfilter").value || "").toLowerCase().trim(), tp = $("ltype").value;
  const byDoc = new Map();
  all.forEach(g => { if (!byDoc.has(g.docId)) byDoc.set(g.docId, []); byDoc.get(g.docId).push(g); });
  const entries = state.index.docs.map(d => ({ doc: d, name: d.name, folder: d.folder, snips: byDoc.get(d.id) || [] }));
  byDoc.forEach((snips, id) => { if (!docById(id)) entries.push({ doc: null, name: snips[0].file + " (rimosso)", folder: "", snips }); });
  entries.sort((a, b) => a.name.localeCompare(b.name, "it", { numeric: true }));
  const ul = $("llist");
  ul.innerHTML = "";
  let nDocs = 0, nSnips = 0;
  entries.forEach(e => {
    let snips = e.snips;
    if (tp === "__none") { if (snips.length) return; }
    else if (tp) { snips = snips.filter(g => g.type === tp); if (!snips.length) return; }
    if (q) {
      const nameHit = (e.name + " " + e.folder).toLowerCase().includes(q);
      if (!nameHit) { snips = snips.filter(g => (g.cell + " " + g.texts.join(" ") + " #" + g.id).toLowerCase().includes(q)); if (!snips.length) return; }
    }
    nDocs++; nSnips += snips.length;
    const head = h("li", { class: "ldoc" },
      h("div", { class: "lbody" }, h("div", { class: "lname", text: (e.doc && e.doc.type !== "pdf" ? "🖼️ " : "📄 ") + e.name }),
        h("div", { class: "lmeta", text: (e.folder ? "📁 " + e.folder + " · " : "") + `${e.snips.length} snip` })));
    if (e.doc) head.addEventListener("click", () => { showTab("snip"); openDoc(e.doc.id); });
    ul.append(head);
    snips.forEach(g => {
      const [cls, lab] = TAGS[g.type] || ["snip", g.type];
      const del = h("button", { text: "🗑", title: "Elimina questo snip (definitivo)" });
      armConfirm(del, "🗑", () => deleteSnipIds([g.id]));
      del.addEventListener("click", ev => ev.stopPropagation());
      const txt = g.texts.join(" + ");
      const li = h("li", { class: "lsnip" }, h("span", { class: "lid", text: "#" + g.id }), h("span", { class: "tag " + cls, text: lab }),
        h("div", { class: "lbody" }, h("div", { text: g.cell + "  ·  " + (txt.length > 60 ? txt.slice(0, 60) + "…" : txt) }),
          h("div", { class: "lmeta", text: `pag. ${g.page}` })), del);
      li.addEventListener("click", () => goToSnip(g));
      ul.append(li);
    });
  });
  $("lsummary").textContent = entries.length ? `${nDocs} documenti · ${nSnips} snip` : "Nessun documento importato.";
}

async function selectCell(addr) {
  return Excel.run(async ctx => {
    const r = rangeByAddress(ctx, addr);
    r.worksheet.activate();
    r.select();
    await ctx.sync();
  });
}

async function goToSnip(g) {
  try {
    await selectCell(g.cell);
    const ids = [...new Set(state.log.filter(r => r.cell === g.cell).map(r => r.id))];
    await focusCell(g.cell, ids, Math.max(0, ids.indexOf(g.id)));
  } catch (e) { setStatus("Errore: " + (e.message || e), "err"); }
}

// Numeri contenuti in un rettangolo di Sum Snip (anche tanti, se è stato trascinato su una colonna)
function sumRowValues(row) {
  const fmt = $("numfmt").value, want = Number(row.value);
  for (const f of [fmt, fmt === "it" ? "en" : "it"]) {
    const v = SnipLib.extractNumbers(row.text, f).nums.map(n => n.v);
    if (v.length && Math.abs(v.reduce((a, b) => a + b, 0) - want) < 1e-6) return v;
  }
  return [want];
}

// X su un riquadro: per un Sum Snip toglie SOLO quel riquadro (e i suoi numeri) dalla somma;
// se era l'ultimo, o se non è una somma, elimina tutto lo snip come Delete Snip.
async function deleteSnipPart(row) {
  state.lastCell = null;
  const sibs = state.log.filter(r => r.id === row.id);
  if (row.type !== "SUM SNIP" || sibs.length <= 1) { await deleteSnipIds([row.id]); return; }
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  state.busy = true;
  try {
    state.log = state.log.filter(r => r !== row);
    const left = state.log.filter(r => r.id === row.id);
    const values = left.flatMap(sumRowValues);
    const formula = SnipLib.sumFormula(values);
    await Excel.run(async ctx => {
      rangeByAddress(ctx, row.cell).formulas = [[formula]];
      await ctx.sync();
    });
    await rewriteLog();
    drawMarks(); renderList(); renderOrg();
    setStatus(`Tolto ${row.text.length > 40 ? row.text.slice(0, 40) + "…" : row.text} dalla somma di ${row.cell}. Ora: ${formula}`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

async function deleteSnipIds(ids) {
  state.lastCell = null;
  if (state.busy) { setStatus("Attendi la fine dell'operazione in corso.", "err"); return; }
  state.busy = true;
  try {
    const set = new Set(ids);
    const cells = [...new Set(state.log.filter(r => set.has(r.id)).map(r => r.cell))];
    state.log = state.log.filter(r => !set.has(r.id));
    await Excel.run(async ctx => {
      for (const addr of cells) {
        const c = rangeByAddress(ctx, addr);
        const left = [...new Map(state.log.filter(r => r.cell === addr && r.type === "SNIP").map(r => [r.id, r])).values()];
        const other = state.log.some(r => r.cell === addr && r.type !== "SNIP");
        if (!left.length && !other) c.clear("All");
        else if (left.length && !other) {
          c.values = [[left.length === 1 && left[0].value !== "" && !isNaN(Number(left[0].value)) && String(Number(left[0].value)) === String(left[0].value) ? Number(left[0].value) : left.map(r => r.text).join(" - ")]];
        }
      }
      await ctx.sync();
    });
    await rewriteLog();
    state.focusId = null; state.nav = null; updateNav(); drawMarks(); renderList(); renderOrg();
    setStatus(`Eliminat${ids.length === 1 ? "o" : "i"} ${ids.length} snip (${cells.join(", ")}).`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

const colNum = c => c.split("").reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

// "Foglio1!A1:B3,Foglio1!D5" -> aree {sheet, c1, r1, c2, r2}
function parseAreas(addr) {
  const areas = [];
  String(addr).split(/,(?=(?:[^']*'[^']*')*[^']*$)/).forEach(part => {
    const m = /^(?:(.*)!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(part.trim());
    if (!m) return;
    const sheet = (m[1] || "").replace(/^'(.*)'$/, "$1").replace(/''/g, "'");
    const c1 = colNum(m[2]), r1 = +m[3], c2 = m[4] ? colNum(m[4]) : c1, r2 = m[5] ? +m[5] : r1;
    areas.push({ sheet, c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) });
  });
  return areas;
}

function cellInAreas(cellAddr, areas) {
  const a = parseAreas(cellAddr)[0];
  if (!a) return false;
  return areas.some(r => r.sheet === a.sheet && a.c1 >= r.c1 && a.c1 <= r.c2 && a.r1 >= r.r1 && a.r1 <= r.r2);
}

async function selectedAddress() {
  return Excel.run(async ctx => {
    const r = ctx.workbook.getSelectedRanges();
    r.load("address");
    await ctx.sync();
    return r.address;
  });
}

async function deleteOnSelection() {
  try {
    const addr = await selectedAddress();
    const areas = parseAreas(addr);
    const ids = [...new Set(state.log.filter(r => cellInAreas(r.cell, areas)).map(r => r.id))];
    if (!ids.length) { setStatus(`Nessuno snip nelle celle selezionate (${addr}).`, "err"); return; }
    await deleteSnipIds(ids);
  } catch (e) { setStatus("Errore: " + (e.message || e), "err"); }
}
