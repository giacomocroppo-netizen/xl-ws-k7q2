/* global Office, Excel, pdfjsLib, Tesseract, PDFLib, SnipLib, ExportLib */
"use strict";

const PDFJS_VERSION = "3.11.174";
const LOG_SHEET = "_SNIP_LOG";
const NS_INDEX = "urn:snipper:index:v1";   // elenco dei documenti incorporati
const NS_DOC = "urn:snipper:doc:v1";       // contenuto dei documenti (a blocchi)
const CHUNK = 2000000;                     // caratteri per blocco
const WARN_MB = 15;

const state = {
  mode: "snip",
  index: { docs: [] },     // documenti incorporati
  log: [],                 // snip già fatti (copia del foglio nascosto _SNIP_LOG)
  docId: null,
  pdf: null, image: null,
  page: 1, numPages: 1,
  scale: 1.6, dpr: 1,
  viewW: 0, viewH: 0,
  textItems: [], hasTextLayer: false,
  sumItems: [],            // snip in attesa di somma
  focusId: null,           // numero snip evidenziato
  busy: false
};

const $ = id => document.getElementById(id);
const canvas = $("canvas");
const overlay = $("overlay");
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);

function setStatus(msg, kind) {
  const el = $("status");
  el.textContent = msg;
  el.className = "status" + (kind ? " " + kind : "");
}
const fmtSize = b => b >= 1048576 ? (b / 1048576).toFixed(1).replace(".", ",") + " MB" : Math.max(1, Math.round(b / 1024)) + " KB";
const curDoc = () => state.index.docs.find(d => d.id === state.docId);

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
      : "Nessun documento in questo file. Premi «➕ Importa».");
  } catch (e) {
    console.error(e);
    setStatus("Errore all'avvio: " + (e.message || e), "err");
  }
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
  $("import").addEventListener("change", onImport);
  $("tab-snip").addEventListener("click", () => showTab("snip"));
  $("tab-org").addEventListener("click", () => { showTab("org"); renderOrg(); });
  $("docselect").addEventListener("change", e => e.target.value && openDoc(e.target.value));
  $("mode-snip").addEventListener("click", () => setMode("snip"));
  $("mode-sum").addEventListener("click", () => setMode("sum"));
  $("prev").addEventListener("click", () => gotoPage(state.page - 1));
  $("next").addEventListener("click", () => gotoPage(state.page + 1));
  $("zoomin").addEventListener("click", () => zoom(1.2));
  $("zoomout").addEventListener("click", () => zoom(1 / 1.2));
  $("findsnip").addEventListener("click", findSnip);
  $("sumwrite").addEventListener("click", writeSum);
  $("sumclear").addEventListener("click", () => { state.sumItems = []; renderSum(); drawMarks(); });
  $("docfilter").addEventListener("input", renderOrg);
  setupSelection();
}

function showTab(t) {
  $("view-snip").hidden = t !== "snip";
  $("view-org").hidden = t !== "org";
  $("tab-snip").classList.toggle("active", t === "snip");
  $("tab-org").classList.toggle("active", t === "org");
}

function setMode(m) {
  state.mode = m;
  $("mode-snip").classList.toggle("active", m === "snip");
  $("mode-sum").classList.toggle("active", m === "sum");
  $("sumpanel").hidden = m !== "sum";
  setStatus(m === "snip"
    ? "SNIP: trascina un rettangolo sul valore; finisce nella cella selezionata."
    : "SUM SNIP: seleziona i numeri uno alla volta, poi premi «Scrivi somma in cella».");
}

function refreshAll() {
  renderDocSelect();
  renderOrg();
}

// =====================================================================
//  Documenti incorporati nel file Excel (parti XML personalizzate)
// =====================================================================
async function readIndex() {
  return Excel.run(async ctx => {
    const col = ctx.workbook.customXmlParts.getByNamespace(NS_INDEX);
    col.load("items/id");
    await ctx.sync();
    if (!col.items.length) return { docs: [] };
    const xml = col.items[0].getXml();
    await ctx.sync();
    try {
      const d = new DOMParser().parseFromString(xml.value, "application/xml");
      const idx = JSON.parse(d.documentElement.textContent);
      idx.docs = idx.docs || [];
      return idx;
    } catch (e) { return { docs: [] }; }
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
      for (let i = 1; i < col.items.length; i++) col.items[i].delete();   // eventuali doppioni
    } else {
      ctx.workbook.customXmlParts.add(xml);
    }
    await ctx.sync();
  });
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

async function loadBytes(doc) {
  let b64 = "";
  await Excel.run(async ctx => {
    for (const id of doc.parts) {
      const x = ctx.workbook.customXmlParts.getItem(id).getXml();
      await ctx.sync();
      const m = /<c[^>]*>([\s\S]*)<\/c>/.exec(x.value);
      b64 += m ? m[1] : "";
    }
  });
  return SnipLib.b64ToBytes(b64);
}

async function deleteParts(ids) {
  return Excel.run(async ctx => {
    ids.forEach(id => ctx.workbook.customXmlParts.getItem(id).delete());
    await ctx.sync();
  });
}

// =====================================================================
//  Importazione
// =====================================================================
async function onImport(ev) {
  const files = Array.from(ev.target.files || []);
  ev.target.value = "";
  if (!files.length || state.busy) return;
  state.busy = true;
  let lastId = null, added = 0;
  try {
    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      setStatus(`Importo ${f.name} (${i + 1}/${files.length})… non chiudere Excel.`);
      if (state.index.docs.some(d => d.name === f.name && d.size === f.size)) {
        lastId = state.index.docs.find(d => d.name === f.name && d.size === f.size).id;
        setStatus(`«${f.name}» è già nel file: l'ho riaperto.`);
        continue;
      }
      const isPdf = /\.pdf$/i.test(f.name) || f.type === "application/pdf";
      const isPng = /\.png$/i.test(f.name) || f.type === "image/png";
      const bytes = new Uint8Array(await f.arrayBuffer());
      let pages = 1;
      if (isPdf) {
        const test = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;   // controlla che sia valido
        pages = test.numPages;
        test.destroy();
      }
      const parts = await storeBytes(bytes);
      const doc = {
        id: uid(), name: f.name, type: isPdf ? "pdf" : "image",
        mime: isPdf ? "application/pdf" : (isPng ? "image/png" : "image/jpeg"),
        size: f.size, pages, folder: "", parts, added: new Date().toISOString()
      };
      state.index.docs.push(doc);
      await writeIndex();
      lastId = doc.id;
      added++;
    }
    refreshAll();
    const big = files.some(f => f.size > WARN_MB * 1048576);
    if (lastId) {
      showTab("snip");
      await openDoc(lastId);
    }
    setStatus(`Importati ${added} documento/i. Ricorda di salvare il file Excel.` +
      (big ? " Attenzione: file grandi appesantiscono molto la cartella di lavoro." : ""), big ? "err" : "ok");
  } catch (e) {
    console.error(e);
    setStatus("Importazione non riuscita: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

// =====================================================================
//  Schermata ORGANIZZA
// =====================================================================
function renderDocSelect() {
  const sel = $("docselect");
  sel.innerHTML = "";
  if (!state.index.docs.length) {
    sel.appendChild(new Option("(nessun documento)", ""));
  } else {
    state.index.docs.forEach(d => sel.appendChild(new Option(d.name, d.id)));
    if (state.docId) sel.value = state.docId;
  }
}

function snipCount(docId) {
  return new Set(state.log.filter(r => r.docId === docId).map(r => r.id)).size;
}

function renderOrg() {
  const q = ($("docfilter").value || "").toLowerCase();
  const docs = state.index.docs.filter(d => !q || d.name.toLowerCase().includes(q));
  const total = state.index.docs.reduce((a, d) => a + d.size, 0);
  $("orgsummary").textContent = state.index.docs.length
    ? `${state.index.docs.length} documenti · ${fmtSize(total)} incorporati nel file`
    : "Nessun documento importato.";

  const folders = [...new Set(state.index.docs.map(d => d.folder).filter(Boolean))].sort();
  const dl = $("folders");
  dl.innerHTML = "";
  folders.forEach(f => dl.appendChild(new Option(f)));

  const list = $("doclist");
  list.innerHTML = "";
  const groups = {};
  docs.forEach(d => (groups[d.folder || ""] = groups[d.folder || ""] || []).push(d));
  Object.keys(groups).sort((a, b) => (a === "" ? 1 : b === "" ? -1 : a.localeCompare(b))).forEach(f => {
    const h = document.createElement("div");
    h.className = "folderhead";
    h.textContent = f ? "📁 " + f : "Senza cartella";
    list.appendChild(h);
    groups[f].sort((a, b) => a.name.localeCompare(b.name)).forEach(d => list.appendChild(docCard(d)));
  });
}

function docCard(d) {
  const el = document.createElement("div");
  el.className = "doc";
  const n = snipCount(d.id);
  el.innerHTML = `<div class="docname"></div>
    <div class="meta"><span></span> · Cartella: <input list="folders" placeholder="nessuna" /></div>
    <div class="actions"></div>`;
  el.querySelector(".docname").textContent = (d.type === "pdf" ? "📄 " : "🖼️ ") + d.name;
  el.querySelector(".meta span").textContent = `${d.pages} pag. · ${fmtSize(d.size)} · ${n} snip`;
  const inp = el.querySelector("input");
  inp.value = d.folder || "";
  inp.addEventListener("change", async () => {
    d.folder = inp.value.trim();
    try { await writeIndex(); renderOrg(); } catch (e) { setStatus("Errore: " + (e.message || e), "err"); }
  });

  const act = el.querySelector(".actions");
  const mk = (txt, fn, cls) => {
    const b = document.createElement("button");
    b.textContent = txt;
    if (cls) b.className = cls;
    b.addEventListener("click", () => fn(b));
    act.appendChild(b);
    return b;
  };
  mk("Apri", () => { showTab("snip"); openDoc(d.id); });
  mk("⬇ Originale", () => downloadOriginal(d));
  mk(`⬇ Con spunte${n ? " (" + n + ")" : ""}`, () => downloadAnnotated(d));
  mk("🗑 Elimina", b => {
    if (!b.classList.contains("armed")) {
      b.classList.add("armed");
      b.textContent = "Sicuro? Clicca di nuovo";
      setTimeout(() => { b.classList.remove("armed"); b.textContent = "🗑 Elimina"; }, 4000);
    } else {
      removeDoc(d);
    }
  }, "danger");
  return el;
}

async function removeDoc(d) {
  if (state.busy) return;
  state.busy = true;
  try {
    setStatus("Elimino…");
    await deleteParts(d.parts);
    state.index.docs = state.index.docs.filter(x => x.id !== d.id);
    await writeIndex();
    const had = state.log.some(r => r.docId === d.id);
    if (had) {
      state.log = state.log.filter(r => r.docId !== d.id);
      await rewriteLog();
    }
    if (state.docId === d.id) closeDoc();
    refreshAll();
    setStatus(`«${d.name}» eliminato.` + (had ? " Le celle già spuntate restano in Excel ma non sono più collegate al documento." : ""), "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
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
  state.docId = null; state.pdf = null; state.image = null;
  canvas.width = canvas.height = 0;
  canvas.style.width = canvas.style.height = "0px";
  $("stage").style.display = "none";
  $("empty").style.display = "";
  $("pageinfo").textContent = "–";
  drawMarks();
}

async function openDoc(id) {
  const d = state.index.docs.find(x => x.id === id);
  if (!d) return;
  try {
    setStatus(`Apro «${d.name}»…`);
    const bytes = await loadBytes(d);
    state.pdf = null; state.image = null;
    if (d.type === "pdf") {
      state.pdf = await pdfjsLib.getDocument({ data: bytes.slice() }).promise;
      state.numPages = state.pdf.numPages;
    } else {
      state.image = await new Promise((res, rej) => {
        const img = new Image();
        img.onload = () => res(img);
        img.onerror = () => rej(new Error("immagine non valida"));
        img.src = URL.createObjectURL(new Blob([bytes], { type: d.mime }));
      });
      state.numPages = 1;
    }
    state.docId = id;
    state.page = 1;
    state.sumItems = []; renderSum();
    $("docselect").value = id;
    $("stage").style.display = "";
    $("empty").style.display = "none";
    await renderPage();
    setStatus("Documento aperto. Trascina un rettangolo sul valore da spuntare.", "ok");
  } catch (e) {
    console.error(e);
    setStatus("Impossibile aprire il documento: " + (e.message || e), "err");
  }
}

async function renderPage() {
  const dpr = state.dpr = window.devicePixelRatio || 1;
  const ctx = canvas.getContext("2d");
  state.textItems = [];
  state.hasTextLayer = false;

  if (state.pdf) {
    const page = await state.pdf.getPage(state.page);
    const viewport = page.getViewport({ scale: state.scale });
    state.viewW = viewport.width; state.viewH = viewport.height;
    canvas.width = Math.floor(viewport.width * dpr);
    canvas.height = Math.floor(viewport.height * dpr);
    canvas.style.width = viewport.width + "px";
    canvas.style.height = viewport.height + "px";
    await page.render({ canvasContext: ctx, viewport: page.getViewport({ scale: state.scale * dpr }) }).promise;
    const tc = await page.getTextContent();
    state.textItems = tc.items.map(it => {
      const t = pdfjsLib.Util.transform(viewport.transform, it.transform);
      return { str: it.str, x: t[4], y: t[5], w: it.width * state.scale, h: Math.hypot(t[2], t[3]) || it.height * state.scale };
    });
    state.hasTextLayer = state.textItems.some(i => i.str && i.str.trim());
  } else if (state.image) {
    const base = Math.min(state.image.naturalWidth, 700) * state.scale / 1.6;
    const w = Math.round(base);
    const h = Math.round(base * state.image.naturalHeight / state.image.naturalWidth);
    state.viewW = w; state.viewH = h;
    canvas.width = Math.floor(w * dpr);
    canvas.height = Math.floor(h * dpr);
    canvas.style.width = w + "px";
    canvas.style.height = h + "px";
    ctx.drawImage(state.image, 0, 0, canvas.width, canvas.height);
  }
  $("pageinfo").textContent = state.docId ? `${state.page} / ${state.numPages}` : "–";
  drawMarks();
}

async function gotoPage(n) {
  if (!state.docId || n < 1 || n > state.numPages) return;
  state.page = n;
  await renderPage();
}

async function zoom(f) {
  state.scale = Math.min(4, Math.max(0.6, state.scale * f));
  if (state.docId) await renderPage();
}

// Riquadri già fatti (dal registro) + quelli in attesa di somma
function drawMarks() {
  overlay.querySelectorAll(".mark").forEach(n => n.remove());
  if (!state.docId) return;
  const W = state.viewW, H = state.viewH;
  const add = (r, cls, label) => {
    const d = document.createElement("div");
    d.className = "mark " + cls;
    Object.assign(d.style, {
      left: Math.min(r.u1, r.u2) * W + "px", top: Math.min(r.v1, r.v2) * H + "px",
      width: Math.abs(r.u2 - r.u1) * W + "px", height: Math.abs(r.v2 - r.v1) * H + "px"
    });
    if (label) { const b = document.createElement("b"); b.textContent = label; d.appendChild(b); }
    overlay.appendChild(d);
  };
  state.log.filter(r => r.docId === state.docId && r.page === state.page)
    .forEach(r => add(r, (r.type === "SUM SNIP" ? "sum " : "") + (r.id === state.focusId ? "focus" : ""), "#" + r.id));
  state.sumItems.filter(s => s.page === state.page).forEach(s => add(s, "pending", "+"));
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
    const p = pos(e);
    start = p;
    box = document.createElement("div");
    box.className = "sel";
    overlay.appendChild(box);
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
    const f = {
      u1: Math.min(start.x, p.x) / p.w, v1: Math.min(start.y, p.y) / p.h,
      u2: Math.max(start.x, p.x) / p.w, v2: Math.max(start.y, p.y) / p.h
    };
    const tiny = (f.u2 - f.u1) * p.w < 4 || (f.v2 - f.v1) * p.h < 4;
    start = null;
    if (box) { box.remove(); box = null; }
    if (tiny) return;
    await onSnip(f);
  });
}

// =====================================================================
//  Lettura dell'area (testo del PDF, altrimenti OCR)
// =====================================================================
async function readArea(f) {
  if (state.pdf && state.hasTextLayer) {
    const t = SnipLib.textInRect(state.textItems, {
      x1: f.u1 * state.viewW, y1: f.v1 * state.viewH, x2: f.u2 * state.viewW, y2: f.v2 * state.viewH
    });
    if (t) return { text: t, source: "testo PDF" };
  }
  return { text: await ocrArea(f), source: "OCR" };
}

async function ocrArea(f) {
  setStatus("Lettura con OCR… (la prima volta scarica i dati della lingua, ci vuole qualche secondo)");
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
async function onSnip(f) {
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
    const addr = await writeToCell(value, "#C6EFCE");
    const id = nextSnipId();
    await addLog([{ id, cell: addr, type: "SNIP", page: state.page, ...f, text, value }]);
    state.focusId = null;
    drawMarks();
    setStatus(`#${id}: scritto ${isNum ? value : "«" + text + "»"} in ${addr} (${source}).`, "ok");
    if ($("movedown").checked) await moveDown();
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
    const li = document.createElement("li");
    li.textContent = `${it.text}  →  ${it.value}`;
    const b = document.createElement("button");
    b.textContent = "✕";
    b.title = "Rimuovi";
    b.addEventListener("click", () => { state.sumItems.splice(i, 1); renderSum(); drawMarks(); });
    li.appendChild(b);
    ul.appendChild(li);
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
    setStatus(`#${id}: somma scritta in ${addr}: ${formula}`, "ok");
    if ($("movedown").checked) await moveDown();
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  } finally {
    state.busy = false;
  }
}

// =====================================================================
//  Trova lo snip della cella selezionata
// =====================================================================
async function findSnip() {
  try {
    const addr = await Excel.run(async ctx => {
      const c = ctx.workbook.getSelectedRange().getCell(0, 0);
      c.load("address");
      await ctx.sync();
      return c.address;
    });
    const rows = state.log.filter(r => r.cell === addr);
    if (!rows.length) { setStatus(`Nessuno snip sulla cella ${addr}.`, "err"); return; }
    const first = rows[0];
    if (!state.index.docs.some(d => d.id === first.docId)) { setStatus("Il documento di questo snip non è più nel file.", "err"); return; }
    showTab("snip");
    if (state.docId !== first.docId) await openDoc(first.docId);
    state.focusId = first.id;
    state.page = first.page;
    await renderPage();
    setStatus(`Snip #${first.id} di ${addr}: pagina ${first.page}, «${first.text}».`, "ok");
  } catch (e) {
    console.error(e);
    setStatus("Errore: " + (e.message || e), "err");
  }
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

async function moveDown() {
  return Excel.run(async ctx => {
    ctx.workbook.getSelectedRange().getCell(0, 0).getOffsetRange(1, 0).select();
    await ctx.sync();
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
    const h = sh.getRange("A1:N1");
    h.numberFormat = [LOG_FORMAT.map(() => "@")];
    h.values = [LOG_HEADER];
    h.format.font.bold = true;
    sh.visibility = Excel.SheetVisibility.hidden;
    await ctx.sync();
  }
  return sh;
}

async function addLog(partial) {
  const doc = curDoc();
  const when = new Date().toLocaleString("it-IT");
  const rows = partial.map(p => ({
    when, docId: state.docId, file: doc ? doc.name : "", fmt: $("numfmt").value, ...p
  }));
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
