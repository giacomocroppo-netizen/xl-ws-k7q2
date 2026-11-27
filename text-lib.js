// Ricerca di parole nel testo delle pagine (logica pura).
// Una pagina è un elenco di pezzi di testo: {s, u1, v1, u2, v2}  (coordinate in frazioni 0..1 della pagina)
(function (root) {
  "use strict";

  // un carattere -> un carattere (minuscolo, senza accenti): così le posizioni restano allineate
  function normChar(c) {
    const base = c.normalize("NFD")[0] || c;
    return (base.toLowerCase()[0]) || base;
  }
  function norm(s) {
    let o = "";
    for (let i = 0; i < s.length; i++) o += normChar(s[i]);
    return o;
  }
  const squash = s => String(s).replace(/\s+/g, " ").trim();

  function buildPage(items) {
    const spans = [];
    let text = "";
    items.forEach(it => {
      const s = squash(it.s);
      if (!s) return;
      if (text) text += " ";
      spans.push({ a: text.length, b: text.length + s.length, it, s });
      text += s;
    });
    return { text, norm: norm(text), spans };
  }

  // Cerca "query" in una pagina. Ritorna [{pos, len, snippet:{before,match,after}, boxes:[{u1,v1,u2,v2}]}]
  function search(page, query, max) {
    const q = norm(squash(query));
    if (!q) return [];
    const hits = [];
    let from = 0;
    while (hits.length < (max || 50)) {
      const i = page.norm.indexOf(q, from);
      if (i < 0) break;
      const j = i + q.length;
      const boxes = [];
      for (const sp of page.spans) {
        if (sp.b <= i || sp.a >= j) continue;
        const a = Math.max(i, sp.a) - sp.a, b = Math.min(j, sp.b) - sp.a, n = sp.s.length;
        const w = sp.it.u2 - sp.it.u1;
        boxes.push({ u1: sp.it.u1 + (a / n) * w, v1: sp.it.v1, u2: sp.it.u1 + (b / n) * w, v2: sp.it.v2 });
      }
      const s0 = Math.max(0, i - 35), s1 = Math.min(page.text.length, j + 35);
      hits.push({
        pos: i, len: q.length, boxes,
        snippet: { before: (s0 > 0 ? "…" : "") + page.text.slice(s0, i), match: page.text.slice(i, j), after: page.text.slice(j, s1) + (s1 < page.text.length ? "…" : "") }
      });
      from = j;
    }
    return hits;
  }

  const api = { norm, buildPage, search };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TextLib = api;
})(typeof self !== "undefined" ? self : this);
