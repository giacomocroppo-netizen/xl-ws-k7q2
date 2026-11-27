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

  function mkHit(page, i, j) {
    const boxes = [];
    for (const sp of page.spans) {
      if (sp.b <= i || sp.a >= j) continue;
      const a = Math.max(i, sp.a) - sp.a, b = Math.min(j, sp.b) - sp.a, n = sp.s.length;
      const w = sp.it.u2 - sp.it.u1;
      boxes.push({ u1: sp.it.u1 + (a / n) * w, v1: sp.it.v1, u2: sp.it.u1 + (b / n) * w, v2: sp.it.v2 });
    }
    const s0 = Math.max(0, i - 35), s1 = Math.min(page.text.length, j + 35);
    return {
      pos: i, len: j - i, boxes,
      snippet: { before: (s0 > 0 ? "…" : "") + page.text.slice(s0, i), match: page.text.slice(i, j), after: page.text.slice(j, s1) + (s1 < page.text.length ? "…" : "") }
    };
  }

  // Forma "canonica" di un importo: 202.330 = 202330 = 202.330,00 -> "202330";  1.234,5 = 1234,50 -> "1234.5"
  function canon(tok) {
    const t = String(tok).replace(/^[^\d]+|[^\d]+$/g, "").replace(/['\s]/g, "");
    if (!t) return "";
    const m = /^(.*?)[.,](\d{1,2})$/.exec(t);
    const ip = (m ? m[1] : t).replace(/[.,]/g, "").replace(/^0+(?=\d)/, "");
    const dp = m ? m[2].replace(/0+$/, "") : "";
    return ip + (dp ? "." + dp : "");
  }
  // la query è un importo? (1234, 1.234, 1.234,56, 1234,5 ... ma non una data come 03.11.2026)
  function isAmount(q) {
    const s = String(q).trim().replace(/^[-+(]\s*/, "").replace(/\s*\)$/, "");
    return /^\d+([.,]\d{1,2})?$/.test(s) || /^\d{1,3}([.,']\d{3})+([.,]\d{1,2})?$/.test(s);
  }

  // Cerca "query" in una pagina. Ritorna [{pos, len, snippet:{before,match,after}, boxes:[{u1,v1,u2,v2}]}]
  function search(page, query, max) {
    const cap = max || 50;
    const q = norm(squash(query));
    if (!q) return [];
    const hits = [];
    let from = 0;
    while (hits.length < cap) {
      const i = page.norm.indexOf(q, from);
      if (i < 0) break;
      hits.push(mkHit(page, i, i + q.length));
      from = i + q.length;
    }
    if (isAmount(query)) {
      const want = canon(query);
      const re = /\d[\d.,']*\d|\d/g;
      let m;
      while ((m = re.exec(page.text)) && hits.length < cap * 2) {
        if (canon(m[0]) !== want) continue;
        if (hits.some(h => h.pos <= m.index && m.index < h.pos + h.len)) continue;
        hits.push(mkHit(page, m.index, m.index + m[0].length));
      }
      hits.sort((a, b) => a.pos - b.pos);
    }
    return hits.slice(0, cap);
  }

  const api = { norm, buildPage, search, canon, isAmount };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TextLib = api;
})(typeof self !== "undefined" ? self : this);
