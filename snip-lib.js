// Logica "pura" (senza interfaccia): lettura dei numeri e estrazione del testo da un'area.
(function (root) {
  "use strict";

  // Converte un testo in numero. fmt = "it" (1.234,56) oppure "en" (1,234.56). Ritorna null se non è un numero.
  function parseNumber(raw, fmt) {
    if (raw == null) return null;
    let s = String(raw).trim();
    if (!s) return null;
    s = s.replace(/[  \s]/g, "");           // spazi
    s = s.replace(/[€$£]|EUR|USD/gi, "");               // valute
    if (/^[–—−]$/.test(s) || s === "-") return 0;       // trattino = zero (tipico nei bilanci)
    s = s.replace(/[–—−]/g, "-");                       // trattini lunghi -> meno

    let neg = false;
    if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1); }   // (1.234) = negativo
    if (/-$/.test(s)) { neg = !neg; s = s.slice(0, -1); }         // 1.234- = negativo
    if (/^-/.test(s)) { neg = !neg; s = s.slice(1); }
    if (/^\+/.test(s)) s = s.slice(1);
    s = s.replace(/%$/, "");

    if (!/^[0-9.,]+$/.test(s) || !/[0-9]/.test(s)) return null;

    const dec = fmt === "en" ? "." : ",";
    const thou = fmt === "en" ? "," : ".";
    const lastDec = s.lastIndexOf(dec);
    const lastThou = s.lastIndexOf(thou);

    let intPart, fracPart = "";
    if (lastDec !== -1) {
      // Il separatore decimale deve comparire una sola volta e dopo gli eventuali separatori di migliaia
      if (s.indexOf(dec) !== lastDec) return null;
      if (lastThou > lastDec) return null;
      intPart = s.slice(0, lastDec);
      fracPart = s.slice(lastDec + 1);
      if (/[^0-9]/.test(fracPart)) return null;
      intPart = intPart.split(thou).join("");
    } else if (lastThou !== -1) {
      // Solo separatori "di migliaia": validi se a gruppi di 3 cifre; altrimenti lo leggiamo come decimale "all'americana"
      if (new RegExp("^\\d{1,3}(\\" + thou + "\\d{3})+$").test(s)) {
        intPart = s.split(thou).join("");
      } else if (s.split(thou).length === 2) {
        intPart = s.split(thou)[0];
        fracPart = s.split(thou)[1];
      } else return null;
    } else {
      intPart = s;
    }
    if (/[^0-9]/.test(intPart)) return null;
    const n = parseFloat((intPart || "0") + (fracPart ? "." + fracPart : ""));
    if (!isFinite(n)) return null;
    return neg ? -n : n;
  }

  // items: elenco di pezzi di testo già in coordinate "schermo": {str, x, y (linea di base), w, h}
  // rect: {x1,y1,x2,y2}. Ritorna il testo dentro l'area (le righe sono unite da spazio).
  function textInRect(items, rect) {
    const picked = [];
    for (const it of items) {
      if (!it.str || !it.str.trim()) continue;
      // il centro della riga di testo deve cadere dentro l'area: così le righe vicine (sopra/sotto) restano fuori
      const cy = it.y - it.h * 0.4;
      if (cy < rect.y1 || cy > rect.y2) continue;
      const left = it.x, right = it.x + it.w;
      if (right < rect.x1 || left > rect.x2) continue;
      // Se il blocco è solo in parte dentro l'area, tagliamo i caratteri in proporzione alla larghezza
      let s = it.str;
      if (left < rect.x1 || right > rect.x2) {
        const n = s.length, cw = it.w / n;
        let a = 0, b = n;
        for (let i = 0; i < n; i++) {
          const cx = left + cw * (i + 0.5);
          if (cx < rect.x1) a = i + 1;
          if (cx > rect.x2 && b === n) b = i;
        }
        s = s.slice(a, b);
      }
      if (s.trim()) picked.push({ s, x: it.x, y: it.y, h: it.h });
    }
    picked.sort((p, q) => (Math.abs(p.y - q.y) > Math.min(p.h, q.h) * 0.5) ? p.y - q.y : p.x - q.x);
    // ricomponi: stessa riga -> spazio, riga diversa -> spazio (poi compattiamo)
    return cleanRead(picked.map(p => p.s.trim()).join(" "));
  }

  // toglie dal testo letto le linee verticali delle tabelle ( | ¦ │ ) e compatta gli spazi
  function cleanRead(s) {
    return String(s || "").replace(/[|¦│┃‖]+/g, " ").replace(/\s+/g, " ").trim();
  }

  // Formula di somma "tracciabile": =1234.5+200-30
  function sumFormula(values) {
    const fmt = v => String(Math.abs(v));
    let f = "=";
    values.forEach((v, i) => {
      if (i === 0) f += (v < 0 ? "-" : "") + fmt(v);
      else f += (v < 0 ? "-" : "+") + fmt(v);
    });
    return f;
  }

  // ---- Utilità per incorporare i file nell'Excel ----
  function bytesToB64(bytes) {
    let out = "";
    const step = 0x8000;
    for (let i = 0; i < bytes.length; i += step) {
      out += String.fromCharCode.apply(null, bytes.subarray(i, i + step));
    }
    return btoa(out);
  }
  function b64ToBytes(b64) {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }
  function chunkString(s, n) {
    const out = [];
    for (let i = 0; i < s.length; i += n) out.push(s.slice(i, i + n));
    return out.length ? out : [""];
  }
  function escapeXml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  }
  function unescapeXml(s) {
    return String(s).replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&amp;/g, "&");
  }

  // Punto (u,v) della pagina "come la vedi" (u verso destra, v verso il basso, valori 0..1)
  // -> coordinate PDF (origine in basso a sinistra), tenendo conto della rotazione della pagina.
  function mapPoint(u, v, box, angle) {
    const a = ((angle % 360) + 360) % 360;
    const W = box.width, H = box.height;
    if (a === 90) return { x: box.x + v * W, y: box.y + u * H };
    if (a === 180) return { x: box.x + (1 - u) * W, y: box.y + v * H };
    if (a === 270) return { x: box.x + (1 - v) * W, y: box.y + (1 - u) * H };
    return { x: box.x + u * W, y: box.y + (1 - v) * H };
  }

  const api = { parseNumber, textInRect, cleanRead, sumFormula, bytesToB64, b64ToBytes, chunkString, escapeXml, unescapeXml, mapPoint };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.SnipLib = api;
})(typeof self !== "undefined" ? self : this);
