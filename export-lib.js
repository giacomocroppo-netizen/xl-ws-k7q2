// Crea una copia del documento con i riquadri e le etichette delle spunte disegnati sopra.
(function (root) {
  "use strict";
  const Lib = (typeof module !== "undefined" && module.exports) ? require("./snip-lib.js") : root.SnipLib;

  const clean = s => String(s).replace(/[^\x20-\x7E\xA0-\xFF]/g, "?");

  // PDFLib: la libreria pdf-lib. doc: {type:"pdf"|"image", mime}. rows: righe del registro di quel documento.
  async function annotate(PDFLib, bytes, doc, rows) {
    const { PDFDocument, StandardFonts, rgb, degrees } = PDFLib;
    let pdf;
    if (doc.type === "pdf") {
      pdf = await PDFDocument.load(bytes, { ignoreEncryption: true });
    } else {
      pdf = await PDFDocument.create();
      const img = doc.mime === "image/png" ? await pdf.embedPng(bytes) : await pdf.embedJpg(bytes);
      const pg = pdf.addPage([img.width, img.height]);
      pg.drawImage(img, { x: 0, y: 0, width: img.width, height: img.height });
    }
    const font = await pdf.embedFont(StandardFonts.HelveticaBold);
    const pages = pdf.getPages();

    for (const r of rows) {
      const page = pages[r.page - 1];
      if (!page) continue;
      const box = page.getMediaBox();
      const angle = ((page.getRotation().angle % 360) + 360) % 360;
      const dispW = (angle === 90 || angle === 270) ? box.height : box.width;
      const dispH = (angle === 90 || angle === 270) ? box.width : box.height;

      const p1 = Lib.mapPoint(r.u1, r.v1, box, angle);
      const p2 = Lib.mapPoint(r.u2, r.v2, box, angle);
      const x = Math.min(p1.x, p2.x), y = Math.min(p1.y, p2.y);
      const w = Math.abs(p2.x - p1.x), h = Math.abs(p2.y - p1.y);
      const isSum = r.type === "SUM SNIP";
      const col = isSum ? rgb(0.10, 0.44, 0.77) : rgb(0.13, 0.45, 0.27);

      page.drawRectangle({ x, y, width: w, height: h, borderColor: col, borderWidth: 1.2, color: col, opacity: 0.15, borderOpacity: 0.95 });

      // etichetta "#12 Foglio1!B5", sopra il riquadro (o sotto se non c'è spazio)
      const label = clean(`#${r.id} ${isSum ? "SUM " : ""}${r.cell}`);
      const fs = Math.max(6, Math.min(10, dispW / 80));
      const tw = font.widthOfTextAtSize(label, fs);
      const labH = fs + 3;
      const above = r.v1 * dispH - labH - 1 >= 0;
      const anchorV = above ? r.v1 - 1 / dispH : r.v2 + (labH + 1) / dispH;
      const a = Lib.mapPoint(Math.min(r.u1, r.u2), anchorV, box, angle);
      const th = angle * Math.PI / 180, cos = Math.cos(th), sin = Math.sin(th);
      page.drawRectangle({ x: a.x, y: a.y, width: tw + 4, height: labH, color: col, rotate: degrees(angle) });
      const ox = 2, oy = 2.2;
      page.drawText(label, {
        x: a.x + ox * cos - oy * sin, y: a.y + ox * sin + oy * cos,
        size: fs, font, color: rgb(1, 1, 1), rotate: degrees(angle)
      });
    }
    return pdf.save();
  }

  const api = { annotate };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.ExportLib = api;
})(typeof self !== "undefined" ? self : this);
