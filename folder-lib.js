// Gestione di cartelle e sottocartelle dei documenti (logica pura, senza interfaccia).
// index = { docs: [{id, name, folder}], folders: ["Cliente", "Cliente/2025", ...] }
(function (root) {
  "use strict";

  const cleanName = s => String(s == null ? "" : s).replace(/[\\\/]+/g, " ").replace(/\s+/g, " ").trim();
  const join = (parent, name) => { const n = cleanName(name); return parent ? parent + "/" + n : n; };
  const parentOf = p => { const i = p.lastIndexOf("/"); return i < 0 ? "" : p.slice(0, i); };
  const baseOf = p => { const i = p.lastIndexOf("/"); return i < 0 ? p : p.slice(i + 1); };
  const isInside = (p, folder) => p === folder || p.startsWith(folder + "/");
  const cleanPath = p => String(p || "").split("/").map(cleanName).filter(Boolean).join("/");
  function ancestors(p) {
    const parts = p.split("/").filter(Boolean), out = [];
    for (let i = 1; i <= parts.length; i++) out.push(parts.slice(0, i).join("/"));
    return out;
  }
  const byName = (a, b) => a.localeCompare(b, "it", { sensitivity: "base", numeric: true });

  // Rende coerente l'indice: percorsi puliti, cartelle "genitore" sempre presenti, nessun doppione.
  function normalize(index) {
    const set = new Set();
    (index.folders || []).forEach(f => ancestors(cleanPath(f)).forEach(a => set.add(a)));
    (index.docs || []).forEach(d => {
      d.folder = cleanPath(d.folder);
      ancestors(d.folder).forEach(a => set.add(a));
    });
    index.folders = [...set].sort(byName);
    return index;
  }

  function addFolder(index, parent, name) {
    const n = cleanName(name);
    if (!n) return null;
    const p = join(parent, n);
    if (!index.folders.includes(p)) { index.folders.push(p); normalize(index); }
    return p;
  }

  // Cambia il prefisso di tutto ciò che sta dentro "oldPath" (cartelle e documenti).
  function rebase(index, oldPath, newPath) {
    const re = f => isInside(f, oldPath) ? newPath + f.slice(oldPath.length) : f;
    index.folders = index.folders.map(re);
    index.docs.forEach(d => { d.folder = re(d.folder); });
    normalize(index);
  }

  function renameFolder(index, path, newName) {
    const n = cleanName(newName);
    if (!n || !index.folders.includes(path)) return null;
    const np = join(parentOf(path), n);
    if (np === path) return path;
    if (index.folders.includes(np)) return null;           // esiste già una cartella con quel nome
    rebase(index, path, np);
    return np;
  }

  function moveFolder(index, path, newParent) {
    if (!index.folders.includes(path)) return null;
    if (newParent && !index.folders.includes(newParent)) return null;
    if (isInside(newParent, path)) return null;            // non si può spostare una cartella dentro se stessa
    const np = join(newParent, baseOf(path));
    if (np === path) return path;
    if (index.folders.includes(np)) return null;
    rebase(index, path, np);
    return np;
  }

  // Elimina la cartella: il suo contenuto passa alla cartella superiore.
  function deleteFolder(index, path) {
    if (!index.folders.includes(path)) return false;
    const par = parentOf(path);
    const fix = f => {                                    // "A/B/C" eliminando "A/B" diventa "A/C"
      const rest = f.slice(path.length + 1);
      return par ? par + "/" + rest : rest;
    };
    index.folders = index.folders.map(f => f === path ? null : (f.startsWith(path + "/") ? fix(f) : f)).filter(Boolean);
    index.docs.forEach(d => {
      if (d.folder === path) d.folder = par;
      else if (d.folder.startsWith(path + "/")) d.folder = fix(d.folder);
    });
    normalize(index);
    return true;
  }

  function moveDocs(index, ids, folder) {
    const set = new Set(ids);
    index.docs.forEach(d => { if (set.has(d.id)) d.folder = cleanPath(folder); });
    normalize(index);
  }

  // Albero: { path, name, children:[...], docs:[...] }
  function buildTree(index) {
    const rootNode = { path: "", name: "", children: [], docs: [] };
    const nodes = { "": rootNode };
    [...index.folders].sort(byName).forEach(f => {
      const node = { path: f, name: baseOf(f), children: [], docs: [] };
      nodes[f] = node;
      (nodes[parentOf(f)] || rootNode).children.push(node);
    });
    index.docs.forEach(d => (nodes[d.folder] || rootNode).docs.push(d));
    const sortRec = n => {
      n.children.sort((a, b) => byName(a.name, b.name));
      n.docs.sort((a, b) => byName(a.name, b.name));
      n.children.forEach(sortRec);
    };
    sortRec(rootNode);
    return rootNode;
  }

  function countDocs(node) {
    return node.docs.length + node.children.reduce((a, c) => a + countDocs(c), 0);
  }

  // Elenco per i menu a tendina: [{path, label}]
  function options(index) {
    const out = [{ path: "", label: "📂 (cartella principale)" }];
    [...index.folders].sort(byName).forEach(f => {
      const depth = f.split("/").length - 1;
      out.push({ path: f, label: "  ".repeat(depth) + "📁 " + baseOf(f) });
    });
    return out;
  }

  const api = { cleanName, cleanPath, join, parentOf, baseOf, isInside, normalize, addFolder, renameFolder, moveFolder, deleteFolder, moveDocs, buildTree, countDocs, options };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.FolderLib = api;
})(typeof self !== "undefined" ? self : this);
