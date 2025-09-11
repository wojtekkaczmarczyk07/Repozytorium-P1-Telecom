(() => {
  'use strict';

  const NIP_RE = /\b\d{10}\b/;
  const ROW_Y_TOL = 22; // px tolerancja porównania po osi Y (wspólna linia)

  // Pobierz wszystkie roots: document + zagnieżdżone shadowRoot-y
  function collectRoots(root) {
    const roots = [root];
    const stack = [root];
    while (stack.length) {
      const r = stack.pop();
      const walker = (r.body || r).querySelectorAll ? (r.body || r).querySelectorAll('*') : [];
      for (const el of walker) {
        if (el.shadowRoot) {
          roots.push(el.shadowRoot);
          stack.push(el.shadowRoot);
        }
      }
    }
    return roots;
  }

  function textHasDB(el) {
    try {
      const t = (el.textContent || '').trim();
      if (/\bDB\b/.test(t)) return true;
      const b = getComputedStyle(el, '::before').content || '';
      const a = getComputedStyle(el, '::after').content || '';
      const norm = (s) => (s||'').replace(/['"]/g,'').trim();
      return /\bDB\b/.test(norm(b)) || /\bDB\b/.test(norm(a));
    } catch(e){ return false; }
  }

  // Zbierz kandydatów DB (elementy z tekstem/pseudo „DB”) + ich prostokąty
  function collectDBRects(root) {
    const rects = [];
    const nodes = (root.querySelectorAll ? root.querySelectorAll('button, [role="button"], .btn, .badge, span, div') : []);
    for (const el of nodes) {
      if (!textHasDB(el)) continue;
      const r = el.getBoundingClientRect();
      if (r && r.width > 0 && r.height > 0) {
        rects.push({y: r.top + r.height/2, el});
      }
    }
    return rects;
  }

  // Zbierz NIP y-centers + tekst NIP (z nodeValue)
  function collectNIPRects(root) {
    const out = [];
    try {
      const walker = (root.ownerDocument || root).createTreeWalker(root, NodeFilter.SHOW_TEXT);
      let node;
      while ((node = walker.nextNode())) {
        const txt = node.nodeValue || '';
        const m = txt.match(NIP_RE);
        if (!m) continue;
        const parent = node.parentElement;
        if (!parent) continue;
        const r = parent.getBoundingClientRect();
        if (!r || r.height <= 0 || r.width <= 0) continue;
        out.push({ nip: m[0], y: r.top + r.height/2, node, el: parent });
      }
    } catch(e){}
    return out;
  }

  function buildNipsWithDBSet() {
    const roots = collectRoots(document);
    const nipPoints = [];
    const dbPoints = [];
    for (const r of roots) {
      // Uwaga: dla shadowRoot używamy jego host.ownerDocument do walkerów
      const base = r.host ? r : (r.body || r);
      if (!base) continue;
      nipPoints.push(...collectNIPRects(base));
      dbPoints.push(...collectDBRects(base));
    }
    const flagged = new Set();
    if (nipPoints.length === 0 || dbPoints.length === 0) return flagged;
    for (const n of nipPoints) {
      let bestDY = Infinity;
      for (const d of dbPoints) {
        const dy = Math.abs(d.y - n.y);
        if (dy < bestDY) bestDY = dy;
      }
      if (bestDY <= ROW_Y_TOL) {
        flagged.add(n.nip);
      }
    }
    return flagged;
  }

  function addDbToText(text, nipsWithDB) {
    if (!text) return text;
    let out = text;
    for (const nip of nipsWithDB) {
      const re = new RegExp(`(${nip})(?!\\sDB)`, 'g');
      out = out.replace(re, `$1 DB`);
    }
    return out;
  }

  function onCopy() {
    try {
      const sel = window.getSelection && window.getSelection();
      if (!sel || sel.rangeCount === 0) return;
      const selectedText = sel.toString();
      if (!selectedText || selectedText.trim().length < 5) return;

      // Zmapuj w DOM (shadow-aware), które NIP-y mają DB (po geometrii)
      const nipsWithDB = buildNipsWithDBSet();
      if (nipsWithDB.size === 0) return;

      const patched = addDbToText(selectedText, nipsWithDB);
      if (patched === selectedText) return;

      // Nadpisz schowek po domyślnym kopiowaniu (strona nie widzi zmiany)
      setTimeout(() => {
        navigator.clipboard.writeText(patched).catch(()=>{});
      }, 0);
    } catch (e) {}
  }

  document.addEventListener('copy', onCopy, true);
})();