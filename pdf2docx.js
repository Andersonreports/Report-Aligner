/*
 * Client-side PDF -> .docx converter.
 *
 * Reports sometimes only exist as PDF. This extracts their text (and, where it
 * can, tables) using pdf.js and emits a *minimal* .docx. The output is NOT the
 * final product: it is fed straight into aligner.js, which restyles it to match
 * a template (fonts, sizes, spacing, margins, header, table widths). So this
 * converter only has to recover STRUCTURE and TEXT correctly - it does not need
 * to get any formatting right, because the aligner overwrites all of that.
 *
 * A PDF has no notion of paragraph/heading/table; it is positioned glyphs plus
 * (sometimes) vector lines. We reconstruct:
 *   - lines      : text items grouped by baseline (y)
 *   - paragraphs : lines grouped by vertical spacing + left edge
 *   - tables      : runs of lines that share consistent column boundaries
 * Headings are deliberately left as plain paragraphs: the aligner re-identifies
 * them from their TEXT (its TITLE_TEXTS / HEADING2_TEXTS dictionary), which is
 * far more reliable than guessing from PDF font metrics.
 *
 * Everything runs in the browser; no file leaves the device.
 */
(function (global) {
  'use strict';

  const PT_TO_TWIP = 20; // 1 point = 20 twips

  function xmlEscape(s) {
    return String(s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ---------- 1. Extract positioned text (+ ruling lines) per page ----------
  //
  // All coordinates are kept in PDF user space with y pointing UP (origin at
  // the bottom-left), which is the space both the text-item transforms and the
  // path operators live in - so text and ruling lines can be compared directly.
  // Larger y == higher on the page.

  function matMul(m, n) {
    return [
      m[0] * n[0] + m[2] * n[1], m[1] * n[0] + m[3] * n[1],
      m[0] * n[2] + m[2] * n[3], m[1] * n[2] + m[3] * n[3],
      m[0] * n[4] + m[2] * n[5] + m[4], m[1] * n[4] + m[3] * n[5] + m[5],
    ];
  }
  function matApply(m, x, y) {
    return [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];
  }

  // Walk the page's operator list, tracking the CTM, and collect the bounding
  // box of every stroked/filled path. Returns rectangles in user space.
  function harvestPathRects(opList, OPS) {
    let ctm = [1, 0, 0, 1, 0, 0];
    const stack = [];
    const rects = [];
    for (let i = 0; i < opList.fnArray.length; i++) {
      const fn = opList.fnArray[i];
      const a = opList.argsArray[i];
      if (fn === OPS.save) stack.push(ctm.slice());
      else if (fn === OPS.restore) ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === OPS.transform) ctm = matMul(ctm, a);
      else if (fn === OPS.constructPath) {
        const segOps = a[0], segArgs = a[1];
        let j = 0;
        const pts = [];
        for (const op of segOps) {
          if (op === OPS.rectangle) {
            const x = segArgs[j++], y = segArgs[j++], w = segArgs[j++], h = segArgs[j++];
            pts.push(matApply(ctm, x, y), matApply(ctm, x + w, y + h));
          } else if (op === OPS.moveTo || op === OPS.lineTo) {
            const x = segArgs[j++], y = segArgs[j++];
            pts.push(matApply(ctm, x, y));
          } else if (op === OPS.curveTo) { j += 6; }
          else if (op === OPS.curveTo2 || op === OPS.curveTo3) { j += 4; }
        }
        if (pts.length) {
          const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
          rects.push({ x0: Math.min(...xs), x1: Math.max(...xs), y0: Math.min(...ys), y1: Math.max(...ys) });
        }
      }
    }
    return rects;
  }

  // Turn path rectangles into thin ruling lines (vertical + horizontal). Both a
  // filled cell/header box and a plain bordered cell contribute their four
  // edges, so grid lines are recovered whether or not a cell is shaded.
  function rectsToRulings(rects, pageW, pageH) {
    const pageArea = pageW * pageH;
    const vLines = []; // { x, y0, y1 }
    const hLines = []; // { y, x0, x1 }
    const THIN = 2.5;
    for (const r of rects) {
      const w = r.x1 - r.x0, h = r.y1 - r.y0;
      if (w * h > pageArea * 0.5) continue; // page background / watermark
      const thin = Math.min(w, h) <= THIN;
      if (thin && Math.max(w, h) > 4) {
        if (h >= w) vLines.push({ x: (r.x0 + r.x1) / 2, y0: r.y0, y1: r.y1 });
        else hLines.push({ y: (r.y0 + r.y1) / 2, x0: r.x0, x1: r.x1 });
      } else if (w > 4 && h > 4) {
        // A cell/box outline or fill: contribute all four edges.
        vLines.push({ x: r.x0, y0: r.y0, y1: r.y1 }, { x: r.x1, y0: r.y0, y1: r.y1 });
        hLines.push({ y: r.y0, x0: r.x0, x1: r.x1 }, { y: r.y1, x0: r.x0, x1: r.x1 });
      }
    }
    return { vLines, hLines };
  }

  async function extractPages(arrayBuffer, pdfjsLib) {
    const loadingTask = pdfjsLib.getDocument({ data: arrayBuffer });
    const pdf = await loadingTask.promise;
    const OPS = pdfjsLib.OPS;
    const pages = [];
    for (let n = 1; n <= pdf.numPages; n++) {
      const page = await pdf.getPage(n);
      const viewport = page.getViewport({ scale: 1 });
      const content = await page.getTextContent();
      const items = [];
      for (const it of content.items) {
        if (typeof it.str !== 'string' || !it.str.length) continue;
        const t = it.transform; // [a,b,c,d,e,f]; user space, y up
        const fontHeight = Math.hypot(t[2], t[3]) || it.height || 0;
        items.push({ str: it.str, x: t[4], y: t[5], w: it.width || 0, h: fontHeight, font: it.fontName || '' });
      }
      let rulings = { vLines: [], hLines: [] };
      try {
        const opList = await page.getOperatorList();
        rulings = rectsToRulings(harvestPathRects(opList, OPS), viewport.width, viewport.height);
      } catch (e) { /* text-only fallback */ }
      pages.push({ items, rulings, width: viewport.width, height: viewport.height });
    }
    return pages;
  }

  // ---------- 2. Small geometry helpers ----------

  // Merge nearby scalar values into representative cluster centres.
  function clusterVals(vals, tol) {
    if (!vals.length) return [];
    const sorted = vals.slice().sort((a, b) => a - b);
    const out = [];
    let grp = [sorted[0]];
    for (let i = 1; i < sorted.length; i++) {
      if (sorted[i] - grp[grp.length - 1] <= tol) grp.push(sorted[i]);
      else { out.push(grp.reduce((a, b) => a + b, 0) / grp.length); grp = [sorted[i]]; }
    }
    out.push(grp.reduce((a, b) => a + b, 0) / grp.length);
    return out;
  }

  // Group text items into visual lines (top-to-bottom; y points up).
  function groupLines(items) {
    if (!items.length) return [];
    const sorted = items.slice().sort((a, b) => (b.y - a.y) || (a.x - b.x));
    const lines = [];
    let cur = null;
    for (const it of sorted) {
      const tol = Math.max(2, (it.h || 8) * 0.5);
      if (cur && Math.abs(it.y - cur.y) <= tol) {
        cur.items.push(it);
      } else {
        cur = { y: it.y, items: [it] };
        lines.push(cur);
      }
    }
    for (const line of lines) {
      line.items.sort((a, b) => a.x - b.x);
      line.x0 = line.items[0].x;
      line.h = Math.max(...line.items.map(i => i.h || 0));
      line.text = joinItems(line.items);
    }
    return lines;
  }

  // Join items on one line, inserting a space where there's a horizontal gap
  // wider than roughly one space character.
  function joinItems(items) {
    let out = '';
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (i > 0) {
        const prev = items[i - 1];
        const gap = it.x - (prev.x + prev.w);
        const spaceW = (prev.h || 8) * 0.25;
        if (gap > spaceW && !/\s$/.test(out) && !/^\s/.test(it.str)) out += ' ';
      }
      out += it.str;
    }
    return out.replace(/\s+/g, ' ').trim();
  }

  // ---------- 3. Reconstruct tables from ruling lines ----------
  //
  // A table's row height varies per table (and can be much larger than the
  // thin gap between two unrelated ruling lines), so there is no single "gap
  // size" threshold that safely tells rows-of-one-table apart from
  // start-of-another-table. What's actually true regardless of row height:
  // every ruling line belonging to the SAME table touches or overlaps another
  // ruling line of that same table (shared corners, a column divider crossing
  // a row divider, etc.), while lines from two visually separate tables never
  // touch. So tables are found as connected components of the ruling-line
  // "touch" graph, not by any fixed distance threshold.

  const EDGE_TOL = 3; // pt: tolerance for "do these lines touch"

  function unionFind(n) {
    const parent = Array.from({ length: n }, (_, i) => i);
    function find(x) { while (parent[x] !== x) { parent[x] = parent[parent[x]]; x = parent[x]; } return x; }
    function union(a, b) { const ra = find(a), rb = find(b); if (ra !== rb) parent[ra] = rb; }
    return { find, union };
  }

  function rangesTouch(a0, a1, b0, b1, tol) {
    return a0 <= b1 + tol && b0 <= a1 + tol;
  }

  function detectTables(page) {
    const { vLines, hLines } = page.rulings;
    if (!hLines.length || !vLines.length) return [];

    // segs[i]: { type: 'v'|'h', ... } — indices are shared with the union-find.
    const segs = [
      ...vLines.map(l => ({ type: 'v', x: l.x, y0: Math.min(l.y0, l.y1), y1: Math.max(l.y0, l.y1) })),
      ...hLines.map(l => ({ type: 'h', y: l.y, x0: Math.min(l.x0, l.x1), x1: Math.max(l.x0, l.x1) })),
    ];
    const uf = unionFind(segs.length);
    for (let i = 0; i < segs.length; i++) {
      for (let j = i + 1; j < segs.length; j++) {
        const a = segs[i], b = segs[j];
        let touch;
        if (a.type === 'v' && b.type === 'v') {
          touch = Math.abs(a.x - b.x) <= EDGE_TOL && rangesTouch(a.y0, a.y1, b.y0, b.y1, EDGE_TOL);
        } else if (a.type === 'h' && b.type === 'h') {
          touch = Math.abs(a.y - b.y) <= EDGE_TOL && rangesTouch(a.x0, a.x1, b.x0, b.x1, EDGE_TOL);
        } else {
          const v = a.type === 'v' ? a : b, h = a.type === 'v' ? b : a;
          touch = v.x >= h.x0 - EDGE_TOL && v.x <= h.x1 + EDGE_TOL && h.y >= v.y0 - EDGE_TOL && h.y <= v.y1 + EDGE_TOL;
        }
        if (touch) uf.union(i, j);
      }
    }

    const groups = {};
    for (let i = 0; i < segs.length; i++) {
      const root = uf.find(i);
      (groups[root] = groups[root] || []).push(segs[i]);
    }

    const tables = [];
    for (const groupSegs of Object.values(groups)) {
      const hSegs = groupSegs.filter(s => s.type === 'h');
      const vSegs = groupSegs.filter(s => s.type === 'v');
      if (!hSegs.length || !vSegs.length) continue;
      const rowEdges = clusterVals(hSegs.map(s => s.y), EDGE_TOL).sort((a, b) => b - a);
      const colEdges = clusterVals(vSegs.map(s => s.x), EDGE_TOL).sort((a, b) => a - b);
      if (rowEdges.length < 2 || colEdges.length < 2) continue;
      const nRows = rowEdges.length - 1, nCols = colEdges.length - 1;
      if (nRows * nCols < 2) continue; // a single bordered box, not a table
      tables.push({ yTop: rowEdges[0], yBot: rowEdges[rowEdges.length - 1], rowEdges, colEdges, nRows, nCols });
    }
    return tables;
  }

  // ---------- 4. Emit OOXML ----------

  function paraXml(text) {
    if (!text) return '<w:p/>';
    return `<w:p><w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r></w:p>`;
  }

  function cellItemsXml(items) {
    if (!items.length) return '<w:p/>';
    return groupLines(items).map(l => paraXml(l.text)).join('') || '<w:p/>';
  }

  function tableXml(table, items) {
    const { rowEdges, colEdges, nRows, nCols } = table;
    const widthsTw = [];
    for (let c = 0; c < nCols; c++) widthsTw.push(Math.max(120, Math.round((colEdges[c + 1] - colEdges[c]) * PT_TO_TWIP)));
    const grid = widthsTw.map(w => `<w:gridCol w:w="${w}"/>`).join('');

    // Bucket items into grid cells by their centre point.
    const cellItems = Array.from({ length: nRows }, () => Array.from({ length: nCols }, () => []));
    for (const it of items) {
      const cx = it.x + (it.w || 0) / 2;
      const cy = it.y + (it.h || 0) * 0.3;
      let c = -1;
      for (let k = 0; k < nCols; k++) { if (cx >= colEdges[k] - EDGE_TOL && cx < colEdges[k + 1] + EDGE_TOL) { c = k; break; } }
      let r = -1;
      for (let k = 0; k < nRows; k++) { if (cy <= rowEdges[k] + EDGE_TOL && cy > rowEdges[k + 1] - EDGE_TOL) { r = k; break; } }
      if (r >= 0 && c >= 0) cellItems[r][c].push(it);
    }

    let rowsXml = '';
    for (let r = 0; r < nRows; r++) {
      let tcs = '';
      for (let c = 0; c < nCols; c++) {
        tcs += `<w:tc><w:tcPr><w:tcW w:w="${widthsTw[c]}" w:type="dxa"/></w:tcPr>${cellItemsXml(cellItems[r][c])}</w:tc>`;
      }
      rowsXml += `<w:tr>${tcs}</w:tr>`;
    }

    return `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders>` +
      ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']
        .map(s => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('') +
      `</w:tblBorders></w:tblPr><w:tblGrid>${grid}</w:tblGrid>${rowsXml}</w:tbl>`;
  }

  // Merge consecutive body lines into paragraphs (wrapped lines share a left
  // edge and sit one line-height apart); a bigger gap or an indent starts a new
  // paragraph.
  function linesToParagraphs(lines) {
    const paras = [];
    let cur = null;
    for (const line of lines) {
      if (cur) {
        const gap = cur.y - line.y;
        const sameLeft = Math.abs(line.x0 - cur.x0) < 12;
        if (gap > 0 && gap < cur.h * 1.8 && sameLeft) {
          cur.text += ' ' + line.text; cur.y = line.y; continue;
        }
      }
      cur = { y: line.y, x0: line.x0, h: line.h, text: line.text };
      paras.push(cur);
    }
    return paras;
  }

  function buildPageXml(page) {
    const tables = detectTables(page);
    const inTable = (it) => tables.some(t =>
      it.x >= t.colEdges[0] - EDGE_TOL && it.x <= t.colEdges[t.colEdges.length - 1] + EDGE_TOL &&
      it.y <= t.yTop + EDGE_TOL && it.y >= t.yBot - EDGE_TOL);

    const elements = [];
    // Text that isn't inside any table -> paragraphs.
    const looseItems = page.items.filter(it => !inTable(it));
    for (const p of linesToParagraphs(groupLines(looseItems))) {
      elements.push({ y: p.y, xml: paraXml(p.text) });
    }
    // Tables, each fed the items that fall within its bounds.
    for (const t of tables) {
      const its = page.items.filter(it =>
        it.x >= t.colEdges[0] - EDGE_TOL && it.x <= t.colEdges[t.colEdges.length - 1] + EDGE_TOL &&
        it.y <= t.yTop + EDGE_TOL && it.y >= t.yBot - EDGE_TOL);
      elements.push({ y: t.yTop, xml: tableXml(t, its) });
    }
    // Emit top-to-bottom.
    elements.sort((a, b) => b.y - a.y);
    return elements.map(e => e.xml).join('');
  }

  function buildDocumentXml(pages) {
    let body = '';
    for (const page of pages) body += buildPageXml(page);
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:body>${body}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` +
      `</w:sectPr></w:body></w:document>`;
  }

  const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
    `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>` +
    `<w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `</w:styles>`;

  const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
    `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
    `</Types>`;

  const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships>`;

  const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
    `</Relationships>`;

  async function buildDocxBlob(documentXml) {
    const zip = new JSZip();
    zip.file('[Content_Types].xml', CONTENT_TYPES);
    zip.folder('_rels').file('.rels', ROOT_RELS);
    const word = zip.folder('word');
    word.file('document.xml', documentXml);
    word.file('styles.xml', STYLES_XML);
    word.folder('_rels').file('document.xml.rels', DOC_RELS);
    const isNode = typeof process !== 'undefined' && !!(process.versions && process.versions.node);
    return zip.generateAsync({
      type: isNode ? 'nodebuffer' : 'blob',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      compression: 'DEFLATE',
    });
  }

  // ---------- Public entry point ----------

  // Returns a Blob (browser) of a minimal .docx reconstructed from the PDF.
  async function convertPdfToDocx(arrayBuffer, opts) {
    opts = opts || {};
    const pdfjsLib = opts.pdfjsLib || global.pdfjsLib;
    if (!pdfjsLib) throw new Error('pdf.js (pdfjsLib) is not loaded.');
    const pages = await extractPages(arrayBuffer, pdfjsLib);
    const documentXml = buildDocumentXml(pages);
    return buildDocxBlob(documentXml);
  }

  const api = { convertPdfToDocx, _internals: { groupLines, detectTables, buildPageXml, buildDocumentXml, extractPages } };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  global.PdfToDocx = api;
})(typeof window !== 'undefined' ? window : globalThis);
