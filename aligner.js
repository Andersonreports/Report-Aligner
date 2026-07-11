/*
 * Client-side Word (.docx) alignment fixer.
 *
 * A .docx file is a zip of XML parts. This reads the reference template's
 * word/styles.xml (per-style alignment/font/spacing) and word/document.xml
 * (page setup), then rewrites the uploaded report's word/document.xml
 * (+ any headers/footers) so paragraphs/runs/tables/page-setup match the
 * template. Everything happens in the browser; no file is ever uploaded
 * to a server.
 */

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const PPR_ORDER = ['pStyle','keepNext','keepLines','pageBreakBefore','framePr','widowControl',
  'numPr','suppressLineNumbers','pBdr','shd','tabs','suppressAutoHyphens','kinsoku','wordWrap',
  'overflowPunct','topLinePunct','autoSpaceDE','autoSpaceDN','bidi','adjustRightInd','snapToGrid',
  'spacing','ind','contextualSpacing','mirrorIndents','suppressOverlap','jc','textDirection',
  'textAlignment','textboxTightWrap','outlineLvl','divId','cnfStyle','rPr','sectPr','pPrChange'];

const RPR_ORDER = ['rStyle','rFonts','b','bCs','i','iCs','caps','smallCaps','strike','dstrike',
  'outline','shadow','emboss','imprint','noProof','snapToGrid','vanish','webHidden','color',
  'spacing','w','kern','position','sz','szCs','highlight','u','effect','bdr','shd','fitText',
  'vertAlign','rtl','cs','em','lang','eastAsianLayout','specVanish','oMath','rPrChange'];

const SECTPR_ORDER = ['headerReference','footerReference','footnotePr','endnotePr','type','pgSz',
  'pgMar','paperSrc','pgBorders','lnNumType','pgNumType','cols','formProt','vAlign','noEndnote',
  'titlePg','textDirection','bidi','rtlGutter','docGrid','printerSettings','sectPrChange'];

function canon(s) {
  return (s || '').toLowerCase().replace(/[\s_-]/g, '');
}

function childrenByLocalName(el, name) {
  const out = [];
  for (const child of Array.from(el.childNodes)) {
    if (child.nodeType === 1 && child.localName === name) out.push(child);
  }
  return out;
}

function firstChildByLocalName(el, name) {
  const list = childrenByLocalName(el, name);
  return list.length ? list[0] : null;
}

function getOrCreateOrdered(doc, parent, name, orderArr) {
  const existing = firstChildByLocalName(parent, name);
  if (existing) return existing;
  const el = doc.createElementNS(W_NS, 'w:' + name);
  const idx = orderArr.indexOf(name);
  let inserted = false;
  if (idx !== -1) {
    for (const child of Array.from(parent.childNodes)) {
      if (child.nodeType !== 1) continue;
      const childIdx = orderArr.indexOf(child.localName);
      if (childIdx !== -1 && childIdx > idx) {
        parent.insertBefore(el, child);
        inserted = true;
        break;
      }
    }
  }
  if (!inserted) parent.appendChild(el);
  return el;
}

function getOrCreatePPr(doc, pEl) {
  let pPr = firstChildByLocalName(pEl, 'pPr');
  if (!pPr) {
    pPr = doc.createElementNS(W_NS, 'w:pPr');
    pEl.insertBefore(pPr, pEl.firstChild);
  }
  return pPr;
}

function getOrCreateRPr(doc, rEl) {
  let rPr = firstChildByLocalName(rEl, 'rPr');
  if (!rPr) {
    rPr = doc.createElementNS(W_NS, 'w:rPr');
    rEl.insertBefore(rPr, rEl.firstChild);
  }
  return rPr;
}

function directRuns(pEl) {
  return Array.from(pEl.childNodes).filter(n => n.nodeType === 1 && n.localName === 'r');
}

// ---------- Extract a style rule from a <w:style> element ----------

function extractRule(styleEl) {
  return extractRuleFromPPrRPr(firstChildByLocalName(styleEl, 'pPr'), firstChildByLocalName(styleEl, 'rPr'));
}

function extractRuleFromPPrRPr(pPr, rPr) {
  const rule = {};
  if (pPr) {
    const jc = firstChildByLocalName(pPr, 'jc');
    if (jc) rule.jc = jc.getAttributeNS(W_NS, 'val') || jc.getAttribute('w:val');

    const spacing = firstChildByLocalName(pPr, 'spacing');
    if (spacing) {
      rule.spacingBefore = spacing.getAttribute('w:before');
      rule.spacingAfter = spacing.getAttribute('w:after');
      rule.spacingLine = spacing.getAttribute('w:line');
      rule.spacingLineRule = spacing.getAttribute('w:lineRule');
    }

    const ind = firstChildByLocalName(pPr, 'ind');
    if (ind) {
      rule.indLeft = ind.getAttribute('w:left');
      rule.indFirstLine = ind.getAttribute('w:firstLine');
    }
  }

  if (rPr) {
    const rFonts = firstChildByLocalName(rPr, 'rFonts');
    if (rFonts) {
      rule.fontAscii = rFonts.getAttribute('w:ascii');
      rule.fontHAnsi = rFonts.getAttribute('w:hAnsi');
      rule.fontEastAsia = rFonts.getAttribute('w:eastAsia');
      rule.fontCs = rFonts.getAttribute('w:cs');
    }
    const sz = firstChildByLocalName(rPr, 'sz');
    if (sz) rule.sz = sz.getAttribute('w:val');
    const szCs = firstChildByLocalName(rPr, 'szCs');
    if (szCs) rule.szCs = szCs.getAttribute('w:val');

    const b = firstChildByLocalName(rPr, 'b');
    if (b) rule.bold = b.getAttribute('w:val') !== '0' && b.getAttribute('w:val') !== 'false';
    const i = firstChildByLocalName(rPr, 'i');
    if (i) rule.italic = i.getAttribute('w:val') !== '0' && i.getAttribute('w:val') !== 'false';

    const color = firstChildByLocalName(rPr, 'color');
    if (color) rule.color = color.getAttribute('w:val');
  }
  return rule;
}

function applyRuleToParagraph(doc, pEl, rule) {
  const pPr = getOrCreatePPr(doc, pEl);

  if (rule.jc) {
    getOrCreateOrdered(doc, pPr, 'jc', PPR_ORDER).setAttribute('w:val', rule.jc);
  }
  if (rule.spacingBefore || rule.spacingAfter || rule.spacingLine) {
    const spacing = getOrCreateOrdered(doc, pPr, 'spacing', PPR_ORDER);
    if (rule.spacingBefore) spacing.setAttribute('w:before', rule.spacingBefore);
    if (rule.spacingAfter) spacing.setAttribute('w:after', rule.spacingAfter);
    if (rule.spacingLine) spacing.setAttribute('w:line', rule.spacingLine);
    if (rule.spacingLineRule) spacing.setAttribute('w:lineRule', rule.spacingLineRule);
  }
  if (rule.indLeft || rule.indFirstLine) {
    const ind = getOrCreateOrdered(doc, pPr, 'ind', PPR_ORDER);
    if (rule.indLeft) ind.setAttribute('w:left', rule.indLeft);
    if (rule.indFirstLine) ind.setAttribute('w:firstLine', rule.indFirstLine);
  }

  for (const r of directRuns(pEl)) {
    const rPr = getOrCreateRPr(doc, r);
    if (rule.fontAscii) {
      const rFonts = getOrCreateOrdered(doc, rPr, 'rFonts', RPR_ORDER);
      rFonts.setAttribute('w:ascii', rule.fontAscii);
      if (rule.fontHAnsi) rFonts.setAttribute('w:hAnsi', rule.fontHAnsi);
      if (rule.fontEastAsia) rFonts.setAttribute('w:eastAsia', rule.fontEastAsia);
      if (rule.fontCs) rFonts.setAttribute('w:cs', rule.fontCs);
    }
    if (rule.sz) getOrCreateOrdered(doc, rPr, 'sz', RPR_ORDER).setAttribute('w:val', rule.sz);
    if (rule.szCs) getOrCreateOrdered(doc, rPr, 'szCs', RPR_ORDER).setAttribute('w:val', rule.szCs);
    if (rule.bold === true) getOrCreateOrdered(doc, rPr, 'b', RPR_ORDER).removeAttribute('w:val');
    if (rule.bold === false) getOrCreateOrdered(doc, rPr, 'b', RPR_ORDER).setAttribute('w:val', '0');
    if (rule.italic === true) getOrCreateOrdered(doc, rPr, 'i', RPR_ORDER).removeAttribute('w:val');
    if (rule.italic === false) getOrCreateOrdered(doc, rPr, 'i', RPR_ORDER).setAttribute('w:val', '0');
    if (rule.color) getOrCreateOrdered(doc, rPr, 'color', RPR_ORDER).setAttribute('w:val', rule.color);
  }
}

// ---------- Template profile ----------

function resolveEffectiveRule(styleId, styleById, rootRule, cache) {
  if (!styleId) return rootRule;
  if (cache[styleId]) return cache[styleId];
  const style = styleById[styleId];
  if (!style) return rootRule;
  // Break potential basedOn cycles (malformed documents) by seeding the cache
  // with the root rule before recursing.
  cache[styleId] = rootRule;
  const base = style.basedOnId ? resolveEffectiveRule(style.basedOnId, styleById, rootRule, cache) : rootRule;
  const resolved = { ...base, ...style.ownRule };
  cache[styleId] = resolved;
  return resolved;
}

function buildTemplateProfile(stylesXmlDoc, documentXmlDoc) {
  const profile = { rulesByCanonName: {}, pgMar: null, pgSz: null, tableStyleId: null, imageJc: null,
    tableStyleXmlById: {} };

  const docDefaults = firstChildByLocalName(stylesXmlDoc.documentElement, 'docDefaults');
  const rPrDefault = docDefaults ? firstChildByLocalName(docDefaults, 'rPrDefault') : null;
  const pPrDefault = docDefaults ? firstChildByLocalName(docDefaults, 'pPrDefault') : null;
  const rootRule = extractRuleFromPPrRPr(
    pPrDefault ? firstChildByLocalName(pPrDefault, 'pPr') : null,
    rPrDefault ? firstChildByLocalName(rPrDefault, 'rPr') : null);

  const styleEls = Array.from(stylesXmlDoc.getElementsByTagName('w:style'))
    .filter(s => s.getAttribute('w:type') === 'paragraph');

  const styleById = {};
  for (const styleEl of styleEls) {
    const id = styleEl.getAttribute('w:styleId');
    if (!id) continue;
    const basedOnEl = firstChildByLocalName(styleEl, 'basedOn');
    styleById[id] = {
      name: (firstChildByLocalName(styleEl, 'name') || {}).getAttribute ? firstChildByLocalName(styleEl, 'name').getAttribute('w:val') : id,
      basedOnId: basedOnEl ? basedOnEl.getAttribute('w:val') : null,
      ownRule: extractRule(styleEl),
    };
  }

  const cache = {};
  for (const [id, style] of Object.entries(styleById)) {
    if (!style.name) continue;
    profile.rulesByCanonName[canon(style.name)] = resolveEffectiveRule(id, styleById, rootRule, cache);
  }

  const tableStyleEls = Array.from(stylesXmlDoc.getElementsByTagName('w:style'))
    .filter(s => s.getAttribute('w:type') === 'table');
  for (const styleEl of tableStyleEls) {
    const id = styleEl.getAttribute('w:styleId');
    if (id) profile.tableStyleXmlById[id] = new XMLSerializer().serializeToString(styleEl);
  }

  const body = documentXmlDoc.getElementsByTagName('w:body')[0];
  let sectPr = null;
  if (body) sectPr = firstChildByLocalName(body, 'sectPr');
  if (!sectPr) {
    const all = documentXmlDoc.getElementsByTagName('w:sectPr');
    if (all.length) sectPr = all[0];
  }
  if (sectPr) {
    const pgSz = firstChildByLocalName(sectPr, 'pgSz');
    if (pgSz) profile.pgSz = { w: pgSz.getAttribute('w:w'), h: pgSz.getAttribute('w:h'), orient: pgSz.getAttribute('w:orient') };
    const pgMar = firstChildByLocalName(sectPr, 'pgMar');
    if (pgMar) {
      profile.pgMar = {};
      for (const attr of ['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter']) {
        const v = pgMar.getAttribute('w:' + attr);
        if (v != null) profile.pgMar[attr] = v;
      }
    }
  }

  const firstTbl = documentXmlDoc.getElementsByTagName('w:tbl')[0];
  if (firstTbl) {
    const tblPr = firstChildByLocalName(firstTbl, 'tblPr');
    const tblStyle = tblPr ? firstChildByLocalName(tblPr, 'tblStyle') : null;
    if (tblStyle) profile.tableStyleId = tblStyle.getAttribute('w:val');
  }

  for (const p of Array.from(documentXmlDoc.getElementsByTagName('w:p'))) {
    if (paragraphIsImageOnly(p)) {
      const pPr = firstChildByLocalName(p, 'pPr');
      const jc = pPr ? firstChildByLocalName(pPr, 'jc') : null;
      if (jc) profile.imageJc = jc.getAttribute('w:val');
      break;
    }
  }

  return profile;
}

function paragraphIsImageOnly(pEl) {
  const text = Array.from(pEl.getElementsByTagName('w:t')).map(t => t.textContent).join('').trim();
  if (text) return false;
  return pEl.getElementsByTagName('w:drawing').length > 0 || pEl.getElementsByTagName('w:pict').length > 0;
}

function styleIdToCanonName(stylesXmlDoc) {
  const map = {};
  for (const styleEl of Array.from(stylesXmlDoc.getElementsByTagName('w:style'))) {
    const id = styleEl.getAttribute('w:styleId');
    const nameEl = firstChildByLocalName(styleEl, 'name');
    const name = nameEl ? nameEl.getAttribute('w:val') : id;
    if (id) map[id] = canon(name);
  }
  return map;
}

function estimateBodyFontSize(documentXmlDoc, styleIdMap) {
  const counts = {};
  for (const p of Array.from(documentXmlDoc.getElementsByTagName('w:p'))) {
    const pPr = firstChildByLocalName(p, 'pPr');
    const pStyle = pPr ? firstChildByLocalName(pPr, 'pStyle') : null;
    const styleId = pStyle ? pStyle.getAttribute('w:val') : null;
    const cName = styleId ? styleIdMap[styleId] : 'normal';
    if (cName !== 'normal') continue;
    for (const r of directRuns(p)) {
      const rPr = firstChildByLocalName(r, 'rPr');
      const sz = rPr ? firstChildByLocalName(rPr, 'sz') : null;
      if (sz) {
        const v = sz.getAttribute('w:val');
        counts[v] = (counts[v] || 0) + 1;
      }
    }
  }
  let best = null, bestCount = -1;
  for (const [v, c] of Object.entries(counts)) {
    if (c > bestCount) { best = v; bestCount = c; }
  }
  return best ? parseInt(best, 10) : 22; // default 11pt
}

// Real-world report generators frequently tag the same logical heading with
// different Word style names from one draft to the next (or no heading style
// at all). Matching on the heading's own TEXT is far more reliable than
// matching on whatever style name happened to get applied. These lists are
// seeded from actual Anderson Diagnostics report templates/drafts.
const TITLE_TEXTS = [
  'whole exome sequencing analysis',
  'whole mitochondrial genome sequencing',
];

const HEADING2_TEXTS = [
  'clinical history', 'results', 'variant interpretation', 'cnv findings',
  'recommendations', 'methodology', 'sequence data attributes', 'disclaimer',
  'references', 'appendix 1: gene coverage', 'additional variant(s)',
  'list of additional uncertain significant variant(s) identified',
  'list of significant variant(s) identified related to the given phenotype',
  'this report has been reviewed and approved by', 'indication based analysis',
];

function normalizeHeadingText(text) {
  return text.trim().toLowerCase().replace(/\s+/g, ' ').replace(/[:.]+$/, '');
}

function matchesDictionary(text, dictionary) {
  const norm = normalizeHeadingText(text);
  if (!norm) return false;
  return dictionary.some(entry => norm === entry || norm.startsWith(entry));
}

function paragraphLooksLikeUntaggedHeading(pEl, bodySize) {
  const text = Array.from(pEl.getElementsByTagName('w:t')).map(t => t.textContent).join('').trim();
  if (!text || text.length > 120) return false;
  if (/[.,;]$/.test(text)) return false;
  const runs = directRuns(pEl).filter(r => Array.from(r.getElementsByTagName('w:t')).some(t => t.textContent.trim()));
  if (!runs.length) return false;
  let allBold = true;
  let anyBigger = false;
  for (const r of runs) {
    const rPr = firstChildByLocalName(r, 'rPr');
    const b = rPr ? firstChildByLocalName(rPr, 'b') : null;
    const isBold = b && b.getAttribute('w:val') !== '0' && b.getAttribute('w:val') !== 'false';
    if (!isBold) allBold = false;
    const sz = rPr ? firstChildByLocalName(rPr, 'sz') : null;
    if (sz && parseInt(sz.getAttribute('w:val'), 10) > bodySize) anyBigger = true;
  }
  return allBold || anyBigger;
}

// ---------- Apply profile to a document/header/footer XML part ----------

function applyProfileToPart(doc, documentXmlDoc, profile, styleIdMap, bodySize, options, titleState) {
  for (const p of Array.from(documentXmlDoc.getElementsByTagName('w:p'))) {
    if (paragraphIsImageOnly(p)) {
      const pPr = getOrCreatePPr(doc, p);
      getOrCreateOrdered(doc, pPr, 'jc', PPR_ORDER).setAttribute('w:val', profile.imageJc || 'center');
      continue;
    }

    const text = Array.from(p.getElementsByTagName('w:t')).map(t => t.textContent).join('');

    const pPr = firstChildByLocalName(p, 'pPr');
    const pStyleEl = pPr ? firstChildByLocalName(pPr, 'pStyle') : null;
    const styleId = pStyleEl ? pStyleEl.getAttribute('w:val') : null;
    const cName = styleId ? (styleIdMap[styleId] || canon(styleId)) : 'normal';

    let rule = null;

    if (matchesDictionary(text, TITLE_TEXTS)) {
      rule = profile.rulesByCanonName['title'];
    } else if (titleState && !titleState.seenFirstParagraph && text.trim()) {
      titleState.seenFirstParagraph = true;
      rule = profile.rulesByCanonName['title'];
    } else if (matchesDictionary(text, HEADING2_TEXTS)) {
      rule = profile.rulesByCanonName['heading2'];
    }

    if (!rule && options.heuristicHeadings && (cName === 'normal' || cName === 'bodytext')) {
      if (paragraphLooksLikeUntaggedHeading(p, bodySize)) {
        rule = profile.rulesByCanonName['heading2'] || profile.rulesByCanonName['heading1'];
      }
    }

    if (!rule) rule = profile.rulesByCanonName[cName];

    if (rule) applyRuleToParagraph(doc, p, rule);
  }

  if (options.normalizeTableStyle && profile.tableStyleId) {
    for (const tbl of Array.from(documentXmlDoc.getElementsByTagName('w:tbl'))) {
      let tblPr = firstChildByLocalName(tbl, 'tblPr');
      if (!tblPr) {
        tblPr = doc.createElementNS(W_NS, 'w:tblPr');
        tbl.insertBefore(tblPr, tbl.firstChild);
      }
      let tblStyle = firstChildByLocalName(tblPr, 'tblStyle');
      if (!tblStyle) {
        tblStyle = doc.createElementNS(W_NS, 'w:tblStyle');
        tblPr.insertBefore(tblStyle, tblPr.firstChild);
      }
      tblStyle.setAttribute('w:val', profile.tableStyleId);
    }
  }
}

function ensureTableStyleAvailable(reportStylesXmlDoc, profile) {
  if (!profile.tableStyleId) return false;
  const already = Array.from(reportStylesXmlDoc.getElementsByTagName('w:style'))
    .some(s => s.getAttribute('w:type') === 'table' && s.getAttribute('w:styleId') === profile.tableStyleId);
  if (already) return false;

  const rawXml = profile.tableStyleXmlById[profile.tableStyleId];
  if (!rawXml) return false;

  const wrapped = `<w:root xmlns:w="${W_NS}">${rawXml}</w:root>`;
  const parsed = new DOMParser().parseFromString(wrapped, 'application/xml');
  const styleEl = parsed.getElementsByTagName('w:style')[0];
  if (!styleEl) return false;

  const imported = reportStylesXmlDoc.importNode(styleEl, true);
  const stylesRoot = reportStylesXmlDoc.getElementsByTagName('w:styles')[0];
  if (!stylesRoot) return false;
  stylesRoot.appendChild(imported);
  return true;
}

function applyPageSetup(doc, documentXmlDoc, profile) {
  for (const sectPr of Array.from(documentXmlDoc.getElementsByTagName('w:sectPr'))) {
    if (profile.pgSz) {
      const pgSz = getOrCreateOrdered(doc, sectPr, 'pgSz', SECTPR_ORDER);
      if (profile.pgSz.w) pgSz.setAttribute('w:w', profile.pgSz.w);
      if (profile.pgSz.h) pgSz.setAttribute('w:h', profile.pgSz.h);
      if (profile.pgSz.orient) pgSz.setAttribute('w:orient', profile.pgSz.orient);
      else pgSz.removeAttribute('w:orient');
    }
    if (profile.pgMar) {
      const pgMar = getOrCreateOrdered(doc, sectPr, 'pgMar', SECTPR_ORDER);
      for (const [k, v] of Object.entries(profile.pgMar)) {
        pgMar.setAttribute('w:' + k, v);
      }
    }
  }
}

// ---------- Top-level entry points ----------

async function loadXml(zip, path) {
  const file = zip.file(path);
  if (!file) return null;
  const text = await file.async('string');
  return new DOMParser().parseFromString(text, 'application/xml');
}

async function extractTemplateProfile(templateArrayBuffer) {
  const zip = await JSZip.loadAsync(templateArrayBuffer);
  const stylesXml = await loadXml(zip, 'word/styles.xml');
  const documentXml = await loadXml(zip, 'word/document.xml');
  if (!stylesXml || !documentXml) {
    throw new Error('This does not look like a valid .docx file (missing styles.xml/document.xml).');
  }
  const profile = buildTemplateProfile(stylesXml, documentXml);
  profile.defaultHeader = await extractDefaultHeaderPart(zip);
  return profile;
}

async function extractDefaultHeaderPart(zip) {
  const docRelsXml = await loadXml(zip, 'word/_rels/document.xml.rels');
  const documentXml = await loadXml(zip, 'word/document.xml');
  if (!docRelsXml || !documentXml) return null;

  const sectPr = Array.from(documentXml.getElementsByTagName('w:sectPr'))[0];
  if (!sectPr) return null;
  const headerRefs = childrenByLocalName(sectPr, 'headerReference');
  const defaultRef = headerRefs.find(r => r.getAttribute('w:type') === 'default') || headerRefs[0];
  if (!defaultRef) return null;
  const rId = defaultRef.getAttribute('r:id');

  const rel = Array.from(docRelsXml.getElementsByTagName('Relationship'))
    .find(r => r.getAttribute('Id') === rId);
  if (!rel) return null;
  const target = rel.getAttribute('Target'); // e.g. "header1.xml"
  const partPath = 'word/' + target;
  const partFile = zip.file(partPath);
  if (!partFile) return null;

  const xml = await partFile.async('string');
  const media = [];
  const headerRelsPath = 'word/_rels/' + target + '.rels';
  const headerRelsFile = zip.file(headerRelsPath);
  if (headerRelsFile) {
    const headerRelsXml = new DOMParser().parseFromString(await headerRelsFile.async('string'), 'application/xml');
    for (const r of Array.from(headerRelsXml.getElementsByTagName('Relationship'))) {
      const relTarget = r.getAttribute('Target'); // relative to word/, e.g. "media/image1.jpeg"
      const mediaPath = 'word/' + relTarget.replace(/^\.\.\//, '');
      const mediaFile = zip.file(mediaPath);
      if (mediaFile) {
        media.push({ rId: r.getAttribute('Id'), target: relTarget.replace(/^\.\.\//, ''), bytes: await mediaFile.async('uint8array') });
      }
    }
  }
  return { xml, media, originalName: target };
}

function extToContentType(ext) {
  const map = { jpeg: 'image/jpeg', jpg: 'image/jpeg', png: 'image/png', gif: 'image/gif', bmp: 'image/bmp' };
  return map[ext.toLowerCase()] || 'image/jpeg';
}

async function injectTemplateHeaderIfMissing(zip, profile) {
  if (!profile.defaultHeader) return;

  const existingHeaderParts = Object.keys(zip.files).filter(n => /^word\/header\d*\.xml$/.test(n));
  if (existingHeaderParts.length > 0) return; // report already has its own header; don't touch it

  const documentXml = await loadXml(zip, 'word/document.xml');
  const docRelsXml = await loadXml(zip, 'word/_rels/document.xml.rels');
  const contentTypesXml = await loadXml(zip, '[Content_Types].xml');
  if (!documentXml || !docRelsXml || !contentTypesXml) return;

  const usedRIds = Array.from(docRelsXml.getElementsByTagName('Relationship')).map(r => r.getAttribute('Id'));
  let counter = 1;
  const nextRId = () => {
    let id;
    do { id = 'rIdAligner' + (counter++); } while (usedRIds.includes(id));
    usedRIds.push(id);
    return id;
  };

  const existingMedia = Object.keys(zip.files).filter(n => /^word\/media\//.test(n));
  const mediaRenames = {}; // originalTarget -> newTarget (if collision)
  let mediaCounter = existingMedia.length + 1;
  for (const m of profile.defaultHeader.media) {
    let target = m.target.replace(/^\.\.\//, '');
    if (existingMedia.includes('word/' + target)) {
      const ext = target.split('.').pop();
      target = `media/aligner_header_img${mediaCounter++}.${ext}`;
    }
    mediaRenames[m.rId] = target;
    zip.file('word/' + target, m.bytes);
    const ext = target.split('.').pop();
    if (!Array.from(contentTypesXml.getElementsByTagName('Default')).some(d => d.getAttribute('Extension') === ext)) {
      const def = contentTypesXml.createElementNS(contentTypesXml.documentElement.namespaceURI, 'Default');
      def.setAttribute('Extension', ext);
      def.setAttribute('ContentType', extToContentType(ext));
      contentTypesXml.documentElement.appendChild(def);
    }
  }

  let headerXmlText = profile.defaultHeader.xml;
  const headerRIdMap = {};
  for (const m of profile.defaultHeader.media) {
    headerRIdMap[m.rId] = nextRId();
  }
  for (const [oldRId, newRId] of Object.entries(headerRIdMap)) {
    headerXmlText = headerXmlText.split(`r:embed="${oldRId}"`).join(`r:embed="${newRId}"`);
  }

  const headerPartName = 'word/header1.xml';
  zip.file(headerPartName, headerXmlText);

  if (profile.defaultHeader.media.length > 0) {
    const relsDoc = new DOMParser().parseFromString(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>',
      'application/xml');
    for (const m of profile.defaultHeader.media) {
      const rel = relsDoc.createElementNS(relsDoc.documentElement.namespaceURI, 'Relationship');
      rel.setAttribute('Id', headerRIdMap[m.rId]);
      rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
      rel.setAttribute('Target', mediaRenames[m.rId]);
      relsDoc.documentElement.appendChild(rel);
    }
    zip.file('word/_rels/header1.xml.rels', new XMLSerializer().serializeToString(relsDoc));
  }

  const headerDocRId = nextRId();
  const rel = docRelsXml.createElementNS(docRelsXml.documentElement.namespaceURI, 'Relationship');
  rel.setAttribute('Id', headerDocRId);
  rel.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header');
  rel.setAttribute('Target', 'header1.xml');
  docRelsXml.documentElement.appendChild(rel);

  const override = contentTypesXml.createElementNS(contentTypesXml.documentElement.namespaceURI, 'Override');
  override.setAttribute('PartName', '/word/header1.xml');
  override.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
  contentTypesXml.documentElement.appendChild(override);

  for (const sectPr of Array.from(documentXml.getElementsByTagName('w:sectPr'))) {
    // Continuous section breaks don't start a new page, so they must not own a
    // header reference of their own: a header can only render at the top of a
    // physical page, and giving a continuous break one forces Word to promote
    // it to a hard page break, splitting the flow at that point.
    const typeEl = firstChildByLocalName(sectPr, 'type');
    if (typeEl && typeEl.getAttribute('w:val') === 'continuous') continue;
    const ref = getOrCreateOrdered(documentXml, sectPr, 'headerReference', SECTPR_ORDER);
    ref.setAttribute('w:type', 'default');
    ref.setAttribute('r:id', headerDocRId);
  }

  zip.file('word/document.xml', new XMLSerializer().serializeToString(documentXml));
  zip.file('word/_rels/document.xml.rels', new XMLSerializer().serializeToString(docRelsXml));
  zip.file('[Content_Types].xml', new XMLSerializer().serializeToString(contentTypesXml));
}

async function fixReportDocx(reportArrayBuffer, profile, options) {
  options = Object.assign({ heuristicHeadings: false, normalizeTableStyle: false, injectMissingHeader: true }, options || {});
  const zip = await JSZip.loadAsync(reportArrayBuffer);

  const documentXml = await loadXml(zip, 'word/document.xml');
  const stylesXml = await loadXml(zip, 'word/styles.xml');
  if (!documentXml || !stylesXml) {
    throw new Error('This does not look like a valid .docx file (missing styles.xml/document.xml).');
  }
  const styleIdMap = styleIdToCanonName(stylesXml);
  const bodySize = estimateBodyFontSize(documentXml, styleIdMap);

  let stylesXmlDirty = ensureTableStyleAvailable(stylesXml, profile);

  applyPageSetup(documentXml, documentXml, profile);
  const titleState = { seenFirstParagraph: false };
  applyProfileToPart(documentXml, documentXml, profile, styleIdMap, bodySize, options, titleState);
  zip.file('word/document.xml', new XMLSerializer().serializeToString(documentXml));
  if (stylesXmlDirty) {
    zip.file('word/styles.xml', new XMLSerializer().serializeToString(stylesXml));
  }

  const partNames = Object.keys(zip.files).filter(n => /^word\/(header|footer)\d*\.xml$/.test(n));
  for (const name of partNames) {
    const partXml = await loadXml(zip, name);
    if (!partXml) continue;
    applyProfileToPart(partXml, partXml, profile, styleIdMap, bodySize, options, null);
    zip.file(name, new XMLSerializer().serializeToString(partXml));
  }

  if (options.injectMissingHeader) {
    await injectTemplateHeaderIfMissing(zip, profile);
  }

  return zip.generateAsync({
    type: 'blob',
    mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

// Exposed to app.js
window.WordAligner = { extractTemplateProfile, fixReportDocx };
