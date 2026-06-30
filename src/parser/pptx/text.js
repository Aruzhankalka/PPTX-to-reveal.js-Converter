const { asArray } = require('./xml');
const { emuToPx, pptxRotationToDegrees } = require('./units');
const { resolveSolidFillCss } = require('./color');

/**
 * Placeholder types that are pure presentation metadata with no visual text.
 * 'hdr' (header) has no rendered equivalent in reveal.js.
 * 'ftr' (footer), 'sldNum' (slide number), and 'dt' (date/time) ARE rendered
 * when they carry actual text content — they must not be skipped here.
 */
const SKIP_PLACEHOLDER_TYPES = new Set(['hdr']);

/**
 * PPTX-spec default font sizes per placeholder type and indent level (0-based).
 * Used only as absolute last resort — after run sz, paragraph defRPr,
 * shape lstStyle, AND master txStyles have all been tried.
 * Values match ECMA-376 Annex E (Office Open XML default theme).
 */
const PLACEHOLDER_DEFAULT_SIZES = {
  title:    { 0: '40pt' },
  ctrTitle: { 0: '40pt' },
  subTitle: { 0: '28pt' },
  body:     { 0: '24pt', 1: '20pt', 2: '18pt', 3: '18pt', 4: '18pt' },
};

/**
 * Look up a font size from the master's txStyles for a given placeholder type
 * and indent level (0-based).
 * txStyles entries are objects: { size, lineSpacing, spaceBefore, spaceAfter }
 */
function sizeFromTxStyles(txStyles, phType, indentLevel) {
  if (!txStyles) return null;
  // ftr/sldNum/dt have their own per-placeholder styling from the layout/master
  // lstStyle; the master body txStyles font sizes do not apply to them.
  if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt' || phType === 'hdr') return null;
  const lvl = (indentLevel || 0) + 1; // txStyles keys are 1-based
  let section;
  if (phType === 'title' || phType === 'ctrTitle') section = txStyles.title;
  else if (phType === 'subTitle') section = txStyles.body;
  else if (phType == null) section = txStyles.other || txStyles.body; // drawing objects
  else section = txStyles.body;
  if (!section) return null;
  const entry = section[lvl] || section[1] || null;
  if (!entry) return null;
  // Entry is now an object with a .size field; guard against old string format.
  return typeof entry === 'string' ? entry : (entry.size || null);
}

/**
 * Look up line-spacing, space-before, and space-after from txStyles per-level
 * entries as CSS-ready values, for use as fallbacks when the paragraph's own
 * <a:pPr> doesn't specify these properties.
 *
 * Returns null when txStyles is absent or the level has no spacing data.
 */
function spacingFromTxStyles(txStyles, phType, indentLevel) {
  if (!txStyles) return null;
  if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt' || phType === 'hdr') return null;
  const lvl = (indentLevel || 0) + 1;
  let section;
  if (phType === 'title' || phType === 'ctrTitle') section = txStyles.title;
  else if (phType === 'subTitle') section = txStyles.body;
  else if (phType == null) section = txStyles.other; // drawing objects: never inherit body spacing
  else section = txStyles.body;
  if (!section) return null;
  const entry = section[lvl] || section[1] || null;
  if (!entry || typeof entry === 'string') return null;
  const result = {};
  if (entry.lineSpacing != null) result['line-spacing']  = entry.lineSpacing;
  if (entry.spaceBefore != null) result['space-before']  = entry.spaceBefore;
  if (entry.spaceAfter  != null) result['space-after']   = entry.spaceAfter;
  return Object.keys(result).length > 0 ? result : null;
}

/**
 * Extract line-spacing, space-before, and space-after from a raw <a:lstStyle>
 * node at the given indent level (0-based).  Mirrors the logic in
 * extractParagraphFormatting so returned values are CSS-ready strings.
 *
 * Returns null when the node is absent or the level has no explicit spacing.
 * A non-null result (even an empty object) signals that this level is defined
 * in the lstStyle, allowing callers to skip the master txStyles fallback.
 */
function spacingFromLstStyle(lstStyle, indentLevel) {
  if (!lstStyle) return null;
  const lvlNum = Math.min(9, Math.max(1, (indentLevel || 0) + 1));
  const lvlPPr = lstStyle[`a:lvl${lvlNum}pPr`];
  if (!lvlPPr) return null;

  const result = {};

  const lnSpc = lvlPPr['a:lnSpc'];
  if (lnSpc) {
    const pct = lnSpc['a:spcPct'];
    const pts = lnSpc['a:spcPts'];
    if (pct && pct['@_val']) {
      const v = Number(pct['@_val']) / 100000;
      if (!Number.isNaN(v)) result['line-spacing'] = String(v);
    } else if (pts && pts['@_val']) {
      const v = Number(pts['@_val']) / 100;
      if (!Number.isNaN(v)) result['line-spacing'] = `${v}pt`;
    }
  }

  for (const [attr, key] of [['a:spcBef', 'space-before'], ['a:spcAft', 'space-after']]) {
    const node = lvlPPr[attr];
    if (!node) continue;
    const pts = node['a:spcPts'];
    const pct = node['a:spcPct'];
    if (pts && pts['@_val'] != null) {
      const v = Number(pts['@_val']) / 100;
      if (!Number.isNaN(v)) result[key] = `${v}pt`;
    } else if (pct && pct['@_val'] != null) {
      const v = Number(pct['@_val']) / 100000;
      if (!Number.isNaN(v)) result[key] = `${v}em`;
    }
  }

  // Return the result object (possibly empty) to signal the level is defined,
  // or null when the level itself doesn't exist in the lstStyle.
  return result;
}

/**
 * Look up the latin font family from txStyles per-level entries, for use as
 * a fallback when the run's own <a:rPr> has no <a:latin typeface="...">.
 * Mirrors sizeFromTxStyles: same section-selection rules, same level-fallback
 * pattern.  Drawing objects (phType === null) use otherStyle only — they must
 * not inherit the body placeholder font.
 *
 * Returns a CSS font-family string (e.g. 'Calibri') or null.
 */
function fontFromTxStyles(txStyles, phType, indentLevel) {
  if (!txStyles) return null;
  if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt' || phType === 'hdr') return null;
  const lvl = (indentLevel || 0) + 1;
  let section;
  if (phType === 'title' || phType === 'ctrTitle') section = txStyles.title;
  else if (phType === 'subTitle') section = txStyles.body;
  else if (phType == null) section = txStyles.other; // drawing objects: never inherit body font
  else section = txStyles.body;
  if (!section) return null;
  const entry = section[lvl] || section[1] || null;
  if (!entry || typeof entry === 'string') return null;
  return entry.font ?? null;
}

/**
 * Look up bold/italic/underline flags from master txStyles for a given
 * placeholder type and indent level (0-based).
 * Returns an object with only keys that are explicitly true in the txStyles entry.
 */
function biuFromTxStyles(txStyles, phType, indentLevel) {
  if (!txStyles) return {};
  if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt' || phType === 'hdr') return {};
  const lvl = (indentLevel || 0) + 1;
  let section;
  if (phType === 'title' || phType === 'ctrTitle') section = txStyles.title;
  else if (phType === 'subTitle') section = txStyles.body;
  else if (phType == null) section = txStyles.other || txStyles.body;
  else section = txStyles.body;
  if (!section) return {};
  const entry = section[lvl] || section[1] || null;
  if (!entry || typeof entry === 'string') return {};
  const biu = {};
  if (entry.bold) biu.weight = 'bold';
  if (entry.italic) biu.italics = true;
  if (entry.underline) biu['text-decoration'] = 'underline';
  return biu;
}

/**
 * Look up a font size from a shape's <a:lstStyle> for a given indent level.
 * Returns a CSS string like '24pt', or null if not found.
 */
function sizeFromLstStyle(lstStyle, indentLevel) {
  if (!lstStyle) return null;
  const lvlNum = Math.min(9, Math.max(1, (indentLevel || 0) + 1));
  const lvlPPr = lstStyle[`a:lvl${lvlNum}pPr`];
  const defRPr = lvlPPr && lvlPPr['a:defRPr'];
  if (!defRPr || !defRPr['@_sz']) return null;
  const pt = Number(defRPr['@_sz']) / 100;
  return Number.isNaN(pt) ? null : `${pt}pt`;
}

/**
 * Get the <a:defRPr> node from a shape's <a:lstStyle> at the given indent level.
 * Used as the lowest-priority fallback for bold/italic/underline inheritance.
 */
function lstStyleDefRPr(lstStyle, indentLevel) {
  if (!lstStyle) return null;
  const lvlNum = Math.min(9, Math.max(1, (indentLevel || 0) + 1));
  const lvlPPr = lstStyle[`a:lvl${lvlNum}pPr`];
  return (lvlPPr && lvlPPr['a:defRPr']) || null;
}

/**
 * Extract bold/italic/underline from any rPr-like node (rPr or defRPr).
 * Accepts both "1" and "true" since OOXML ST_TextBooleanType allows both.
 * Only returns keys that are explicitly enabled; absent or false/0 → nothing.
 */
function extractBIU(node) {
  if (!node) return {};
  const f = {};
  const b = node['@_b'];
  if (b === '1' || b === 'true') f.weight = 'bold';
  const i = node['@_i'];
  if (i === '1' || i === 'true') f.italics = true;
  const u = node['@_u'];
  if (u && u !== 'none') f['text-decoration'] = 'underline';
  return f;
}

/**
 * PPTX-spec defaults per placeholder type + indent level (last-resort fallback).
 * ftr/sldNum/dt/hdr have their own per-placeholder styling from the layout/master;
 * they must not receive the body fallback size.
 */
function placeholderFallbackSize(phType, indentLevel) {
  if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt' || phType === 'hdr') return null;
  const lvl = Math.max(0, Math.min(4, indentLevel || 0));
  const table = PLACEHOLDER_DEFAULT_SIZES[phType] || PLACEHOLDER_DEFAULT_SIZES.body;
  return table[lvl] || null;
}

/**
 * Extract a run's formatting from <a:rPr> attributes.
 * Handles bold, italics, underline/strikethrough, color, font family, font size.
 * Returns empty string when no formatting is present, so we can omit style="".
 *
 * @param {object} rPr - the parsed <a:rPr> node
 * @returns {object|undefined}
 */
function extractRunFormatting(rPr) {
  if (!rPr) return undefined;
  const f = { ...extractBIU(rPr) };

  if (rPr['@_strike'] && rPr['@_strike'] !== 'noStrike') {
    f['text-decoration'] = 'strikethrough';
  }
  // <a:rPr sz="2400"> means 24pt — sz is in 100ths of a point.
  if (rPr['@_sz']) {
    const pt = Number(rPr['@_sz']) / 100;
    if (!Number.isNaN(pt)) f.size = pt + 'pt';
  }
  // Font color: explicit hex or theme slot reference
  const fillColor = resolveSolidFillCss(rPr['a:solidFill']);
  if (fillColor) f.color = fillColor;
  // Font family: <a:latin typeface="Arial"/>
  const latin = rPr['a:latin'];
  if (latin && latin['@_typeface']) f.font = latin['@_typeface'];

  return Object.keys(f).length > 0 ? f : undefined;
}

/**
 * Extract a paragraph's formatting from <a:pPr>.
 * Reads alignment, indent level, list type, line spacing, and space before/after.
 */

// ---------------------------------------------------------------------------
// Bullet resolution helpers (lstStyle / txStyles cascade)
// ---------------------------------------------------------------------------

/**
 * Extract bullet info from a single <a:lvl{N}pPr> node.
 * Returns { type:'bullets'|'numbered', char? } | { type:'none' } | null.
 */
function bulletFromLvlPPr(lvlPPr) {
  if (!lvlPPr) return null;
  if (lvlPPr['a:buNone'] !== undefined) return { type: 'none' };
  if (lvlPPr['a:buChar'] !== undefined) {
    const char = lvlPPr['a:buChar']['@_char'];
    return { type: 'bullets', char: char || '•' };
  }
  if (lvlPPr['a:buAutoNum'] !== undefined) return { type: 'numbered' };
  return null;
}

/**
 * Look up bullet info in a <a:lstStyle> node at the given indent level
 * (0-based; maps to a:lvl1pPr … a:lvl9pPr).
 */
function bulletFromLstStyle(lstStyleNode, indentLevel) {
  if (!lstStyleNode) return null;
  const key = `a:lvl${Math.min(9, Math.max(1, (indentLevel || 0) + 1))}pPr`;
  return bulletFromLvlPPr(lstStyleNode[key]);
}

/**
 * Look up bullet info from master txStyles at the given placeholder type and
 * indent level.  Mirrors the pattern used by spacingFromTxStyles.
 */
function bulletFromTxStyles(txStyles, phType, indentLevel) {
  if (!txStyles) return null;
  const styleKey = phType === 'title' ? 'p:titleStyle' : 'p:bodyStyle';
  const style = txStyles[styleKey];
  if (!style) return null;
  const key = `a:lvl${Math.min(9, Math.max(1, (indentLevel || 0) + 1))}pPr`;
  return bulletFromLvlPPr(style[key]);
}

function extractParagraphFormatting(pPr) {
  if (!pPr) return undefined;
  const f = {};

  if (pPr['@_algn']) {
    const map = { l: 'left', r: 'right', ctr: 'center', just: 'justify' };
    if (map[pPr['@_algn']]) f.align = map[pPr['@_algn']];
  }
  if (pPr['@_lvl']) {
    const lvl = Number(pPr['@_lvl']);
    if (!Number.isNaN(lvl)) f['indent-level'] = lvl;
  }

  // List type
  if (pPr['a:buNone'] !== undefined) f['list-type'] = 'none';
  else if (pPr['a:buAutoNum'] !== undefined) f['list-type'] = 'numbered';
  else if (pPr['a:buChar'] !== undefined) f['list-type'] = 'bullets';

  // <a:lnSpc> — line spacing
  // spcPct val is in 1000ths of a percent (100000 = 100% = unitless 1.0)
  // spcPts val is in 100ths of a point
  const lnSpc = pPr['a:lnSpc'];
  if (lnSpc) {
    const pct = lnSpc['a:spcPct'];
    const pts = lnSpc['a:spcPts'];
    if (pct && pct['@_val']) {
      const v = Number(pct['@_val']) / 100000;
      if (!Number.isNaN(v)) f['line-spacing'] = String(v);
    } else if (pts && pts['@_val']) {
      const v = Number(pts['@_val']) / 100;
      if (!Number.isNaN(v)) f['line-spacing'] = `${v}pt`;
    }
  }

  // <a:spcBef> / <a:spcAft> — paragraph spacing (pts val in 100ths of a point)
  for (const [attr, key] of [['a:spcBef', 'space-before'], ['a:spcAft', 'space-after']]) {
    const node = pPr[attr];
    if (!node) continue;
    const pts = node['a:spcPts'];
    const pct = node['a:spcPct'];
    if (pts && pts['@_val'] != null) {
      const v = Number(pts['@_val']) / 100;
      if (!Number.isNaN(v)) f[key] = `${v}pt`;
    } else if (pct && pct['@_val'] != null) {
      const v = Number(pct['@_val']) / 100000;
      if (!Number.isNaN(v)) f[key] = `${v}em`;
    }
  }

  return Object.keys(f).length > 0 ? f : undefined;
}

/**
 * Convert a single <a:r> to an IR run.
 * fallbackSize: CSS size string (e.g. '24pt') used when <a:rPr> has no explicit sz.
 * slideRels:   parsed relationship map for this slide (optional); used to resolve
 *              hyperlink rIds from <a:hlinkClick r:id="..."/>.
 */
function runToIr(aR, fallbackSize, fallbackBIU, slideRels, fallbackFont) {
  const text = aR['a:t'];
  let textValue = '';
  if (typeof text === 'string') textValue = text;
  else if (text && typeof text === 'object' && '#text' in text) textValue = text['#text'];

  const formatting = extractRunFormatting(aR['a:rPr']) || {};

  // Apply fallback size when the run carries no explicit sz
  if (fallbackSize && !formatting.size) {
    formatting.size = fallbackSize;
  }

  // Inherit latin font family from txStyles when the run has no explicit <a:latin>
  if (fallbackFont && !formatting.font) {
    formatting.font = fallbackFont;
  }

  // Inherit bold/italic/underline from defRPr cascade when not set on the run.
  // Respect explicit "not bold/italic/underline" (b="0", i="0", u="none") on the run's rPr.
  if (fallbackBIU) {
    const rPr = aR['a:rPr'];
    if (fallbackBIU.weight != null && formatting.weight == null) {
      const b = rPr && rPr['@_b'];
      if (b !== '0' && b !== 'false') formatting.weight = fallbackBIU.weight;
    }
    if (fallbackBIU.italics != null && formatting.italics == null) {
      const i = rPr && rPr['@_i'];
      if (i !== '0' && i !== 'false') formatting.italics = fallbackBIU.italics;
    }
    if (fallbackBIU['text-decoration'] != null && formatting['text-decoration'] == null) {
      const u = rPr && rPr['@_u'];
      if (u !== 'none') formatting['text-decoration'] = fallbackBIU['text-decoration'];
    }
  }

  const run = { text: textValue };
  if (Object.keys(formatting).length > 0) run.formatting = formatting;

  // Hyperlink: <a:rPr><a:hlinkClick r:id="rId2"/></a:rPr>
  const rPr = aR['a:rPr'];
  const hlinkClick = rPr && rPr['a:hlinkClick'];
  if (hlinkClick && slideRels) {
    const rId  = hlinkClick['@_r:id'];
    const rel  = rId && slideRels[rId];
    const href = rel && rel.target;
    // Only include absolute URLs (http, https, mailto, ftp, tel …)
    if (href && /^[a-z][a-z0-9+\-.]*:/i.test(href)) {
      run.link = { href, target: '_blank' };
      // Apply the theme hyperlink color when the run has no explicit color,
      // so linked text is visually distinct even without browser default styling.
      if (!formatting.color) {
        if (!run.formatting) run.formatting = {};
        run.formatting.color = 'var(--theme-link)';
      }
    }
  }

  return run;
}

/**
 * Convert a single <a:p> to an IR paragraph.
 * lstStyle:       the slide shape's <a:lstStyle> node (may be null)
 * phType:         PPTX placeholder type string (e.g. 'body', 'title') or null
 * layoutLstStyle: the layout/master placeholder's <a:lstStyle> node — sits
 *                 between the slide's own lstStyle and the master txStyles in
 *                 the OOXML BIU inheritance chain. May be null.
 */
// skipMasterSizeFallbacks: when true, skip steps 3 & 4 of the size cascade
// (sizeFromTxStyles and placeholderFallbackSize). Used for embedded shape text
// so that the master's body-text size (e.g. 22pt) is not applied to shape
// labels that have no explicit run size — those labels stay unsized and the
// browser uses its default, which is far more appropriate than a slide heading.
function paragraphToIr(aP, idx, lstStyle, phType, txStyles, layoutLstStyle, slideRels, skipMasterSizeFallbacks) {
  // Determine indent level for size fallback lookup
  const indentLevel = aP['a:pPr'] ? (Number(aP['a:pPr']['@_lvl']) || 0) : 0;

  // Build fallback size cascade:
  //   1. Paragraph-level <a:pPr><a:defRPr sz="...">
  //   2. Shape's <a:lstStyle> for this indent level
  //   3. Master txStyles per placeholder type and indent level  [skipped for shape text]
  //   4. PPTX-spec hard-coded defaults per placeholder type     [skipped for shape text]
  let fallbackSize = null;
  const paraDefRPr = aP['a:pPr'] && aP['a:pPr']['a:defRPr'];
  if (paraDefRPr && paraDefRPr['@_sz']) {
    const pt = Number(paraDefRPr['@_sz']) / 100;
    if (!Number.isNaN(pt)) fallbackSize = `${pt}pt`;
  }
  if (!fallbackSize) fallbackSize = sizeFromLstStyle(lstStyle, indentLevel);
  if (!fallbackSize) fallbackSize = sizeFromLstStyle(layoutLstStyle, indentLevel);
  if (!fallbackSize && !skipMasterSizeFallbacks) fallbackSize = sizeFromTxStyles(txStyles, phType, indentLevel);
  if (!fallbackSize && !skipMasterSizeFallbacks) fallbackSize = placeholderFallbackSize(phType, indentLevel);

  // Font family fallback: applied to runs with no explicit <a:latin>, same gating
  // as size (skipped for shape text when skipMasterSizeFallbacks is set — but font
  // must always be propagated for shapes so the browser doesn't use a wider default).
  const fallbackFont = fontFromTxStyles(txStyles, phType, indentLevel);

  // Build fallback bold/italic/underline from the four-level defRPr cascade
  // (OOXML inheritance order, lowest → highest priority):
  //   1. Master txStyles defRPr
  //   2. Layout/master placeholder <a:lstStyle> defRPr  (layoutLstStyle)
  //   3. Slide shape's own <a:lstStyle> defRPr
  //   4. Paragraph's own <a:pPr><a:defRPr>
  // Each level's explicit "not" (b="0", u="none") clears what lower levels set.
  const txStylesBIU     = biuFromTxStyles(txStyles, phType, indentLevel);
  const layoutLstDefRPr = lstStyleDefRPr(layoutLstStyle, indentLevel);
  const lstDefRPr       = lstStyleDefRPr(lstStyle, indentLevel);

  const fallbackBIU = Object.assign({}, txStylesBIU);

  Object.assign(fallbackBIU, extractBIU(layoutLstDefRPr));
  if (layoutLstDefRPr) {
    if (layoutLstDefRPr['@_b'] === '0' || layoutLstDefRPr['@_b'] === 'false') delete fallbackBIU.weight;
    if (layoutLstDefRPr['@_i'] === '0' || layoutLstDefRPr['@_i'] === 'false') delete fallbackBIU.italics;
    if (layoutLstDefRPr['@_u'] === 'none') delete fallbackBIU['text-decoration'];
  }

  Object.assign(fallbackBIU, extractBIU(lstDefRPr));
  if (lstDefRPr) {
    if (lstDefRPr['@_b'] === '0' || lstDefRPr['@_b'] === 'false') delete fallbackBIU.weight;
    if (lstDefRPr['@_i'] === '0' || lstDefRPr['@_i'] === 'false') delete fallbackBIU.italics;
    if (lstDefRPr['@_u'] === 'none') delete fallbackBIU['text-decoration'];
  }

  Object.assign(fallbackBIU, extractBIU(paraDefRPr));
  if (paraDefRPr) {
    if (paraDefRPr['@_b'] === '0' || paraDefRPr['@_b'] === 'false') delete fallbackBIU.weight;
    if (paraDefRPr['@_i'] === '0' || paraDefRPr['@_i'] === 'false') delete fallbackBIU.italics;
    if (paraDefRPr['@_u'] === 'none') delete fallbackBIU['text-decoration'];
  }

  // Parse explicit tab stops from <a:pPr><a:tabLst><a:tab l="..." algn="..."/>.
  // fast-xml-parser may return a single object or an array; asArray normalises both.
  // Absent or empty <a:tabLst> → tabStops stays empty; the generator uses a fallback.
  const paraPPr = aP['a:pPr'];
  const tabLst  = paraPPr && paraPPr['a:tabLst'];
  const tabStops = asArray(tabLst && tabLst['a:tab'])
    .filter((t) => t && t['@_l'] != null)
    .map((t) => ({ pos: Number(t['@_l']), align: t['@_algn'] || 'l' }))
    .sort((a, b) => a.pos - b.pos);

  // Build runs from <a:r> elements, splitting on literal U+0009 tab characters.
  // PowerPoint sometimes writes "Text\tMore" as a single <a:t> instead of using
  // separate <a:r> and <a:tab/> elements.  Split here so the generator receives
  // discrete { type:'tab' } markers that it can render as positioned spacers.
  // Each split part gets an independent copy of the formatting object so that
  // any post-processing mutation (e.g. normAutofit scaling) affects only one run.
  const runs = [];
  for (const aR of asArray(aP['a:r'])) {
    const run = runToIr(aR, fallbackSize, fallbackBIU, slideRels, fallbackFont);
    if (run.text && run.text.includes('\t')) {
      const parts = run.text.split('\t');
      for (let pi = 0; pi < parts.length; pi++) {
        if (pi > 0) runs.push({ type: 'tab' });
        runs.push({
          ...run,
          text: parts[pi],
          formatting: run.formatting ? { ...run.formatting } : undefined,
        });
      }
    } else {
      runs.push(run);
    }
  }

  // <a:fld> (field elements: slide number, date/time) share the same child
  // structure as <a:r> (<a:rPr> + <a:t>), so runToIr handles them directly.
  // Fields always contain their pre-computed text in <a:t>, so no index lookup
  // is needed — PowerPoint already wrote the correct value there.
  for (const aFld of asArray(aP['a:fld'])) {
    runs.push(runToIr(aFld, fallbackSize, fallbackBIU, slideRels, fallbackFont));
  }

  // Handle <a:tab/> XML elements: some PPTX producers (e.g. LibreOffice) place
  // a self-closing <a:tab/> between sibling <a:r> elements rather than embedding
  // a \t character.  fast-xml-parser does not preserve inter-element order across
  // different tag names, so we insert these as a best-effort heuristic: one tab
  // marker per <a:tab/> element, inserted before the last run in the paragraph.
  // Skip if \t-splitting already produced tab markers (both methods in one para
  // is extremely rare and merging them would duplicate markers).
  const tabNodes = asArray(aP['a:tab']);
  const hasTabsFromSplit = runs.some((r) => r.type === 'tab');
  if (tabNodes.length > 0 && !hasTabsFromSplit) {
    const insertAt = Math.max(1, runs.length - tabNodes.length);
    for (let ti = 0; ti < tabNodes.length; ti++) {
      runs.splice(insertAt + ti, 0, { type: 'tab' });
    }
  }

  // Track empty paragraphs (no runs after all run-building logic).
  // The font size is resolved here so the generator can render the blank line at
  // the correct height.  We use the same cascade as runs but stop before the
  // hard-coded placeholder defaults — those should not force a size on blank lines
  // in test fixtures that carry no explicit sizing.
  const isEmptyPara = runs.length === 0;
  let emptyParaSize = null;
  if (isEmptyPara) {
    const endParaRPr = aP['a:endParaRPr'];
    if (endParaRPr && endParaRPr['@_sz'] != null) {
      const pt = Number(endParaRPr['@_sz']) / 100;
      if (!Number.isNaN(pt)) emptyParaSize = `${pt}pt`;
    }
    if (!emptyParaSize) emptyParaSize = sizeFromLstStyle(lstStyle, indentLevel);
    if (!emptyParaSize) emptyParaSize = sizeFromLstStyle(layoutLstStyle, indentLevel);
    if (!emptyParaSize && !skipMasterSizeFallbacks) {
      emptyParaSize = sizeFromTxStyles(txStyles, phType, indentLevel);
    }
  }

  // Explicit formatting from this paragraph's own <a:pPr>
  const explicitFormatting = extractParagraphFormatting(aP['a:pPr']);

  // Spacing cascade: paragraph pPr (explicit) > layout lstStyle > master txStyles.
  //
  // When the layout's lstStyle defines an explicit font size for this level it
  // signals that the placeholder has been purposefully customised (e.g. a compact
  // footer-annotation box).  In that case the layout's own spacing takes full
  // precedence — the master body defaults (which are sized for 22pt body text)
  // must not bleed into a 9pt compact box.  When the layout lstStyle has no size
  // override, fall through to the master txStyles as before.
  const layoutHasExplicitSize = sizeFromLstStyle(layoutLstStyle, indentLevel) !== null;
  const txSpacing = layoutHasExplicitSize
    ? spacingFromLstStyle(layoutLstStyle, indentLevel)   // may be {} → no spacing
    : spacingFromTxStyles(txStyles, phType, indentLevel);

  let formatting = explicitFormatting ? { ...explicitFormatting } : null;
  if (txSpacing) {
    if (!formatting) formatting = {};
    for (const [key, val] of Object.entries(txSpacing)) {
      if (formatting[key] == null) formatting[key] = val;
    }
    if (Object.keys(formatting).length === 0) formatting = null;
  }

  // Resolve bullet type from lstStyle cascade when the paragraph has no explicit
  // bullet node in its own <a:pPr>.  Body placeholders typically inherit bullets
  // from the shape's txBody lstStyle rather than setting them per-paragraph.
  if (!formatting || !formatting['list-type']) {
    const pPr = aP['a:pPr'];
    const hasExplicitBu = pPr && (
      pPr['a:buNone']    !== undefined ||
      pPr['a:buChar']    !== undefined ||
      pPr['a:buAutoNum'] !== undefined
    );
    if (!hasExplicitBu) {
      const bulletInfo =
        bulletFromLstStyle(lstStyle, indentLevel) ||
        bulletFromLstStyle(layoutLstStyle, indentLevel) ||
        bulletFromTxStyles(txStyles, phType, indentLevel);
      if (bulletInfo && bulletInfo.type !== 'none') {
        if (!formatting) formatting = {};
        formatting['list-type'] = bulletInfo.type;
        if (bulletInfo.char) formatting['bullet-char'] = bulletInfo.char;
      }
    }
  }

  // For empty paragraphs, attach the resolved font size as formatting.size so
  // the generator can render a blank line whose height matches the surrounding text.
  if (isEmptyPara && emptyParaSize) {
    if (!formatting) formatting = {};
    if (!formatting.size) formatting.size = emptyParaSize;
  }

  const ir = { id: 'p-' + idx, runs };
  if (formatting) ir.formatting = formatting;
  if (tabStops.length > 0) ir.tabStops = tabStops;
  return ir;
}

/**
 * Convert a <p:sp> shape that carries text into an IR text block.
 * Returns null if the shape has no text body or is a skipped metadata placeholder.
 *
 * layoutLstStyle: the layout/master placeholder's <a:lstStyle> node, pre-resolved
 *   by slide.js from loadLayoutGeometry. Null for non-placeholder shapes.
 */
function shapeToTextBlock(pSp, idx, txStyles, layoutLstStyle, slideRels, skipMasterSizeFallbacks) {
  const txBody = pSp['p:txBody'];
  if (!txBody) return null;

  const ph = pSp['p:nvSpPr']
    && pSp['p:nvSpPr']['p:nvPr']
    && pSp['p:nvSpPr']['p:nvPr']['p:ph'];
  const phType = ph ? ph['@_type'] : null;

  // Skip header, date/time, and slide-number placeholders entirely.
  if (ph && SKIP_PLACEHOLDER_TYPES.has(phType)) return null;

  // Shape's <a:lstStyle> provides per-level size fallbacks for runs.
  const lstStyle = txBody['a:lstStyle'];

  const paragraphs = asArray(txBody['a:p']).map((aP, pIdx) =>
    paragraphToIr(aP, pIdx, lstStyle, phType, txStyles, layoutLstStyle || null, slideRels || null, !!skipMasterSizeFallbacks)
  );
  if (paragraphs.length === 0) return null;

  const block = {
    id: 'txt-' + idx,
    paragraphs,
  };

  // Position from <p:spPr><a:xfrm>
  const xfrm = pSp['p:spPr'] && pSp['p:spPr']['a:xfrm'];
  if (xfrm) {
    if (xfrm['a:off']) {
      block.position = {
        x: emuToPx(xfrm['a:off']['@_x']) || 0,
        y: emuToPx(xfrm['a:off']['@_y']) || 0,
      };
    }
    if (xfrm['a:ext']) {
      const w = emuToPx(xfrm['a:ext']['@_cx']);
      const h = emuToPx(xfrm['a:ext']['@_cy']);
      if (w != null) block.width = w;
      if (h != null) block.height = h;
    }
    if (xfrm['@_rot']) {
      const deg = pptxRotationToDegrees(xfrm['@_rot']);
      if (deg !== 0) block.rotation = deg;
    }
  }

  // Footer placeholder without explicit position (position lives in slide layout).
  // Mark it so the generator can place it at the bottom of the slide canvas.
  if (phType === 'ftr' && !block.position) {
    block['footer-placement'] = true;
  }

  // <a:bodyPr> — read vertical anchor, auto-fit mode, and normAutofit from the same node.
  // Anchor fallback (layout/master) is resolved later in slide.js.
  const bodyPr = txBody['a:bodyPr'];
  if (bodyPr && bodyPr['@_anchor']) {
    block['text-anchor'] = bodyPr['@_anchor'];
  }

  // Auto-fit mode — stored as textBlock.autoFit so the generator only applies
  // overflow / shrink behaviour when PowerPoint actually requested it.
  //   'none'  → <a:noAutofit/>   text overflows and is clipped
  //   'norm'  → <a:normAutofit/> font/spacing scaled to fit (fontScale applied below)
  //   'shape' → <a:spAutoFit/>   shape grew to contain text; no fixed bounding box
  // Absent means the PPTX did not specify, treated as 'none' by PowerPoint.
  if (bodyPr) {
    if      (bodyPr['a:noAutofit']  !== undefined) block.autoFit = 'none';
    else if (bodyPr['a:normAutofit'] !== undefined) block.autoFit = 'norm';
    else if (bodyPr['a:spAutoFit']   !== undefined) block.autoFit = 'shape';
  }

  // <a:bodyPr><a:normAutofit> — PowerPoint records how much it scaled fonts/spacing
  // to auto-fit overflowing text. fontScale and lnSpcReduction are in 1000ths of a
  // percent (100000 = 100%).  Apply them now so the HTML matches what PPT shows.
  const normAutofit = bodyPr && bodyPr['a:normAutofit'];
  if (normAutofit) {
    const fontScaleRaw = normAutofit['@_fontScale'];
    const lnSpcRedRaw  = normAutofit['@_lnSpcReduction'];

    const fontScale = fontScaleRaw != null ? Number(fontScaleRaw) / 100000 : 1;
    const lnSpcRed  = lnSpcRedRaw  != null ? Number(lnSpcRedRaw)  / 100000 : 0;

    for (const para of block.paragraphs) {
      // Scale run font sizes
      if (!Number.isNaN(fontScale) && fontScale !== 1 && fontScale > 0) {
        for (const run of para.runs) {
          const sz = run.formatting && run.formatting.size;
          if (sz && typeof sz === 'string' && sz.endsWith('pt')) {
            const scaled = parseFloat(sz) * fontScale;
            run.formatting.size = `${Math.round(scaled * 100) / 100}pt`;
          }
        }
      }

      // Reduce line spacing: multiply the existing spacing by (1 - lnSpcRed)
      // so that lnSpcReduction=0.2 means "20% reduction" regardless of the base.
      if (!Number.isNaN(lnSpcRed) && lnSpcRed > 0) {
        const f = para.formatting;
        const ls = f && f['line-spacing'];
        if (ls && typeof ls === 'string') {
          const unitless = parseFloat(ls);
          if (!Number.isNaN(unitless) && !ls.endsWith('pt')) {
            const reduced = parseFloat(Math.max(0.5, unitless * (1 - lnSpcRed)).toFixed(4));
            para.formatting['line-spacing'] = String(reduced);
          }
        } else {
          // No explicit line spacing — reduce from the PPTX default of 1.0
          if (!para.formatting) para.formatting = {};
          para.formatting['line-spacing'] = String(Math.max(0.5, 1.0 - lnSpcRed));
        }
      }
    }
    // Flag so slide.js does not double-apply layout/master normAutofit.
    block._normAutofitApplied = true;
  }

  return block;
}

module.exports = { shapeToTextBlock, paragraphToIr, runToIr, fontFromTxStyles };
