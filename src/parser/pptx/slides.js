const { readText } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { emuToPx } = require('./units');
const { resolveSolidFillCss } = require('./color');

/**
 * Read presentation.xml + its .rels to determine the ordered list of slide files.
 *
 * Returns an array of { rId, path } in slide order, where path is e.g.
 * 'ppt/slides/slide1.xml'. The order comes from <p:sldIdLst> in presentation.xml,
 * NOT from filename sort — slide7.xml may appear before slide2.xml if the user
 * reordered slides in PowerPoint.
 */
async function listSlides(zip) {
  const presentationXml = await readText(zip, 'ppt/presentation.xml');
  if (!presentationXml) {
    const e = new Error('Missing ppt/presentation.xml — not a valid PPTX');
    e.code = 'INVALID_PPTX';
    throw e;
  }

  const presRelsXml = await readText(zip, 'ppt/_rels/presentation.xml.rels');
  const presRels = parseRelationships(presRelsXml);

  const parsed = parseXml(presentationXml);
  const sldIdLst = parsed
    && parsed['p:presentation']
    && parsed['p:presentation']['p:sldIdLst'];
  if (!sldIdLst) return [];

  const result = [];
  for (const sldId of asArray(sldIdLst['p:sldId'])) {
    const rId = sldId['@_r:id'];
    if (!rId) continue;
    const rel = presRels[rId];
    if (!rel) continue;
    const path = resolveTarget('ppt', rel.target);
    result.push({ rId, path });
  }
  return result;
}

/**
 * Read <p:sldSz> from presentation.xml and return slide canvas dimensions in px.
 * Standard widescreen PPTX is 9144000 × 5143500 EMU → 960 × 540 px.
 * Returns null values when the element is absent (non-standard files).
 */
async function getSlideDimensions(zip) {
  const xml = await readText(zip, 'ppt/presentation.xml');
  if (!xml) return { slideWidth: null, slideHeight: null };

  const parsed = parseXml(xml);
  const sldSz = parsed && parsed['p:presentation'] && parsed['p:presentation']['p:sldSz'];

  return {
    slideWidth:  emuToPx(sldSz && sldSz['@_cx']),
    slideHeight: emuToPx(sldSz && sldSz['@_cy']),
  };
}

/**
 * Convert a parsed <a:lvl1pPr> node (default-run-properties + paragraph
 * defaults for indent level 1) into the IR's flat formatting-bag shape —
 * the same shape used by paragraph.formatting, run.formatting, and
 * master.formatting.
 *
 * Reused for two distinct sources that share this exact XML shape:
 *   - presentation.xml's <p:defaultTextStyle> (-> master.formatting)
 *   - a layout placeholder's own <a:lstStyle> (-> layouts[].placeholders[].formatting)
 *
 * @param {object|null|undefined} lvl1pPr - parsed <a:lvl1pPr> node
 * @returns {object|null} IR-shaped formatting bag, or null when there are no
 *   usable fields (absent node, or an empty/no-op pPr).
 */
function lvl1pPrToFormatting(lvl1pPr) {
  if (!lvl1pPr) return null;

  const f = {};

  // Run-level defaults from <a:defRPr> (font, size, weight, italics,
  // text-decoration, color) — same fields/units as text.js's run formatting.
  const defRPr = lvl1pPr['a:defRPr'];
  if (defRPr) {
    if (defRPr['@_sz']) {
      const pt = Number(defRPr['@_sz']) / 100;
      if (!Number.isNaN(pt)) f.size = `${pt}pt`;
    }
    const b = defRPr['@_b'];
    f.weight = (b === '1' || b === 'true') ? 'bold' : 'normal';
    const i = defRPr['@_i'];
    f.italics = (i === '1' || i === 'true');
    const u = defRPr['@_u'];
    const strike = defRPr['@_strike'];
    if (u && u !== 'none') f['text-decoration'] = 'underline';
    else if (strike && strike !== 'noStrike') f['text-decoration'] = 'strikethrough';

    const latin = defRPr['a:latin'];
    if (latin && latin['@_typeface']) f.font = latin['@_typeface'];

    const fillColor = resolveSolidFillCss(defRPr['a:solidFill']);
    if (fillColor) f.color = fillColor;
  }

  // Paragraph-level defaults from <a:lvl1pPr> itself (align, list-type, line-spacing).
  const algn = lvl1pPr['@_algn'];
  const algnMap = { l: 'left', r: 'right', ctr: 'center', just: 'justify' };
  if (algn && algnMap[algn]) f.align = algnMap[algn];

  if (lvl1pPr['a:buNone'] !== undefined) f['list-type'] = 'none';
  else if (lvl1pPr['a:buAutoNum'] !== undefined) f['list-type'] = 'numbered';
  else if (lvl1pPr['a:buChar'] !== undefined) f['list-type'] = 'bullets';

  const lnSpc = lvl1pPr['a:lnSpc'];
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

  return Object.keys(f).length > 0 ? f : null;
}

/**
 * Read <p:defaultTextStyle><a:lvl1pPr> from presentation.xml — the
 * presentation-wide default text style (spec: master.formatting, "Global
 * preset!"). It sits above every other level in the formatting inheritance
 * chain: master.formatting -> slide master txStyles -> layout/slide
 * lstStyle -> paragraph -> run.
 *
 * Only level 1 is read: unlike txStyles (one entry per indent level, used
 * internally to backfill missing run sizes), master.formatting is a single
 * flat style object — the same shape as paragraph/run formatting.
 *
 * @param {JSZip} zip
 * @returns {Promise<object|null>} IR-shaped formatting bag, or null when
 *   presentation.xml carries no <p:defaultTextStyle> (rare/malformed) or it
 *   resolves to no usable fields.
 */
async function getDefaultTextStyle(zip) {
  const xml = await readText(zip, 'ppt/presentation.xml');
  if (!xml) return null;

  const parsed = parseXml(xml);
  const dts = parsed
    && parsed['p:presentation']
    && parsed['p:presentation']['p:defaultTextStyle'];

  return lvl1pPrToFormatting(dts && dts['a:lvl1pPr']);
}

module.exports = { listSlides, getSlideDimensions, getDefaultTextStyle, lvl1pPrToFormatting };