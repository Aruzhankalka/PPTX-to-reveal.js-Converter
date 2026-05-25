const { asArray } = require('./xml');
const { emuToPx, pptxRotationToDegrees } = require('./units');

/**
 * Extract a run's formatting from <a:rPr> attributes.
 * Sprint 1 covers the FR-06 subset: bold, italics, underline/strikethrough,
 * color, font family, font size.
 */
function extractRunFormatting(rPr) {
  if (!rPr) return undefined;
  const f = {};

  if (rPr['@_b'] === '1') f.weight = 'bold';
  if (rPr['@_i'] === '1') f.italics = true;
  if (rPr['@_u'] && rPr['@_u'] !== 'none') f['text-decoration'] = 'underline';
  if (rPr['@_strike'] && rPr['@_strike'] !== 'noStrike') {
    f['text-decoration'] = 'strikethrough';
  }
  // <a:rPr sz="2400"> means 24pt — sz is in 100ths of a point.
  if (rPr['@_sz']) {
    const pt = Number(rPr['@_sz']) / 100;
    if (!Number.isNaN(pt)) f.size = pt + 'pt';
  }
  // Font color: explicit hex or theme slot reference
  // <a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>
  // <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>
  const fill = rPr['a:solidFill'];
  if (fill) {
    if (fill['a:srgbClr'] && fill['a:srgbClr']['@_val']) {
      f.color = '#' + fill['a:srgbClr']['@_val'];
    } else if (fill['a:schemeClr'] && fill['a:schemeClr']['@_val']) {
      f.color = `var(--theme-${fill['a:schemeClr']['@_val']})`;
    }
  }
  // Font family: <a:latin typeface="Arial"/>
  const latin = rPr['a:latin'];
  if (latin && latin['@_typeface']) f.font = latin['@_typeface'];

  return Object.keys(f).length > 0 ? f : undefined;
}

/**
 * Extract a paragraph's formatting from <a:pPr>.
 */
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

  // List type: <a:buChar/> for bullets, <a:buAutoNum/> for numbered, <a:buNone/> for none
  if (pPr['a:buNone'] !== undefined) f['list-type'] = 'none';
  else if (pPr['a:buAutoNum'] !== undefined) f['list-type'] = 'numbered';
  else if (pPr['a:buChar'] !== undefined) f['list-type'] = 'bullets';

  return Object.keys(f).length > 0 ? f : undefined;
}

/**
 * Convert a single <a:r> to an IR run.
 * <a:r><a:rPr .../><a:t>text</a:t></a:r>
 */
function runToIr(aR) {
  const text = aR['a:t'];
  // <a:t> may be a string, or an object like { '#text': 'value' } depending on
  // whitespace and parser quirks. Normalise.
  let textValue = '';
  if (typeof text === 'string') textValue = text;
  else if (text && typeof text === 'object' && '#text' in text) textValue = text['#text'];

  const formatting = extractRunFormatting(aR['a:rPr']);
  const run = { text: textValue };
  if (formatting) run.formatting = formatting;
  return run;
}

/**
 * Convert a single <a:p> to an IR paragraph.
 * Skips empty paragraphs that contain no runs.
 */
function paragraphToIr(aP, idx) {
  const runs = asArray(aP['a:r']).map(runToIr);

  // Some paragraphs use <a:fld> (fields) or <a:br> (line breaks). For Sprint 1
  // we keep it simple and ignore those; they'll be added in Sprint 2.

  if (runs.length === 0) {
    // Preserve empty paragraph as a single empty run so the slide structure
    // stays intact (otherwise you lose blank lines).
    runs.push({ text: '' });
  }

  const formatting = extractParagraphFormatting(aP['a:pPr']);
  const ir = { id: 'p-' + idx, runs };
  if (formatting) ir.formatting = formatting;
  return ir;
}

/**
 * Convert a <p:sp> shape that carries text into an IR text block.
 * Returns null if the shape has no text body (it's a pure shape, handled
 * by Sprint 2's shapes module).
 */
function shapeToTextBlock(pSp, idx) {
  const txBody = pSp['p:txBody'];
  if (!txBody) return null;

  const paragraphs = asArray(txBody['a:p']).map(paragraphToIr);
  if (paragraphs.length === 0) return null;

  const block = {
    id: 'txt-' + idx,
    paragraphs,
  };

  // Position from <p:spPr><a:xfrm><a:off/><a:ext/></a:xfrm></p:spPr>
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

  return block;
}

module.exports = { shapeToTextBlock, paragraphToIr, runToIr };