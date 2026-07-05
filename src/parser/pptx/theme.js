const { XMLParser } = require('fast-xml-parser');
const { readText, listByPrefix } = require('./zip');
const { parseXml } = require('./xml');

const ORDERED_PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: false,
  preserveOrder: true,
};
const orderedParser = new XMLParser(ORDERED_PARSER_OPTIONS);

// Find the children array of the first matching tag in an ordered-parser node list.
function findOrderedTag(nodes, tag) {
  for (const item of (nodes || [])) {
    if (Array.isArray(item[tag])) return item[tag];
  }
  return null;
}

/**
 * Convert an ordered-parser node list to the flat object format that the
 * regular parser produces, so callers (resolveFill etc.) work unchanged.
 * Same-tag siblings are collected into an array, matching regular-parser
 * behaviour for lists (a:gs, a:ln, …).
 */
function orderedNodeToRegular(orderedNodes) {
  if (!Array.isArray(orderedNodes)) return {};
  const result = {};
  for (const item of orderedNodes) {
    if (!item || typeof item !== 'object') continue;
    const tag = Object.keys(item).find(k => k !== ':@' && k !== '#text');
    if (!tag) continue;
    const attrs = item[':@'] || {};
    const children = Array.isArray(item[tag]) ? item[tag] : [];
    const childObj = orderedNodeToRegular(children);
    const nodeValue = Object.assign({}, attrs, childObj);
    if (tag in result) {
      if (!Array.isArray(result[tag])) result[tag] = [result[tag]];
      result[tag].push(nodeValue);
    } else {
      result[tag] = nodeValue;
    }
  }
  return result;
}

/**
 * Extract the three format-scheme lists from theme XML using the ordered
 * parser so mixed fill types (solidFill / gradFill) keep their 1-based index.
 *
 * Returns { fillStyleLst, lnStyleLst, effectStyleLst } — each is an array
 * whose N-th element (0-based) corresponds to fmtScheme idx=N+1.
 * Returns null when no <a:fmtScheme> is found.
 */
function extractFmtScheme(themeXml) {
  const doc = orderedParser.parse(themeXml);
  const themeChildren = findOrderedTag(doc, 'a:theme');
  if (!themeChildren) return null;
  const elementsChildren = findOrderedTag(themeChildren, 'a:themeElements');
  if (!elementsChildren) return null;
  const fmtSchemeChildren = findOrderedTag(elementsChildren, 'a:fmtScheme');
  if (!fmtSchemeChildren) return null;

  const fmtScheme = { fillStyleLst: [], lnStyleLst: [], effectStyleLst: [] };

  // fillStyleLst — mixed child types (solidFill, gradFill, …) in document order
  const fillLstChildren = findOrderedTag(fmtSchemeChildren, 'a:fillStyleLst');
  for (const item of (fillLstChildren || [])) {
    const tag = Object.keys(item).find(k => k !== ':@' && k !== '#text');
    if (!tag) continue;
    const attrs = item[':@'] || {};
    const children = Array.isArray(item[tag]) ? item[tag] : [];
    fmtScheme.fillStyleLst.push({ [tag]: Object.assign({}, attrs, orderedNodeToRegular(children)) });
  }

  // lnStyleLst — all a:ln children
  const lnLstChildren = findOrderedTag(fmtSchemeChildren, 'a:lnStyleLst');
  for (const item of (lnLstChildren || [])) {
    const tag = Object.keys(item).find(k => k !== ':@' && k !== '#text');
    if (tag !== 'a:ln') continue;
    const attrs = item[':@'] || {};
    const children = Array.isArray(item[tag]) ? item[tag] : [];
    fmtScheme.lnStyleLst.push({ 'a:ln': Object.assign({}, attrs, orderedNodeToRegular(children)) });
  }

  // effectStyleLst — all a:effectStyle children; store their inner content
  const effectLstChildren = findOrderedTag(fmtSchemeChildren, 'a:effectStyleLst');
  for (const item of (effectLstChildren || [])) {
    const tag = Object.keys(item).find(k => k !== ':@' && k !== '#text');
    if (tag !== 'a:effectStyle') continue;
    const children = Array.isArray(item[tag]) ? item[tag] : [];
    fmtScheme.effectStyleLst.push(orderedNodeToRegular(children));
  }

  return fmtScheme;
}

/**
 * The 12 named colour slots defined by OOXML's <a:clrScheme>.
 * Order matches the spec (ECMA-376 §20.1.6.2).
 */
const COLOR_SLOTS = [
  'dk1', 'lt1', 'dk2', 'lt2',
  'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
  'hlink', 'folHlink',
];

/**
 * Resolve a single <a:clrScheme> slot node to a #RRGGBB string.
 *
 * PPTX colour nodes come in three flavours inside a theme:
 *   <a:srgbClr val="4472C4"/>          → straightforward hex
 *   <a:sysClr val="windowText" lastClr="000000"/>  → OS colour; lastClr is the
 *                                                     saved resolved value
 *   <a:prstClr val="black"/>           → named preset; rare in themes, not mapped
 *
 * Returns null when the colour cannot be resolved so callers can omit the slot
 * rather than emit a broken value.
 */
function resolveColor(slotNode) {
  if (!slotNode) return null;
  if (slotNode['a:srgbClr']) {
    const val = slotNode['a:srgbClr']['@_val'];
    return val ? '#' + val : null;
  }
  if (slotNode['a:sysClr']) {
    const last = slotNode['a:sysClr']['@_lastClr'];
    return last ? '#' + last : null;
  }
  return null;
}

/**
 * Parse a theme XML string into a theme IR object.
 * Exported separately so unit tests can drive it without a ZIP.
 *
 * @param {string} xmlString
 * @returns {{
 *   name: string,
 *   colors: Record<string, string>,
 *   fonts: { major?: string, minor?: string },
 *   fmtScheme: object|null
 * } | null}
 */
function parseThemeXml(xmlString) {
  const parsed = parseXml(xmlString);
  const theme = parsed && parsed['a:theme'];
  if (!theme) return null;

  const elements = theme['a:themeElements'];
  if (!elements) return null;

  const result = {
    name: theme['@_name'] || 'Unknown',
    colors: {},
    fonts: {},
    fmtScheme: null,
  };

  // -- Colour scheme --
  const clrScheme = elements['a:clrScheme'];
  if (clrScheme) {
    for (const slot of COLOR_SLOTS) {
      const color = resolveColor(clrScheme['a:' + slot]);
      if (color) result.colors[slot] = color;
    }
  }

  // -- Font scheme (major = headings, minor = body) --
  const fntScheme = elements['a:fontScheme'];
  if (fntScheme) {
    const majorLatin = fntScheme['a:majorFont'] && fntScheme['a:majorFont']['a:latin'];
    if (majorLatin && majorLatin['@_typeface']) {
      result.fonts.major = majorLatin['@_typeface'];
    }
    const minorLatin = fntScheme['a:minorFont'] && fntScheme['a:minorFont']['a:latin'];
    if (minorLatin && minorLatin['@_typeface']) {
      result.fonts.minor = minorLatin['@_typeface'];
    }
  }

  // Parse fmtScheme (fillStyleLst / lnStyleLst / effectStyleLst) using the
  // ordered parser so mixed fill types keep their 1-based positional index.
  const fmtScheme = extractFmtScheme(xmlString);
  if (fmtScheme) result.fmtScheme = fmtScheme;

  return result;
}

/**
 * Find and parse the theme from an open .pptx ZIP.
 *
 * @param {JSZip} zip
 * @param {string} [themePath] - resolved path from the master's .rels file.
 *   When omitted, falls back to the conventional 'ppt/theme/theme1.xml' path,
 *   then scans ppt/theme/ for any .xml file (handles non-standard filenames).
 *   master.js should always pass the resolved path once FR-11 is implemented.
 * @returns {Promise<object|null>} theme IR object, or null if no theme found.
 */
async function parseTheme(zip, themePath) {
  let xml;

  if (themePath) {
    xml = await readText(zip, themePath);
  } else {
    // Conventional path first
    xml = await readText(zip, 'ppt/theme/theme1.xml');
    if (!xml) {
      // Non-standard filename fallback
      const candidates = listByPrefix(zip, 'ppt/theme/').filter((p) => p.endsWith('.xml'));
      if (candidates.length > 0) {
        xml = await readText(zip, candidates[0]);
      }
    }
  }

  if (!xml) return null;
  return parseThemeXml(xml);
}

module.exports = { parseTheme, parseThemeXml };
