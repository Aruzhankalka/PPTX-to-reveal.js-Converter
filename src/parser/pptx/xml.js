const { XMLParser } = require('fast-xml-parser');

const PARSER_OPTIONS = {
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  removeNSPrefix: false,
};

/**
 * Shared XML parser configuration.
 *
 * - ignoreAttributes: false   -> keep attributes (we need them for formatting)
 * - attributeNamePrefix: '@_' -> distinguishes attributes from child nodes
 * - parseTagValue: false      -> never coerce strings to numbers (preserves leading zeros, etc.)
 * - parseAttributeValue: false -> same, for attributes
 * - trimValues: false         -> preserve whitespace in <a:t>...</a:t> exactly
 *
 * Per Specification §6 (NFR-02 row): parseTagValue=false avoids type-coercion overhead.
 */
const parser = new XMLParser(PARSER_OPTIONS);

/**
 * A second parser instance with preserveOrder:true, used only to recover
 * the interleaved document order of p:sp and p:pic children inside spTree.
 * The main parser groups same-tag siblings, losing their relative order.
 */
const orderedParser = new XMLParser({ ...PARSER_OPTIONS, preserveOrder: true });

function parseXml(xmlString) {
  if (!xmlString) return null;
  return parser.parse(xmlString);
}

/**
 * fast-xml-parser collapses single-child arrays to objects. Most slide content
 * is a list (paragraphs, runs, shapes), so we need this normalizer everywhere.
 */
function asArray(value) {
  if (value == null) return [];
  return Array.isArray(value) ? value : [value];
}

/**
 * Walk an ordered-parser node array to find the first child with a given tag.
 * In preserveOrder output each element is { tagName: [...children], ':@': {...attrs} }.
 */
function findOrderedChild(nodes, tag) {
  if (!Array.isArray(nodes)) return null;
  for (const node of nodes) {
    if (node[tag]) return node[tag];
  }
  return null;
}

/**
 * Return the document-order sequence of direct p:spTree children as an array
 * of { tag, idx } objects, where idx is the 0-based count of that tag seen so
 * far (matching the array indices produced by the main parser).
 *
 * Only p:sp, p:pic, and p:grpSp are included; boilerplate children like p:nvGrpSpPr
 * are skipped. p:grpSp entries are emitted so the caller can assign a z-index
 * placeholder for the group's contained pictures.
 *
 * @param {string} slideXml - raw slide XML string
 * @returns {Array<{tag: string, idx: number}>}
 */
function getSpTreeChildOrder(slideXml) {
  if (!slideXml) return [];

  const doc = orderedParser.parse(slideXml);

  const sld   = findOrderedChild(doc, 'p:sld');
  if (!sld) return [];
  const cSld  = findOrderedChild(sld, 'p:cSld');
  if (!cSld) return [];
  const spTree = findOrderedChild(cSld, 'p:spTree');
  if (!spTree) return [];

  const TRACKED = new Set(['p:sp', 'p:pic', 'p:grpSp', 'p:cxnSp']);
  const counters = {};
  const order = [];

  for (const child of spTree) {
    const tag = Object.keys(child).find(k => k !== ':@');
    if (!tag || !TRACKED.has(tag)) continue;
    const idx = counters[tag] ?? 0;
    order.push({ tag, idx });
    counters[tag] = idx + 1;
  }

  return order;
}

/**
 * Generic variant of getSpTreeChildOrder that works for any root element
 * (slide, layout, master).  Pass the raw XML and the root tag name.
 *
 * @param {string} rawXml   - raw XML string (slide, layout, or master)
 * @param {string} rootTag  - root element name, e.g. 'p:sld', 'p:sldLayout', 'p:sldMaster'
 * @returns {Array<{tag: string, idx: number}>}
 */
function getSpTreeOrder(rawXml, rootTag) {
  if (!rawXml) return [];
  const doc = orderedParser.parse(rawXml);
  const root = findOrderedChild(doc, rootTag);
  if (!root) return [];
  const cSld = findOrderedChild(root, 'p:cSld');
  if (!cSld) return [];
  const spTree = findOrderedChild(cSld, 'p:spTree');
  if (!spTree) return [];

  const TRACKED = new Set(['p:sp', 'p:pic', 'p:grpSp', 'p:cxnSp']);
  const counters = {};
  const order = [];
  for (const child of spTree) {
    const tag = Object.keys(child).find((k) => k !== ':@');
    if (!tag || !TRACKED.has(tag)) continue;
    const idx = counters[tag] ?? 0;
    order.push({ tag, idx });
    counters[tag] = idx + 1;
  }
  return order;
}

module.exports = { parseXml, asArray, getSpTreeChildOrder, getSpTreeOrder };