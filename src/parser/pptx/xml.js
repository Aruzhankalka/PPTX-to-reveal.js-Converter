/**
 * Shared XML parsing utilities for the whole PPTX parser: the configured
 * fast-xml-parser instance every module.js uses to turn a raw XML string
 * into a plain object tree, plus helpers for fast-xml-parser's two quirks
 * that the rest of the parser has to work around (single-child arrays
 * collapse to bare objects; same-tag siblings lose their relative order).
 */

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

/**
 * Parse an XML string with the shared parser config (see PARSER_OPTIONS above).
 * @param {string} xmlString - raw XML; falsy input (e.g. a missing/empty part) is valid input
 * @returns {object|null} parsed document tree, or null when xmlString is falsy
 */
function parseXml(xmlString) {
  if (!xmlString) return null;
  return parser.parse(xmlString);
}

/**
 * fast-xml-parser collapses single-child arrays to objects. Most slide content
 * is a list (paragraphs, runs, shapes), so we need this normalizer everywhere.
 *
 * @param {*} value - a parsed node, array of nodes, or null/undefined
 * @returns {Array} value as-is if already an array, [value] if a single node,
 *   [] if value is null/undefined
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
 * Works for any root element (slide, layout, master) — pass the raw XML and
 * the root tag name, e.g. 'p:sld', 'p:sldLayout', 'p:sldMaster'.
 *
 * @param {string} rawXml   - raw XML string (slide, layout, or master)
 * @param {string} rootTag  - root element name
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

  const TRACKED = new Set(['p:sp', 'p:pic', 'p:grpSp', 'p:cxnSp', 'p:graphicFrame']);
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

module.exports = { parseXml, asArray, getSpTreeOrder };