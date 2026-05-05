const { XMLParser } = require('fast-xml-parser');

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
const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: false,
  // OOXML uses namespaces; we keep the prefixed names so we can distinguish
  // p:sp (shape) from a:sp etc.
  removeNSPrefix: false,
});

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

module.exports = { parseXml, asArray };