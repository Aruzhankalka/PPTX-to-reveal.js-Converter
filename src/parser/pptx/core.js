/**
 * Presentation core-properties parser — reads docProps/core.xml for
 * document-level metadata (title/author/creation date) surfaced in the
 * IR's slideset object, independent of any slide/shape content.
 */

const { readText } = require('./zip');
const { parseXml } = require('./xml');

/**
 * Parse docProps/core.xml from the PPTX zip to extract presentation metadata.
 *
 * @param {JSZip} zip - open PPTX archive
 * @returns {Promise<{title: string|null, author: string|null, creationDate: string|null}>}
 *   any field is null if absent. creationDate is ISO date (YYYY-MM-DD), trimmed
 *   from the full W3CDTF datetime (e.g. "2021-02-15T08:56:04Z" → "2021-02-15").
 */
async function parseCoreProps(zip) {
  const xml = await readText(zip, 'docProps/core.xml');
  if (!xml) return { title: null, author: null, creationDate: null };

  const parsed = parseXml(xml);
  const core = parsed && parsed['cp:coreProperties'];
  if (!core) return { title: null, author: null, creationDate: null };

  const title  = stringVal(core['dc:title'])   || null;
  const author = stringVal(core['dc:creator']) || null;

  const rawDate = stringVal(core['dcterms:created']);
  const creationDate = rawDate ? rawDate.slice(0, 10) : null;

  return { title, author, creationDate };
}

// A node can be a plain string (text-only) or an object with '#text' (when it
// also carries attributes like xsi:type). Handle both.
function stringVal(node) {
  if (!node) return null;
  if (typeof node === 'string') return node.trim() || null;
  if (typeof node === 'object' && node['#text']) return String(node['#text']).trim() || null;
  return null;
}

module.exports = { parseCoreProps };
