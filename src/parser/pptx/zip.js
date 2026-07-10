/**
 * ZIP-level access — a .pptx file is a ZIP archive; every other parser
 * module reads its XML/media parts through these four functions rather
 * than touching JSZip directly, so the "part not found" behavior (return
 * null, not throw) is consistent everywhere.
 */

const JSZip = require('jszip');

/**
 * Open a .pptx buffer as a JSZip instance.
 *
 * @param {Buffer} buffer - raw .pptx file bytes
 * @returns {Promise<JSZip>} opened archive
 * @throws {Error} code 'INVALID_PPTX' when buffer is not a valid ZIP archive
 */
async function openPptx(buffer) {
  try {
    return await JSZip.loadAsync(buffer);
  } catch {
    const e = new Error('File is not a valid ZIP archive');
    e.code = 'INVALID_PPTX';
    throw e;
  }
}

/**
 * Read a file inside the .pptx as a UTF-8 string.
 *
 * @param {JSZip} zip - open PPTX archive
 * @param {string} path - zip-relative path (e.g. 'ppt/slides/slide1.xml')
 * @returns {Promise<string|null>} file contents, or null if the file does
 *   not exist (some parts are optional)
 */
async function readText(zip, path) {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async('string');
}

/**
 * Read a file inside the .pptx as a Buffer (used for media).
 *
 * @param {JSZip} zip - open PPTX archive
 * @param {string} path - zip-relative path (e.g. 'ppt/media/image1.png')
 * @returns {Promise<Buffer|null>} file bytes, or null if the file does not exist
 */
async function readBinary(zip, path) {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async('nodebuffer');
}

/**
 * List all entries whose name matches a prefix (e.g. "ppt/slides/").
 *
 * @param {JSZip} zip - open PPTX archive
 * @param {string} prefix - zip-path prefix to match
 * @returns {string[]} matching zip-relative paths, in JSZip's own iteration order
 */
function listByPrefix(zip, prefix) {
  return Object.keys(zip.files).filter((name) => name.startsWith(prefix));
}

module.exports = { openPptx, readText, readBinary, listByPrefix };