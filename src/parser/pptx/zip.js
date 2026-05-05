const JSZip = require('jszip');

/**
 * Open a .pptx buffer as a JSZip instance.
 * Throws a parser error if the buffer is not a valid ZIP archive.
 */
async function openPptx(buffer) {
  try {
    return await JSZip.loadAsync(buffer);
  } catch (err) {
    const e = new Error('File is not a valid ZIP archive');
    e.code = 'INVALID_PPTX';
    throw e;
  }
}

/**
 * Read a file inside the .pptx as a UTF-8 string.
 * Returns null if the file does not exist (some parts are optional).
 */
async function readText(zip, path) {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async('string');
}

/**
 * Read a file inside the .pptx as a Buffer (used for media).
 */
async function readBinary(zip, path) {
  const entry = zip.file(path);
  if (!entry) return null;
  return entry.async('nodebuffer');
}

/**
 * List all entries whose name matches a prefix (e.g. "ppt/slides/").
 */
function listByPrefix(zip, prefix) {
  return Object.keys(zip.files).filter((name) => name.startsWith(prefix));
}

module.exports = { openPptx, readText, readBinary, listByPrefix };