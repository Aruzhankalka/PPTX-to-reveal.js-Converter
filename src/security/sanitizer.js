'use strict';

function logRemoval(type, detail) {
  console.warn(`[SANITIZER] Removed ${type}: ${detail}`);
}

function removeVBA(zip) {
  zip.forEach((relativePath) => {
    if (relativePath.toLowerCase().includes('vbaproject.bin')) {
      zip.remove(relativePath);
      logRemoval('VBA macro', relativePath);
    }
  });
}

async function stripScriptsFromSVGs(zip) {
  const svgFiles = [];
  zip.forEach((relativePath, file) => {
    if (relativePath.toLowerCase().endsWith('.svg')) {
      svgFiles.push({ relativePath, file });
    }
  });

  for (const { relativePath, file } of svgFiles) {
    let content = await file.async('string');
    const original = content;
    content = content.replace(/<script[\s\S]*?<\/script>/gi, '');
    content = content.replace(/\s+on\w+="[^"]*"/gi, '');
    content = content.replace(/\s+on\w+='[^']*'/gi, '');
    if (content !== original) {
      zip.file(relativePath, content);
      logRemoval('script in SVG', relativePath);
    }
  }
}

async function checkHTMLImports(zip) {
  const xmlFiles = [];
  zip.forEach((relativePath, file) => {
    if (relativePath.toLowerCase().endsWith('.xml') ||
        relativePath.toLowerCase().endsWith('.rels')) {
      xmlFiles.push({ relativePath, file });
    }
  });

  for (const { relativePath, file } of xmlFiles) {
    const content = await file.async('string');
    if (content.includes('text/html') || content.includes('import ')) {
      logRemoval('HTML import', relativePath);
      throw Object.assign(
        new Error(`Rejected: HTML import detected in ${relativePath}`),
        { code: 'HTML_IMPORT_DETECTED', statusCode: 400 }
      );
    }
  }
}

/**
 * Sanitize an already-opened PPTX JSZip instance before parsing, in place.
 * Removes VBA macros, strips scripts from SVGs, and rejects HTML imports.
 * Every removal is logged for audit purposes (NFR-08).
 *
 * Mutates `zip` directly rather than returning a re-serialized buffer — the
 * caller reuses the same instance for parsing, so there's no need to pay for
 * a full zip.generateAsync() round-trip just to hand the bytes back.
 *
 * @param {JSZip} zip - an opened JSZip instance (mutated in place)
 * @returns {Promise<void>}
 */
async function sanitize(zip) {
  removeVBA(zip);
  await stripScriptsFromSVGs(zip);
  await checkHTMLImports(zip);
}

module.exports = { sanitize };