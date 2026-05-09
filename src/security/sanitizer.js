'use strict';

const JSZip = require('jszip');

/**
 * Sanitizer — NFR-08
 * Strips dangerous content from a PPTX buffer before parsing.
 * Logs every removal for audit purposes.
 */

// ── Helper: logger ────────────────────────────────────────────────────────────
function logRemoval(type, detail) {
  console.warn(`[SANITIZER] Removed ${type}: ${detail}`);
}

// ── 1. Remove vbaProject.bin ──────────────────────────────────────────────────
function removeVBA(zip) {
  let removed = false;
  zip.forEach((relativePath) => {
    if (relativePath.toLowerCase().includes('vbaproject.bin')) {
      zip.remove(relativePath);
      logRemoval('VBA macro', relativePath);
      removed = true;
    }
  });
  return removed;
}

// ── 2. Strip <script> tags from SVG files ─────────────────────────────────────
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
    // Remove <script>...</script> blocks
    content = content.replace(/<script[\s\S]*?<\/script>/gi, '');
    // Remove inline event handlers like onload, onclick etc.
    content = content.replace(/\s+on\w+="[^"]*"/gi, '');
    content = content.replace(/\s+on\w+='[^']*'/gi, '');
    if (content !== original) {
      zip.file(relativePath, content);
      logRemoval('script in SVG', relativePath);
    }
  }
}

// ── 3. Refuse HTML imports ────────────────────────────────────────────────────
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

// ── Main export ───────────────────────────────────────────────────────────────
/**
 * Sanitize a PPTX buffer.
 * @param {Buffer} buffer - Raw PPTX file buffer
 * @returns {Promise<Buffer>} - Cleaned PPTX buffer
 */
async function sanitize(buffer) {
  // Load the ZIP
  const zip = await JSZip.loadAsync(buffer);

  // 1. Remove VBA macros
  removeVBA(zip);

  // 2. Strip scripts from SVGs
  await stripScriptsFromSVGs(zip);

  // 3. Check for HTML imports
  await checkHTMLImports(zip);

  // Return cleaned buffer
  const cleanedBuffer = await zip.generateAsync({ type: 'nodebuffer' });
  return cleanedBuffer;
}

module.exports = { sanitize };