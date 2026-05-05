/**
 * Escape a string for safe insertion into HTML text content.
 * Prevents XSS by neutralizing HTML metacharacters in user-supplied text.
 *
 * Per Specification NFR-08: no execution of code embedded in input.
 * Even though the sanitizer (Sprint 1, step 4) strips macros and scripts
 * from the source PPTX, defense in depth requires us to also escape any
 * text that ends up in generated HTML.
 *
 * @param {string} str
 * @returns {string}
 */
function escapeHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/**
 * Escape a string for safe use as a CSS attribute value (e.g. inside style="").
 * Strips characters that would let user input break out of the attribute.
 *
 * @param {string} str
 * @returns {string}
 */
function escapeCss(str) {
  if (str == null) return '';
  // CSS in HTML attributes is delimited by quotes; the dangerous characters
  // are quote, semicolon (statement break), and angle brackets.
  return String(str).replace(/["'<>;]/g, '');
}

module.exports = { escapeHtml, escapeCss }; 