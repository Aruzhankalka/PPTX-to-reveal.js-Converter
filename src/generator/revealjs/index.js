/**
 * Public entry point of the reveal.js generator (parser/pptx is the
 * mirror-image entry point on the other side of the pipeline). Thin wrapper
 * around html.js's renderDocument that gives callers a stable
 * {html, warnings} return shape independent of how the document gets built.
 */

const { renderDocument } = require('./html');

/**
 * Generate a reveal.js HTML document from an IR document.
 *
 * Per Specification §4.2 step 8-9, the IR must be validated against the
 * schema before generation — but that check happens exactly once, at the end
 * of parsePptx (src/parser/pptx/index.js), which is the sole production path
 * into this function (via src/api/upload.js). generate() does not re-validate;
 * callers that hand it IR the parser never touched (e.g. hand-written test
 * fixtures exercising invalid documents) must call validate() themselves.
 *
 * @param {object} ir - the IR document (assumed already schema-valid)
 * @returns {{ html: string, warnings: string[] }}
 */
function generate(ir) {
  const warnings = [];
  const html = renderDocument(ir);

  return { html, warnings };
}

module.exports = { generate };