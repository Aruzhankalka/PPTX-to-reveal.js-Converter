const { validate } = require('../../ir/validator');
const { renderDocument } = require('./html');

/**
 * Generate a reveal.js HTML document from a validated IR document.
 *
 * Per Specification §4.2 step 8-9: the IR must be validated against the
 * schema before generation. A validation failure is treated as an internal
 * error, since Sprint 1 only ever generates from IR produced by our own
 * parser (or hand-written test fixtures).
 *
 * @param {object} ir - the IR document
 * @returns {{ html: string, warnings: string[] }}
 * @throws {Error} - if the IR fails schema validation
 */
function generate(ir) {
  const result = validate(ir);
  if (!result.valid) {
    const messages = result.errors
      .map((e) => `${e.instancePath || '/'} ${e.message}`)
      .join('; ');
    const err = new Error(`IR validation failed: ${messages}`);
    err.code = 'IR_VALIDATION_FAILED';
    err.details = result.errors;
    throw err;
  }

  const warnings = [];
  const html = renderDocument(ir);

  return { html, warnings };
}

module.exports = { generate };