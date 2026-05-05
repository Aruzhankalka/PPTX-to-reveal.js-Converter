/**
 * IR validator — wraps Ajv around the Sprint 1 schema.
 *
 * Usage:
 *   const { validate } = require('./validator');
 *   const result = validate(irDocument);
 *   if (!result.valid) {
 *     console.error(result.errors);
 *   }
 */

const Ajv = require('ajv');
const schema = require('./schema');

const ajv = new Ajv({ allErrors: true, strict: false });
const validateFn = ajv.compile(schema);

/**
 * Validate an IR document against the Sprint 1 schema.
 * @param {object} doc - the IR document to validate
 * @returns {{ valid: boolean, errors: Array|null }}
 */
function validate(doc) {
  const valid = validateFn(doc);
  return {
    valid,
    errors: valid ? null : validateFn.errors,
  };
}

module.exports = { validate, schema };
