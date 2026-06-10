/**
 * IR validator — wraps Ajv around the Sprint 2/3 schema and adds a
 * cross-document targetId referential-integrity check for FR-14 animations.
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

// ---------------------------------------------------------------------------
// Cross-document targetId validation (FR-14)
//
// JSON Schema cannot express "targetId must reference an id that exists
// somewhere else in the same slide object" without custom keywords.
// We implement it as a post-schema runtime check and merge the errors into
// the same { valid, errors } envelope so callers see a unified result.
// ---------------------------------------------------------------------------

/**
 * Collect every element id defined on a slide (textBlock, shape, mediaItem).
 *
 * @param {object} slideContents - slide.contents object
 * @returns {Set<string>}
 */
function collectSlideIds(slideContents) {
  const ids = new Set();
  if (!slideContents) return ids;

  for (const block of (slideContents.text || [])) {
    if (block && block.id) ids.add(block.id);
  }
  for (const shape of (slideContents.shapes || [])) {
    if (shape && shape.id) ids.add(shape.id);
  }
  for (const media of (slideContents.media || [])) {
    if (media && media.id) ids.add(media.id);
  }

  return ids;
}

/**
 * Validate that every animation.targetId on every slide references a defined
 * element id on that same slide.
 *
 * @param {object} doc - IR document
 * @returns {Array<{message: string, instancePath: string}>} error objects
 *   (empty array when all references are valid)
 */
function validateTargetIds(doc) {
  const errors = [];
  const slides = (doc && doc.slideset && doc.slideset.slides) || [];

  for (let si = 0; si < slides.length; si++) {
    const slide = slides[si];
    const contents = slide && slide.contents;
    const animations = (contents && contents.animations) || [];
    if (animations.length === 0) continue;

    const ids = collectSlideIds(contents);

    for (let ai = 0; ai < animations.length; ai++) {
      const anim = animations[ai];
      if (!anim || !anim.targetId) continue;
      // spid-N targetIds are raw PPTX shape IDs not yet resolved to IR element
      // ids — skip them to avoid false positives until the mapping is implemented.
      if (/^spid-\d+$/.test(anim.targetId)) continue;
      if (!ids.has(anim.targetId)) {
        errors.push({
          message: `animation "${anim.id || ai}" has targetId "${anim.targetId}" which does not match any element id on slide ${si}`,
          instancePath: `/slideset/slides/${si}/contents/animations/${ai}/targetId`,
        });
      }
    }
  }

  return errors;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Validate an IR document against the Sprint 2/3 schema plus targetId refs.
 *
 * @param {object} doc - the IR document to validate
 * @returns {{ valid: boolean, errors: Array|null }}
 */
function validate(doc) {
  const schemaValid = validateFn(doc);
  if (!schemaValid) {
    return { valid: false, errors: validateFn.errors };
  }

  const refErrors = validateTargetIds(doc);
  if (refErrors.length > 0) {
    return { valid: false, errors: refErrors };
  }

  return { valid: true, errors: null };
}

module.exports = { validate, schema, validateTargetIds };
