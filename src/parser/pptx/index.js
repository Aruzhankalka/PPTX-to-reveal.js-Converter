const path = require('path');
const { openPptx, readBinary } = require('./zip');
const { listSlides, getSlideDimensions, getDefaultTextStyle } = require('./slides');
const { parseSlide } = require('./slide');
const { parseMaster } = require('./master');
const { parseFonts } = require('./fonts');
const { parseCoreProps } = require('./core');
const { validate } = require('../../ir/validator');

// Spec's aspect-ratio field is a closed enum (16:9|4:3) — PPTX allows
// arbitrary custom slide sizes, so clamp to whichever of the two is
// numerically closest rather than emitting an out-of-contract ratio
// like "8:5" or "1:1".
function nearestAspectRatio(w, h) {
  if (!w || !h) return '16:9';
  const ratio = w / h;
  const dist169 = Math.abs(ratio - 16 / 9);
  const dist43  = Math.abs(ratio - 4 / 3);
  return dist169 <= dist43 ? '16:9' : '4:3';
}

/**
 * Parse a .pptx buffer (or an already-opened JSZip instance) into a validated
 * IR document plus the list of media files that need to be bundled with the
 * generated reveal.js output.
 *
 * Accepting either input lets the upload route open the zip once — reusing it
 * across structural validation, sanitization, and parsing — while every
 * existing caller that still hands over a raw Buffer keeps working unchanged.
 *
 * Per Specification §4.2 (forward conversion pipeline):
 *  - Step 4: extract text and structure
 *  - Step 5: extract media
 *  - Step 8: validate IR against schema
 *
 * Step 6 (master/layout/theme) — FR-11, FR-12, implemented in Sprint 2.
 * Step 7 (animations) is deferred to Sprint 3.
 *
 * @param {Buffer|JSZip} bufferOrZip - the uploaded .pptx file bytes, or an
 *   already-opened JSZip instance for that same file
 * @param {object} [options]
 * @param {string} [options.filename] - original filename, included in IR metadata
 * @returns {Promise<{
 *   ir: object,
 *   media: Array<{ bundlePath: string, bytes: Buffer }>,
 *   warnings: string[]
 * }>}
 */
async function parsePptx(bufferOrZip, options = {}) {
  const warnings = [];
  const zip = Buffer.isBuffer(bufferOrZip) ? await openPptx(bufferOrZip) : bufferOrZip;

  // 1. Discover slides in document order
  const slideList = await listSlides(zip);

  // 2. Parse master first so txStyles and theme fmtScheme are available for slides.
  //    Slide dimensions and fonts are independent, so they run in parallel with
  //    the (sequential) slide loop via Promise.all.
  const masterResult = await parseMaster(zip);
  const txStyles = (masterResult && masterResult.txStyles) || null;
  const fmtScheme = (masterResult && masterResult.theme && masterResult.theme.fmtScheme) || null;
  // Attach theme colors so deepSubstitutePhClr can compute tint/shade/lumMod values
  if (fmtScheme && masterResult.theme.colors) fmtScheme.colors = masterResult.theme.colors;

  // 3. Parse slides (sequential — each slide may reference the previous),
  //    slide dimensions, and fonts in parallel.
  const slides = [];
  const mediaMap = new Map();

  const [, { slideWidth, slideHeight }, { fonts, fontBytes }, coreProps, defaultTextStyle] = await Promise.all([
    (async () => {
      for (const { path: slidePath } of slideList) {
        const { ir: slideIr, mediaRefs, layoutId, warnings: slideWarnings } = await parseSlide(zip, slidePath, txStyles, fmtScheme);
        if (slideWarnings && slideWarnings.length > 0) warnings.push(...slideWarnings);
        if (layoutId) slideIr['layout-id'] = layoutId;
        slides.push(slideIr);

        for (const ref of mediaRefs) {
          if (mediaMap.has(ref.bundlePath)) continue;
          const bytes = await readBinary(zip, ref.zipPath);
          if (bytes) {
            mediaMap.set(ref.bundlePath, bytes);
          } else {
            warnings.push(`Referenced media not found in archive: ${ref.zipPath}`);
          }
        }
      }
    })(),
    getSlideDimensions(zip),
    parseFonts(zip),
    parseCoreProps(zip),
    getDefaultTextStyle(zip),
  ]);

  // 4. Build the slideset
  // OOXML colour-scheme slot → professor-format CSS variable name.
  // Matches the mapping already used by renderThemeVariables in html.js.
  const SLOT_TO_CSS_VAR = {
    accent1: 'accent1', accent2: 'accent2', accent3: 'accent3',
    accent4: 'accent4', accent5: 'accent5', accent6: 'accent6',
    dk1: 'text-dark',  lt1: 'text-light',
    dk2: 'bg-dark',    lt2: 'bg-light',
    hlink: 'link',     folHlink: 'link-visited',
  };

  const master = {};
  if (masterResult) {
    if (masterResult.theme) {
      master.theme = masterResult.theme;
      const colors = masterResult.theme.colors || {};
      const colorTheme = Object.entries(SLOT_TO_CSS_VAR)
        .filter(([slot]) => colors[slot] != null)
        .map(([slot, cssVar]) => ({ 'css-variable-name': cssVar, color: colors[slot] }));
      if (colorTheme.length > 0) master['color-theme'] = colorTheme;
    }
    if (masterResult.masterName) master.name = masterResult.masterName;
  }
  // FR-11/12 (extended): presentation-wide default text style ("Global preset!")
  if (defaultTextStyle) master.formatting = defaultTextStyle;
  if (slideWidth != null || slideHeight != null) {
    const w = slideWidth  ?? 960;
    const h = slideHeight ?? 540;
    master['slide-dimensions'] = { width: w, height: h };
    master['dimension-units']  = 'px';
    master['aspect-ratio'] = nearestAspectRatio(w, h);
  }
  const layouts = (masterResult && masterResult.layouts) || [];

  const ir = {
    slideset: {
      filename: options.filename || 'unknown.pptx',
      ...(coreProps.title        && { title:           coreProps.title }),
      ...(coreProps.author       && { author:          coreProps.author }),
      ...(coreProps.creationDate && { 'creation-date': coreProps.creationDate }),
      slides,
      ...(fonts.length > 0 && { fonts }),
      master,
      layouts,
    },
  };

  // 5. Validate against the schema
  const v = validate(ir);
  if (!v.valid) {
    const messages = v.errors.map((e) => `${e.instancePath || '/'} ${e.message}`).join('; ');
    const err = new Error(`Parser produced invalid IR: ${messages}`);
    err.code = 'IR_VALIDATION_FAILED';
    err.details = v.errors;
    throw err;
  }

  // 6. Convert media map to array and append extracted font bytes
  const media = [
    ...Array.from(mediaMap.entries()).map(([bundlePath, bytes]) => ({ bundlePath, bytes })),
    ...fontBytes,
  ];

  return { ir, media, warnings };
}

module.exports = { parsePptx };