'use strict';

const { emuToPx } = require('../../parser/pptx/units');
const { renderParagraph } = require('./text');
const { escapeHtml } = require('./escape');

// Stroke widths in the IR are stored in points. Converting to SVG px uses the
// same EMU-based ratio as font sizes: pt × 12700 (EMU/pt) / 9525 (EMU/px).
// This is identical to the standard CSS 1 pt = 4/3 px rule at 96 DPI.
const PT_TO_PX = 12700 / 9525;

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

function fillAttr(fill) {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'solid') {
    // OPEN QUESTION (FR-12): the parser may resolve theme slot references
    // (e.g. <a:schemeClr val="accent1"/>) to a baked hex color, losing the
    // original slot name. If fill.color is always a resolved hex, the
    // requirement to emit var(--theme-accent1) cannot be satisfied without
    // adding a separate 'themeVar' field to the IR shapeFill schema.
    // Current behavior: pass fill.color verbatim — if the parser preserves
    // theme vars (e.g. 'var(--theme-accent1)') they appear unchanged in SVG.
    return fill.color || 'none';
  }
  return 'none'; // gradient/pattern: not yet implemented
}

function strokeAttrs(stroke) {
  if (!stroke || stroke.style === 'none' || !stroke.color) {
    return { color: 'none', widthPx: 0 };
  }
  const widthPx = typeof stroke.width === 'number'
    ? Math.round(stroke.width * PT_TO_PX)
    : 0;
  return { color: stroke.color, widthPx };
}

function emitRect(wPx, hPx, geometry) {
  const rx = (geometry != null && geometry.rx != null) ? Number(geometry.rx) : 0;
  const ry = (geometry != null && geometry.ry != null) ? Number(geometry.ry) : rx;
  return `<rect x="0" y="0" width="${wPx}" height="${hPx}" rx="${rx}" ry="${ry}"/>`;
}

function emitForeignObject(paragraphs, wPx, hPx) {
  if (!paragraphs || paragraphs.length === 0) return '';
  const body = paragraphs.map(renderParagraph).join('');
  // foreignObject requires an explicit XHTML namespace on its HTML root so
  // browsers parse the HTML content correctly inside the SVG document.
  return (
    `<foreignObject x="0" y="0" width="${wPx}" height="${hPx}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml"` +
    ` style="width:100%;height:100%;overflow:hidden;box-sizing:border-box;">` +
    body +
    `</div></foreignObject>`
  );
}

function buildTransform(xPx, yPx, rotation, wPx, hPx) {
  const t = `translate(${xPx},${yPx})`;
  if (!rotation || rotation === 0) return t;
  // rotate() center is in the post-translate local coordinate system,
  // so (wPx/2, hPx/2) is the shape's own center regardless of position.
  return `${t} rotate(${rotation} ${wPx / 2} ${hPx / 2})`;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Emit one IR shape as an SVG <g> fragment.
 *
 * Position and dimensions arrive in EMU; they are converted to CSS pixels
 * with the same 1 px = 9525 EMU factor used by the rest of the geometry
 * pipeline (src/parser/pptx/units.js). The "rect" type is fully implemented;
 * all other types push a warning and return '' so the caller continues safely.
 *
 * @param {object} shape  IR Shape (position.x/y, width, height in EMU; rotation in degrees)
 * @param {{ warnings: string[] }} ctx  shared context — warnings are appended in-place
 * @returns {string}  '<g ...>...</g>' SVG fragment, or '' for unsupported types
 */
function emitShape(shape, ctx) {
  const warnings = ctx && Array.isArray(ctx.warnings) ? ctx.warnings : [];
  const type = shape.type;

  const xPx = emuToPx(shape.position && shape.position.x) ?? 0;
  const yPx = emuToPx(shape.position && shape.position.y) ?? 0;
  const wPx = emuToPx(shape.width)  ?? 0;
  const hPx = emuToPx(shape.height) ?? 0;
  const rotation  = typeof shape.rotation === 'number' ? shape.rotation : 0;
  const opacity   = typeof shape.opacity  === 'number' ? shape.opacity  : 1;

  const fill = fillAttr(shape.fill);
  const { color: sc, widthPx: sw } = strokeAttrs(shape.stroke);
  const transform = buildTransform(xPx, yPx, rotation, wPx, hPx);

  let primitive;
  switch (type) {
    case 'rect':
      primitive = emitRect(wPx, hPx, shape.geometry);
      break;

    // Stubs — each type is a separate branch so adding one later is a single
    // new case block. All unsupported types warn and return '' without throwing.
    case 'ellipse':
    case 'line':
    case 'arrow':
    case 'polyline':
    case 'polygon':
    case 'callout':
    case 'connector':
      warnings.push(`shape type ${type} not yet supported`);
      return '';

    default:
      warnings.push(`shape type ${type} not yet supported`);
      return '';
  }

  const fo = emitForeignObject(shape.text, wPx, hPx);
  const inner = fo
    ? `\n  ${primitive}\n  ${fo}\n`
    : `\n  ${primitive}\n`;

  return (
    `<g transform="${transform}"` +
    ` fill="${escapeHtml(fill)}"` +
    ` stroke="${escapeHtml(sc)}"` +
    ` stroke-width="${sw}"` +
    ` opacity="${opacity}">` +
    inner +
    `</g>`
  );
}

/**
 * Wrap an emitShape result in an absolutely-positioned full-canvas SVG element
 * for use inside a .slide-canvas div.
 *
 * The SVG covers the entire canvas so the <g>'s translate() positions the
 * shape in slide pixel space without double-counting. z-index controls
 * stacking relative to text blocks and images.
 *
 * @param {object} shape  IR Shape
 * @returns {string}      complete <svg> element, or '' for unsupported types
 */
function renderShape(shape) {
  const ctx = { warnings: [] };
  const g = emitShape(shape, ctx);
  if (!g) return '';

  const zIndex = typeof shape['z-index'] === 'number' ? shape['z-index'] : 0;
  const style = [
    'position:absolute',
    'left:0',
    'top:0',
    'width:100%',
    'height:100%',
    'overflow:visible',
    'pointer-events:none',
    `z-index:${zIndex}`,
  ].join(';');

  return (
    `<svg style="${style}" xmlns="http://www.w3.org/2000/svg">` +
    g +
    `</svg>`
  );
}

module.exports = { emitShape, renderShape };
