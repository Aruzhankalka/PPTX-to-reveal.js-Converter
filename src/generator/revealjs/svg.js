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
    const c = fill.color;
    if (!c) return 'none';
    // New IR format: structured Color { space, hex } or { space, ref }
    if (typeof c === 'object') {
      if (c.space === 'srgb'  && c.hex) return `#${c.hex}`;
      if (c.space === 'theme' && c.ref) return `var(--theme-${c.ref})`;
      return 'none';
    }
    // Old IR format: CSS color string (e.g. '#ff0000', 'var(--theme-accent1)')
    return c;
  }
  return 'none'; // gradient/pattern: not yet implemented
}

function strokeAttrs(stroke) {
  // New IR format uses stroke.type; old format uses stroke.style.
  if (!stroke || stroke.type === 'none' || stroke.style === 'none') {
    return { color: 'none', widthPx: 0 };
  }

  // Color: new format is a structured Color object; old format is a CSS string.
  let color = 'none';
  const c = stroke.color;
  if (typeof c === 'object' && c !== null) {
    if (c.space === 'srgb'  && c.hex) color = `#${c.hex}`;
    else if (c.space === 'theme' && c.ref) color = `var(--theme-${c.ref})`;
  } else if (typeof c === 'string' && c) {
    color = c;
  }
  if (!color || color === 'none') return { color: 'none', widthPx: 0 };

  // Width: new format is widthEmu (integer EMU); old format is width in points.
  let widthPx = 0;
  if (stroke.widthEmu != null) {
    widthPx = emuToPx(stroke.widthEmu) ?? 0;
  } else if (typeof stroke.width === 'number') {
    widthPx = Math.round(stroke.width * PT_TO_PX);
  }

  return { color, widthPx };
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
  // New IR: w/h live inside position; old IR: separate shape.width / shape.height fields.
  const wPx = emuToPx(shape.width  ?? (shape.position && shape.position.w)) ?? 0;
  const hPx = emuToPx(shape.height ?? (shape.position && shape.position.h)) ?? 0;
  // New IR: rotation in native PPTX rot units (1/60000 of a degree); old IR: degrees.
  // Values > 360 can only be PPTX rot units — convert them to degrees for SVG.
  const rawRot = typeof shape.rotation === 'number' ? shape.rotation : 0;
  const rotation = rawRot > 360 ? rawRot / 60000 : rawRot;
  const opacity   = typeof shape.opacity  === 'number' ? shape.opacity  : 1;

  const fill = fillAttr(shape.fill);
  const { color: sc, widthPx: sw } = strokeAttrs(shape.stroke);
  const transform = buildTransform(xPx, yPx, rotation, wPx, hPx);

  let primitive;
  switch (type) {
    case 'rect':
      primitive = emitRect(wPx, hPx, shape.geometry);
      break;

    // roundRect: use adjustments.adj (new IR) or geometry.rx (old IR) for corner radius.
    case 'roundRect': {
      const adjVal = shape.adjustments && shape.adjustments.adj;
      // adj is in 1/100000 units; convert to a pixel radius capped at half the shorter side.
      const rxRaw = adjVal != null
        ? Math.round((adjVal / 100000) * Math.min(wPx, hPx) / 2)
        : ((shape.geometry && shape.geometry.rx) ?? 8);
      primitive = emitRect(wPx, hPx, { rx: rxRaw, ry: rxRaw });
      break;
    }

    // Stubs — each type is a separate branch so adding one later is a single
    // new case block. All unsupported types warn and return '' without throwing.
    case 'ellipse':
    case 'line':
    case 'arrow':
    case 'polyline':
    case 'polygon':
    case 'callout':
    case 'connector':
    case 'unknown':
      warnings.push(`shape type ${type} not yet supported`);
      return '';

    default:
      warnings.push(`shape type ${type} not yet supported`);
      return '';
  }

  // New IR: shape.text is a TextBlock object { id, paragraphs }
  // Old IR: shape.text is a plain paragraphs array
  const paragraphs = Array.isArray(shape.text)
    ? shape.text
    : (shape.text && shape.text.paragraphs) || [];
  const fo = emitForeignObject(paragraphs, wPx, hPx);
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

  // New IR uses shape.z; old IR uses shape['z-index']. Accept either.
  const zIndex = typeof shape['z-index'] === 'number' ? shape['z-index']
    : typeof shape.z === 'number' ? shape.z
    : 0;
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
