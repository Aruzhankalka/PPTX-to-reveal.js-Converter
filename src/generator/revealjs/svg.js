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

function colorValue(color) {
  if (!color) return 'none';

  if (typeof color === 'object') {
    if (color.space === 'srgb' && color.hex) return `#${color.hex}`;
    if (color.space === 'theme' && color.ref) return `var(--theme-${color.ref})`;
    return 'none';
  }

  return color;
}

function fillValue(fill) {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'solid') return colorValue(fill.color);
  return 'none';
}

function strokeValue(stroke) {
  if (!stroke || stroke.type === 'none' || stroke.style === 'none') {
    return { color: 'none', widthPx: 0 };
  }

  const color = colorValue(stroke.color);
  if (!color || color === 'none') {
    return { color: 'none', widthPx: 0 };
  }

  let widthPx = 0;
  if (stroke.widthEmu != null) {
    widthPx = emuToPx(stroke.widthEmu) ?? 0;
  } else if (typeof stroke.width === 'number') {
    widthPx = Math.round(stroke.width * PT_TO_PX);
  }

  return { color, widthPx };
}

function svgPaintAttrs(shape) {
  let fill = fillValue(shape.fill);
  let { color, widthPx } = strokeValue(shape.stroke);

  return `fill="${escapeHtml(fill)}" stroke="${escapeHtml(color)}" stroke-width="${widthPx}"`;
}
/**
 * Emit a regular N-sided polygon inscribed in the bounding box.
 * The first vertex is at the top center (startAngle = -π/2).
 */
function emitRegularPolygon(wPx, hPx, sides) {
  const cx = wPx / 2;
  const cy = hPx / 2;
  const pts = [];
  for (let i = 0; i < sides; i++) {
    const a = (2 * Math.PI * i / sides) - Math.PI / 2;
    pts.push(`${(cx + cx * Math.cos(a)).toFixed(2)},${(cy + cy * Math.sin(a)).toFixed(2)}`);
  }
  return `<polygon points="${pts.join(' ')}"/>`;
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

  const fill = fillValue(shape.fill);
  const { color: sc, widthPx: sw } = strokeValue(shape.stroke);
  const paintAttrs = svgPaintAttrs(shape);

  const hasVisiblePaint = fill !== 'none' || (sc !== 'none' && sw > 0);
  const fallbackPaintAttrs = hasVisiblePaint
  ? paintAttrs
  : 'fill="rgba(255, 165, 0, 0.35)" stroke="#ff6600" stroke-width="1"';

  const strokeColorForArrow = colorValue(shape.stroke?.color);
  const arrowFill = strokeColorForArrow !== 'none' ? strokeColorForArrow : fillValue(shape.fill);
  const transform = buildTransform(xPx, yPx, rotation, wPx, hPx);

  let primitive;

  switch (type) {
    case 'rect':
      primitive = emitRect(wPx, hPx, shape.geometry);
      break;
  
    case 'roundRect': {
      const adjVal = shape.adjustments && shape.adjustments.adj;
      const rxRaw = adjVal != null
        ? Math.round((adjVal / 100000) * Math.min(wPx, hPx) / 2)
        : ((shape.geometry && shape.geometry.rx) ?? 8);
  
      primitive = emitRect(wPx, hPx, { rx: rxRaw, ry: rxRaw });
      break;
    }


    // case 'triangle':
    //   primitive = emitRegularPolygon(wPx, hPx, 3);
    //   break;

    // case 'hexagon':
    //   primitive = emitRegularPolygon(wPx, hPx, 6);
    //   break;

    // case 'octagon':
    //   primitive = emitRegularPolygon(wPx, hPx, 8);
    //   break;

    // Stubs — each type is a separate branch so adding one later is a single
    // new case block. All unsupported types warn and return '' without throwing.

    case 'ellipse':
      primitive = `<ellipse cx="${wPx / 2}" cy="${hPx / 2}" rx="${wPx / 2}" ry="${hPx / 2}"  ${paintAttrs} />`;
      break;
  
    case 'triangle':
      primitive = `<polygon points="${wPx / 2},0 ${wPx},${hPx} 0,${hPx}"  ${paintAttrs} />`;
      break;
  
    case 'hexagon':
      primitive = `<polygon points="${wPx * 0.25},0 ${wPx * 0.75},0 ${wPx},${hPx / 2} ${wPx * 0.75},${hPx} ${wPx * 0.25},${hPx} 0,${hPx / 2}"  ${paintAttrs}/>`;
      break;
  
    case 'octagon':
      primitive = `<polygon points="${wPx * 0.3},0 ${wPx * 0.7},0 ${wPx},${hPx * 0.3} ${wPx},${hPx * 0.7} ${wPx * 0.7},${hPx} ${wPx * 0.3},${hPx} 0,${hPx * 0.7} 0,${hPx * 0.3}"  ${paintAttrs}/>`;
      break;
  
    case 'line':
      primitive = `<line x1="0" y1="${hPx / 2}" x2="${wPx}" y2="${hPx / 2}" stroke="${escapeHtml(sc)}" stroke-width="${sw}" fill="none" />`;
      break;
  
    case 'connector':
      primitive = `<line x1="0" y1="${hPx / 2}" x2="${wPx}" y2="${hPx / 2}" stroke="${escapeHtml(sc)}" stroke-width="${sw}" fill="none" />`;
      break;
  
    case 'arrow':
      primitive = `
        <line x1="0" y1="${hPx / 2}" x2="${wPx * 0.8}" y2="${hPx / 2}" stroke="${escapeHtml(sc)}" stroke-width="${sw}" fill="none" />
        <polygon points="${wPx * 0.8},${hPx * 0.25} ${wPx},${hPx / 2} ${wPx * 0.8},${hPx * 0.75}" fill="${escapeHtml(arrowFill)}" stroke="${escapeHtml(sc)}" stroke-width="${sw}" />
      `;
      break;
  
    case 'cloud':
      primitive = `
        <rect x="0" y="0" width="${wPx}" height="${hPx}" rx="${Math.min(wPx, hPx) * 0.25}" ry="${Math.min(wPx, hPx) * 0.25}"  ${paintAttrs} />
      `;
      break;
  
    case 'star7':
      primitive = `<polygon points="${wPx * 0.5},0 ${wPx * 0.6},${hPx * 0.35} ${wPx},${hPx * 0.25} ${wPx * 0.68},${hPx * 0.5} ${wPx * 0.85},${hPx} ${wPx * 0.5},${hPx * 0.7} ${wPx * 0.15},${hPx} ${wPx * 0.32},${hPx * 0.5} 0,${hPx * 0.25} ${wPx * 0.4},${hPx * 0.35}"  ${paintAttrs}/>`;
      break;
  
    case 'flowChartMagneticDisk':
      primitive = `
        <ellipse cx="${wPx / 2}" cy="${hPx * 0.18}" rx="${wPx * 0.45}" ry="${hPx * 0.14}"  ${paintAttrs} />
        <path d="M${wPx * 0.05} ${hPx * 0.18} V${hPx * 0.82} C${wPx * 0.05} ${hPx * 0.95}, ${wPx * 0.95} ${hPx * 0.95}, ${wPx * 0.95} ${hPx * 0.82} V${hPx * 0.18}" ${paintAttrs} />
        <ellipse cx="${wPx / 2}" cy="${hPx * 0.82}" rx="${wPx * 0.45}" ry="${hPx * 0.14}" fill="none" stroke="${escapeHtml(sc)}" stroke-width="${sw}" />
      `;
      break;
  
    case 'arc':
      primitive = `<path d="M${wPx * 0.15} ${hPx * 0.85} C${wPx * 0.15} ${hPx * 0.25}, ${wPx * 0.85} ${hPx * 0.25}, ${wPx * 0.85} ${hPx * 0.85}" fill="none" stroke="${escapeHtml(sc)}" stroke-width="${Math.max(sw, 2)}" />`;
      break;
  
    case 'polyline':
    case 'polygon':
    case 'callout':
    case 'unknown':
    default:
      warnings.push(`shape type ${type} not yet supported`);
      return '';
  }

  // New IR: shape.text is a TextBlock object { id, paragraphs }
  // Old IR: shape.text is a plain paragraphs array
  const paragraphs = Array.isArray(shape.text)
    ? shape.text
    : (shape.text && shape.text.paragraphs) || [];
  // const fo = emitForeignObject(paragraphs, wPx, hPx);
  const shouldRenderText = ctx.renderText !== false;
  const fo = shouldRenderText ? emitForeignObject(paragraphs, wPx, hPx) : '';
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
function renderShape(shape, options = {}) {
  const ctx = { warnings: [], renderText: options.renderText !== false };
  const g = emitShape(shape, ctx);
  if (!g) return '';

  // New IR uses shape.z; old IR uses shape['z-index']. Accept either.
  let zIndex = typeof shape['z-index'] === 'number' ? shape['z-index']
  : typeof shape.z === 'number' ? shape.z
  : 0;

const fill =
  shape.fill?.type === 'solid' && shape.fill?.color?.space === 'srgb'
    ? `#${shape.fill.color.hex}`.toUpperCase()
    : shape.fill?.type === 'solid' && shape.fill?.color?.space === 'theme'
      ? shape.fill.color.ref
      : '';

const isLayoutShape = shape.id?.startsWith('layout-shp');

if (isLayoutShape && (fill === 'bg1' || fill === '#FFFFFF')) {
  zIndex = 1;
} else if (isLayoutShape) {
  zIndex = 2;
}
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