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

function colorToCss(c) {
  if (!c) return null;
  if (typeof c === 'object') {
    if (c.space === 'srgb'  && c.hex) return `#${c.hex}`;
    if (c.space === 'theme' && c.ref) return `var(--theme-${c.ref})`;
    return null;
  }
  return c || null; // old IR: plain CSS string
}

/** Return fill-opacity fraction (0–1) from a shapeColor's alpha field. */
function colorOpacity(c) {
  if (!c || typeof c !== 'object' || c.alpha == null) return 1;
  return c.alpha / 100;
}

function fillAttr(fill) {
  if (!fill || fill.type === 'none') return 'none';
  if (fill.type === 'solid') return colorToCss(fill.color) || 'none';
  if (fill.type === 'gradient') return null; // caller handles via emitGradientDefs
  return 'none';
}

/**
 * Convert a PPTX gradient linear angle (1/60000 degrees, CW from east) to
 * SVG linearGradient x1/y1/x2/y2 in objectBoundingBox % coordinates.
 *
 * ang=0 → left-to-right; ang=5400000 (90°) → top-to-bottom.
 */
function angleToGradientCoords(pptxAngle) {
  const deg = ((pptxAngle || 0) / 60000) % 360;
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return {
    x1: `${((0.5 - cos / 2) * 100).toFixed(2)}%`,
    y1: `${((0.5 - sin / 2) * 100).toFixed(2)}%`,
    x2: `${((0.5 + cos / 2) * 100).toFixed(2)}%`,
    y2: `${((0.5 + sin / 2) * 100).toFixed(2)}%`,
  };
}

/**
 * Build an SVG <defs> block for a gradient fill and return the fill reference.
 * Returns { defs:'', fillRef:'none' } for non-gradient fills.
 *
 * Theme-color stops use style="stop-color:var(...)" so CSS custom properties
 * cascade correctly (SVG attributes don't support var()).
 */
function emitGradientDefs(fill, shapeId) {
  if (!fill || fill.type !== 'gradient') return { defs: '', fillRef: 'none' };
  const gradId = `grad-${String(shapeId).replace(/[^a-zA-Z0-9_-]/g, '-')}`;

  const stopEls = (fill.stops || []).map((s) => {
    const pct = ((s.pos / 100000) * 100).toFixed(2);
    const css = colorToCss(s.color);
    const op  = colorOpacity(s.color);
    // Use style attribute so CSS var() resolves; bare attribute does not support var()
    const colorPart = css && css.startsWith('var(')
      ? `style="stop-color:${css}${op < 1 ? `;stop-opacity:${op}` : ''}"`
      : `stop-color="${css || '#000'}"${op < 1 ? ` stop-opacity="${op}"` : ''}`;
    return `<stop offset="${pct}%" ${colorPart}/>`;
  }).join('');

  let gradEl;
  if (fill.kind === 'radial') {
    gradEl = `<radialGradient id="${gradId}" cx="50%" cy="50%" r="50%" gradientUnits="objectBoundingBox">${stopEls}</radialGradient>`;
  } else {
    const { x1, y1, x2, y2 } = angleToGradientCoords(fill.angle);
    gradEl = `<linearGradient id="${gradId}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" gradientUnits="objectBoundingBox">${stopEls}</linearGradient>`;
  }

  return { defs: `<defs>${gradEl}</defs>`, fillRef: `url(#${gradId})` };
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

function buildTransform(xPx, yPx, rotation, wPx, hPx, flipH, flipV) {
  const parts = [`translate(${xPx},${yPx})`];
  // Flip around the shape centre before rotation so both transforms compose correctly.
  if (flipH || flipV) {
    const cx = wPx / 2, cy = hPx / 2;
    parts.push(`translate(${cx},${cy}) scale(${flipH ? -1 : 1},${flipV ? -1 : 1}) translate(${-cx},${-cy})`);
  }
  if (rotation) {
    parts.push(`rotate(${rotation} ${wPx / 2} ${hPx / 2})`);
  }
  return parts.join(' ');
}

// ---------------------------------------------------------------------------
// Arrow head markers
// ---------------------------------------------------------------------------

const ARROW_MARKER_PATH = {
  triangle: (c) => `<polygon points="0 0, 10 5, 0 10" fill="${c}"/>`,
  stealth:  (c) => `<polygon points="0 0, 10 5, 0 10, 5 5" fill="${c}"/>`,
  arrow:    (c) => `<polyline points="0 0, 10 5, 0 10" fill="none" stroke="${c}" stroke-width="1.5"/>`,
  diamond:  (c) => `<polygon points="5 0, 10 5, 5 10, 0 5" fill="${c}"/>`,
  oval:     (c) => `<ellipse cx="5" cy="5" rx="5" ry="5" fill="${c}"/>`,
};

/**
 * Build SVG <marker> defs and attribute references for stroke arrowheads.
 * Returns { markerDefs, headAttr, tailAttr } — attrs are empty strings when absent.
 */
function buildArrowMarkers(stroke, shapeId) {
  if (!stroke || stroke.type !== 'solid') return { markerDefs: '', headAttr: '', tailAttr: '' };
  const { headEnd, tailEnd } = stroke;
  if (!headEnd && !tailEnd) return { markerDefs: '', headAttr: '', tailAttr: '' };

  const color   = colorToCss(stroke.color) || '#000';
  const safeId  = String(shapeId).replace(/[^a-zA-Z0-9_-]/g, '-');
  const defs    = [];
  let headAttr  = '';
  let tailAttr  = '';

  function makeMarker(end, id, refX, orient) {
    const tpl = end && ARROW_MARKER_PATH[end.type];
    if (!tpl) return '';
    const marker = (
      `<marker id="${id}" markerWidth="10" markerHeight="10"` +
      ` refX="${refX}" refY="5" orient="${orient}" markerUnits="strokeWidth">` +
      tpl(color) +
      `</marker>`
    );
    defs.push(marker);
    return `url(#${id})`;
  }

  // headEnd = start of the line (refX=0, orient=auto-start-reverse so arrow faces inward)
  if (headEnd && headEnd.type && headEnd.type !== 'none') {
    headAttr = makeMarker(headEnd, `mh-${safeId}`, 0, 'auto-start-reverse')
      ? ` marker-start="url(#mh-${safeId})"` : '';
    // Re-check: makeMarker already pushes to defs; we just need the attr string
    if (defs.length) headAttr = ` marker-start="url(#mh-${safeId})"`;
  }
  const defsBeforeTail = defs.length;
  // tailEnd = end of the line (refX=10, orient=auto)
  if (tailEnd && tailEnd.type && tailEnd.type !== 'none') {
    makeMarker(tailEnd, `mt-${safeId}`, 10, 'auto');
    if (defs.length > defsBeforeTail) tailAttr = ` marker-end="url(#mt-${safeId})"`;
  }

  return {
    markerDefs: defs.length ? `<defs>${defs.join('')}</defs>` : '',
    headAttr,
    tailAttr,
  };
}

// ---------------------------------------------------------------------------
// Arrow polygon helpers
// ---------------------------------------------------------------------------

/**
 * Build a 7-point chevron-arrow polygon pointing in the given direction.
 * Points are in shape-local coordinates (0..w, 0..h).
 */
function arrowPolygon(wPx, hPx, direction) {
  const w = wPx, h = hPx;
  switch (direction) {
    case 'left':
      return `${w},${h*0.3} ${w*0.4},${h*0.3} ${w*0.4},0 0,${h*0.5} ${w*0.4},${h} ${w*0.4},${h*0.7} ${w},${h*0.7}`;
    case 'up':
      return `${w*0.3},${h} ${w*0.3},${h*0.4} 0,${h*0.4} ${w*0.5},0 ${w},${h*0.4} ${w*0.7},${h*0.4} ${w*0.7},${h}`;
    case 'down':
      return `${w*0.3},0 ${w*0.3},${h*0.6} 0,${h*0.6} ${w*0.5},${h} ${w},${h*0.6} ${w*0.7},${h*0.6} ${w*0.7},0`;
    default: // right
      return `0,${h*0.3} ${w*0.6},${h*0.3} ${w*0.6},0 ${w},${h*0.5} ${w*0.6},${h} ${w*0.6},${h*0.7} 0,${h*0.7}`;
  }
}

// Map PPTX arrow preset → direction string consumed by arrowPolygon.
const ARROW_DIRECTION = {
  leftArrow:      'left',  leftArrowCallout:  'left',
  upArrow:        'up',    upArrowCallout:    'up',
  downArrow:      'down',  downArrowCallout:  'down',
  rightArrow:     'right', rightArrowCallout: 'right',
  // Chevron / homePlate → approximated as right arrow
  chevron: 'right', homePlate: 'right',
  // Bent/curved arrows → approximated
  bentArrow: 'right', uturnArrow: 'right',
  curvedRightArrow: 'right', curvedLeftArrow:  'left',
  curvedUpArrow:    'up',   curvedDownArrow:  'down',
  stripedRightArrow: 'right', notchedRightArrow: 'right',
  leftRightArrow: 'right', upDownArrow: 'up',
};

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

  // Gradient fills need an SVG <defs> block; collect it alongside the fill ref.
  const rawFill = fillAttr(shape.fill);
  const { defs: gradDefs, fillRef: gradFillRef } = (rawFill === null)
    ? emitGradientDefs(shape.fill, shape.id || String(xPx + yPx))
    : { defs: '', fillRef: '' };
  const fill = rawFill !== null ? rawFill : gradFillRef;

  // Alpha transparency on solid fills → SVG fill-opacity (separate from general opacity).
  const fillOpacity = (shape.fill && shape.fill.type === 'solid')
    ? colorOpacity(shape.fill.color)
    : 1;

  const { color: sc, widthPx: sw } = strokeAttrs(shape.stroke);
  const flipH = !!shape.flipH;
  const flipV = !!shape.flipV;
  const transform = buildTransform(xPx, yPx, rotation, wPx, hPx, flipH, flipV);

  // Line / connector arrowhead markers (built before the switch so shapeId is in scope).
  const { markerDefs, headAttr, tailAttr } =
    (type === 'line' || type === 'connector')
      ? buildArrowMarkers(shape.stroke, shape.id || String(xPx + yPx))
      : { markerDefs: '', headAttr: '', tailAttr: '' };

  let primitive;
  switch (type) {
    case 'rect':
      primitive = emitRect(wPx, hPx, shape.geometry);
      break;

    // roundRect: adjustments is now [{name, value}] array (IR v2) or legacy {adj:N} object.
    case 'roundRect': {
      const adj = shape.adjustments;
      const adjVal = Array.isArray(adj)
        ? (adj.find((a) => a.name === 'adj') || {}).value
        : (adj && adj.adj);
      // adj value is in 1/100000 units; convert to pixel radius capped at half short side.
      const rxRaw = adjVal != null
        ? Math.round((adjVal / 100000) * Math.min(wPx, hPx) / 2)
        : ((shape.geometry && shape.geometry.rx) ?? 8);
      primitive = emitRect(wPx, hPx, { rx: rxRaw, ry: rxRaw });
      break;
    }

    case 'ellipse':
      primitive = `<ellipse cx="${wPx / 2}" cy="${hPx / 2}" rx="${wPx / 2}" ry="${hPx / 2}"/>`;
      break;

    case 'triangle':
      primitive = emitRegularPolygon(wPx, hPx, 3);
      break;

    case 'hexagon':
      primitive = emitRegularPolygon(wPx, hPx, 6);
      break;

    case 'octagon':
      primitive = emitRegularPolygon(wPx, hPx, 8);
      break;

    // Straight line — bounding box corner to corner.
    // flipH/V are handled in buildTransform; line always goes (0,0)→(w,h) in local coords.
    case 'line': {
      primitive = `<line x1="0" y1="0" x2="${wPx}" y2="${hPx}"${headAttr}${tailAttr}/>`;
      break;
    }

    // Connector — straight-line approximation; bent/elbow connectors lose their bends
    // but position and endpoints remain correct.  Arrow markers are applied when present.
    case 'connector': {
      const preset = shape.preset || '';
      const isStraight = !preset || preset.startsWith('straight') || preset === 'line';
      // For straight connectors, the single line uses the full bounding box diagonal.
      // Elbow connectors render as a straight line (approximation without bend-point data).
      primitive = `<line x1="0" y1="0" x2="${wPx}" y2="${hPx}"${headAttr}${tailAttr}/>`;
      void isStraight; // future: render elbow segments when bend data is available
      break;
    }

    // Arrow shapes — 7-point polygon, direction from original PPTX preset name.
    case 'arrow': {
      const preset    = shape.preset || 'rightArrow';
      const direction = ARROW_DIRECTION[preset] || 'right';
      primitive = `<polygon points="${arrowPolygon(wPx, hPx, direction)}"/>`;
      break;
    }

    // Callout — render as rounded rectangle (pointer not implemented).
    case 'callout':
      primitive = emitRect(wPx, hPx, { rx: Math.min(wPx, hPx) * 0.05, ry: Math.min(wPx, hPx) * 0.05 });
      break;

    // Stubs — shapes where we need the full OOXML formula engine (arc, cloud, star, etc.).
    // Render as a tinted bounding-box rect so the colour/position is at least visible.
    case 'polyline':
    case 'polygon':
    case 'unknown':
      // Emit nothing for truly unknown shapes — avoids clutter for decorative unknowns.
      return '';

    default:
      return '';
  }

  // Collect all defs (gradient + arrow markers) into one ctx bucket for renderShape.
  if (ctx) {
    if (gradDefs)    ctx.extraDefs = (ctx.extraDefs || '') + gradDefs;
    if (markerDefs)  ctx.extraDefs = (ctx.extraDefs || '') + markerDefs;
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
    (fillOpacity < 1 ? ` fill-opacity="${fillOpacity}"` : '') +
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
  const ctx = { warnings: [], extraDefs: '' };
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
    (ctx.extraDefs || '') +
    g +
    `</svg>`
  );
}

module.exports = { emitShape, renderShape };