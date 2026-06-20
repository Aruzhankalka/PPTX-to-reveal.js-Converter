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

/**
 * Render embedded shape text as an SVG <foreignObject>.
 *
 * Accepts either the full text object {paragraphs, anchor, insets} (new IR)
 * or a plain paragraph array (old IR / fallback).
 *
 * anchor:  'ctr' → vertically centred; 'b' → bottom-aligned; default top.
 * insets:  body padding in EMU (PPTX default: l=r=91440, t=b=45720 ≈ 10/5 px).
 * Text is clipped to the shape bounding box (overflow:hidden) so it does not
 * bleed into neighbouring shapes.
 */
function emitForeignObject(textOrParagraphs, wPx, hPx) {
  const isArray    = Array.isArray(textOrParagraphs);
  const paragraphs = isArray ? textOrParagraphs : (textOrParagraphs && textOrParagraphs.paragraphs);
  if (!paragraphs || paragraphs.length === 0) return '';

  const body   = paragraphs.map(renderParagraph).join('');
  const anchor = (!isArray && textOrParagraphs && textOrParagraphs.anchor) || 't';
  const ins    = (!isArray && textOrParagraphs && textOrParagraphs.insets) || null;

  // Convert EMU insets to px (1 px = 9525 EMU); fall back to PPTX defaults.
  const lPx = Math.round((ins ? ins.l : 91440) / 9525);
  const rPx = Math.round((ins ? ins.r : 91440) / 9525);
  const tPx = Math.round((ins ? ins.t : 45720) / 9525);
  const bPx = Math.round((ins ? ins.b : 45720) / 9525);

  const justify = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';

  const divStyle = [
    'width:100%', 'height:100%',
    'overflow:hidden', 'box-sizing:border-box',
    `padding:${tPx}px ${rPx}px ${bPx}px ${lPx}px`,
    'display:flex', 'flex-direction:column',
    `justify-content:${justify}`,
  ].join(';');

  // foreignObject requires explicit XHTML namespace on its HTML root.
  return (
    `<foreignObject x="0" y="0" width="${wPx}" height="${hPx}">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="${divStyle}">` +
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

// ---------------------------------------------------------------------------
// Additional shape helpers
// ---------------------------------------------------------------------------

/**
 * Open arc on an ellipse.
 * adj1/adj2: PPTX angles in 1/60000 degrees, 0=east, clockwise.
 * Default: adj1=16200000 (270°=north), adj2=0 (0°=east) → quarter arc upper-right.
 */
function emitArc(wPx, hPx, adj1, adj2) {
  const cx = wPx / 2, cy = hPx / 2;
  const rx = wPx / 2, ry = hPx / 2;
  const toRad = (a) => (a / 60000) * (Math.PI / 180);
  const a1 = toRad(adj1);
  const a2 = toRad(adj2);
  const x1 = (cx + rx * Math.cos(a1)).toFixed(2);
  const y1 = (cy + ry * Math.sin(a1)).toFixed(2);
  const x2 = (cx + rx * Math.cos(a2)).toFixed(2);
  const y2 = (cy + ry * Math.sin(a2)).toFixed(2);
  // CW swing angle — if adj2 ≤ adj1, the arc wraps around 0°/360°
  let swingDeg = (adj2 / 60000) - (adj1 / 60000);
  if (swingDeg <= 0) swingDeg += 360;
  const largeArc = swingDeg > 180 ? 1 : 0;
  // Open arc (no Z) — fill:none handles the "line only" appearance automatically
  return `<path d="M ${x1},${y1} A ${rx.toFixed(2)},${ry.toFixed(2)} 0 ${largeArc} 1 ${x2},${y2}"/>`;
}

/**
 * N-pointed star polygon.
 * innerRatio: inner vertex radius as a fraction of the outer radius (0–1).
 */
function emitStar(wPx, hPx, points, innerRatio) {
  const cx = wPx / 2, cy = hPx / 2;
  const outerRx = wPx / 2, outerRy = hPx / 2;
  const innerRx = outerRx * innerRatio, innerRy = outerRy * innerRatio;
  const pts = [];
  for (let i = 0; i < points * 2; i++) {
    const angle = (i * Math.PI / points) - Math.PI / 2;
    const isOuter = i % 2 === 0;
    const rx = isOuter ? outerRx : innerRx;
    const ry = isOuter ? outerRy : innerRy;
    pts.push(`${(cx + rx * Math.cos(angle)).toFixed(2)},${(cy + ry * Math.sin(angle)).toFixed(2)}`);
  }
  return `<polygon points="${pts.join(' ')}"/>`;
}

/**
 * Approximate OOXML cloud shape using bezier curves.
 * The OOXML cloud requires the formula engine; this gives a recognisable
 * cloud silhouette without implementing ~80 guide variables.
 */
function emitCloud(wPx, hPx) {
  const f = (x, y) => `${(x * wPx).toFixed(1)},${(y * hPx).toFixed(1)}`;
  return (
    `<path d="` +
    `M ${f(0.05,0.65)} ` +
    `Q ${f(-0.02,0.55)} ${f(0.04,0.44)} ` +
    `Q ${f(-0.02,0.24)} ${f(0.18,0.22)} ` +
    `Q ${f(0.16,0.02)} ${f(0.38,0.05)} ` +
    `Q ${f(0.43,-0.03)} ${f(0.53,0.04)} ` +
    `Q ${f(0.60,-0.04)} ${f(0.71,0.07)} ` +
    `Q ${f(0.84,0.00)} ${f(0.93,0.18)} ` +
    `Q ${f(1.04,0.20)} ${f(1.01,0.40)} ` +
    `Q ${f(1.06,0.58)} ${f(0.94,0.66)} ` +
    `Q ${f(0.97,0.90)} ${f(0.76,0.94)} ` +
    `L ${f(0.24,0.94)} ` +
    `Q ${f(0.02,0.92)} ${f(0.05,0.65)} Z` +
    `"/>`
  );
}

/**
 * Render an IR customGeometry object as an SVG <path> string.
 *
 * The custGeom coordinate space (w, h in EMU) is scaled to the shape's pixel
 * bounding box.  Commands are emitted in the order they appear within each
 * path's command array; cross-type ordering between different command tags in
 * the same path may not be preserved (fast-xml-parser limitation noted in the
 * parser) but single-type sequences are correct.
 *
 * Unsupported ops (arcTo — needs wR/hR/stAng/swAng attrs not stored in IR)
 * are skipped with a no-op.
 *
 * @param {object} custGeom  IR customGeometry { w, h, paths }
 * @param {number} wPx       shape width in pixels
 * @param {number} hPx       shape height in pixels
 * @returns {string}  One or more <path> elements, or '' if no paths.
 */
function emitCustomGeometry(custGeom, wPx, hPx) {
  if (!custGeom || !custGeom.paths || custGeom.paths.length === 0) return '';

  const scaleX = custGeom.w > 0 ? wPx / custGeom.w : 1;
  const scaleY = custGeom.h > 0 ? hPx / custGeom.h : 1;
  const sx = (x) => (x * scaleX).toFixed(2);
  const sy = (y) => (y * scaleY).toFixed(2);

  const pathEls = [];
  for (const pathDef of custGeom.paths) {
    const cmds = pathDef.commands || [];
    if (cmds.length === 0) continue;

    const dParts = [];
    for (const cmd of cmds) {
      const pts = cmd.pts || [];
      switch (cmd.op) {
        case 'moveTo':
          if (pts[0]) dParts.push(`M ${sx(pts[0].x)},${sy(pts[0].y)}`);
          break;
        case 'lnTo':
          if (pts[0]) dParts.push(`L ${sx(pts[0].x)},${sy(pts[0].y)}`);
          break;
        case 'cubicBezTo':
          if (pts.length >= 3)
            dParts.push(`C ${sx(pts[0].x)},${sy(pts[0].y)} ${sx(pts[1].x)},${sy(pts[1].y)} ${sx(pts[2].x)},${sy(pts[2].y)}`);
          break;
        case 'quadBezTo':
          if (pts.length >= 2)
            dParts.push(`Q ${sx(pts[0].x)},${sy(pts[0].y)} ${sx(pts[1].x)},${sy(pts[1].y)}`);
          break;
        case 'close':
          dParts.push('Z');
          break;
        case 'arcTo':
          // arcTo needs wR/hR/stAng/swAng which are OOXML formula-evaluated values
          // not stored in the IR; skip gracefully.
          break;
        default:
          break;
      }
    }
    if (dParts.length > 0) {
      pathEls.push(`<path d="${dParts.join(' ')}"/>`);
    }
  }
  return pathEls.join('');
}

/**
 * Flowchart magnetic-disk / database cylinder.
 * Rendered as rect body + top ellipse lid + bottom rim arc.
 */
function emitFlowChartDisk(wPx, hPx) {
  const ry = Math.max(4, Math.round(hPx * 0.12));
  const rx = wPx / 2;
  const bY  = hPx - ry; // bottom ellipse centre y
  return (
    // Body
    `<rect x="0" y="${ry}" width="${wPx}" height="${Math.max(0, hPx - 2 * ry)}"/>` +
    // Top lid (full ellipse)
    `<ellipse cx="${rx}" cy="${ry}" rx="${rx}" ry="${ry}"/>` +
    // Bottom rim (lower half-arc, stroke only)
    `<path d="M 0,${bY} A ${rx},${ry} 0 0 0 ${wPx},${bY}" fill="none"/>`
  );
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
    // OOXML default for adj when avLst is empty: 16667 (1/6 of the shorter side).
    case 'roundRect': {
      const adj = shape.adjustments;
      const adjVal = Array.isArray(adj)
        ? (adj.find((a) => a.name === 'adj') || {}).value
        : (adj && adj.adj);
      // adj value is in 1/100000 units; OOXML default is 16667 when avLst is absent.
      // Convert to pixel radius capped at half the shorter side.
      const ROUNDRECT_DEFAULT_ADJ = 16667;
      const effectiveAdj = adjVal != null ? adjVal : ROUNDRECT_DEFAULT_ADJ;
      const rxRaw = Math.round((effectiveAdj / 100000) * Math.min(wPx, hPx) / 2);
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

    // Arc — open ellipse arc from adj1 angle to adj2 angle (CW).
    case 'arc': {
      const adjs  = shape.adjustments || [];
      const getA  = (name, def) => (Array.isArray(adjs) ? (adjs.find((a) => a.name === name) || {}).value : null) ?? def;
      primitive = emitArc(wPx, hPx, getA('adj1', 16200000), getA('adj2', 0));
      break;
    }

    // Star — N-pointed polygon; inner-radius ratio from adj (OOXML 1/100000 units).
    case 'star': {
      const preset = shape.preset || 'star5';
      const m      = preset.match(/^star(\d+)$/i);
      const n      = m ? parseInt(m[1]) : 5;
      const adjs   = shape.adjustments || [];
      const adjEntry = Array.isArray(adjs) ? adjs.find((a) => a.name === 'adj') : null;
      // OOXML per-star default inner radii (from the spec's fixed-value table)
      const STAR_DEF = { 4:25000, 5:32287, 6:28868, 7:27500, 8:29289, 10:31623,
                         12:33333, 16:35355, 24:37500, 32:38268 };
      const innerRatio = (adjEntry ? adjEntry.value : (STAR_DEF[n] || 30000)) / 100000;
      primitive = emitStar(wPx, hPx, n, innerRatio);
      break;
    }

    // Cloud — bezier approximation (OOXML formula engine not implemented).
    case 'cloud':
      primitive = emitCloud(wPx, hPx);
      break;

    // Flowchart cylinder / magnetic disk.
    case 'flowchartDisk':
      primitive = emitFlowChartDisk(wPx, hPx);
      break;

    // Shapes with IR customGeometry: render the stored path data directly.
    // This covers freeform / scribble shapes whose path was extracted from
    // <a:custGeom> by the parser (type is usually 'unknown').
    case 'polyline':
    case 'polygon':
    case 'unknown':
    default: {
      if (shape.customGeometry) {
        const cg = emitCustomGeometry(shape.customGeometry, wPx, hPx);
        if (cg) { primitive = cg; break; }
      }
      return '';
    }
  }

  // Collect all defs (gradient + arrow markers) into one ctx bucket for renderShape.
  if (ctx) {
    if (gradDefs)    ctx.extraDefs = (ctx.extraDefs || '') + gradDefs;
    if (markerDefs)  ctx.extraDefs = (ctx.extraDefs || '') + markerDefs;
  }

  // Pass the full text object so emitForeignObject can apply anchor + insets.
  // Handles both new IR {id,paragraphs,anchor,insets} and old IR plain array.
  const fo = emitForeignObject(shape.text || [], wPx, hPx);
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