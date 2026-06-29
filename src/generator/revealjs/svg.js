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

// ---------------------------------------------------------------------------
// Font-shrink safety net — glyph-accurate auto-fit for shape text
// ---------------------------------------------------------------------------

// Lazy-loaded fontkit reference. null = not installed, skip shrink entirely.
let _fontkitLoaded = false;
let _fontkit       = null;
function getFontkit() {
  if (!_fontkitLoaded) {
    _fontkitLoaded = true;
    try { _fontkit = require('fontkit'); } catch { _fontkit = null; }
  }
  return _fontkit;
}

// Directories searched in priority order for system font files.
// Set SYSTEM_FONT_DIR env var to prepend a custom path.
const SYSTEM_FONT_DIRS = [
  process.env.SYSTEM_FONT_DIR,
  'C:/Windows/Fonts',
  '/usr/share/fonts/truetype/msttcorefonts',
  '/usr/share/fonts/truetype',
  '/System/Library/Fonts',
  '/usr/local/share/fonts',
].filter(Boolean);

/**
 * Try to load a font family from system font directories.
 * Returns a fontkit Font object, or null when the file cannot be found.
 */
function tryLoadFont(family, bold) {
  const fk = getFontkit();
  if (!fk) return null;
  /* eslint-disable global-require */
  const nodePath = require('path');
  const nodeFs   = require('fs');
  /* eslint-enable global-require */
  const stem     = family.toLowerCase().replace(/\s+/g, '');
  const variants = bold
    ? [stem + 'b', stem + 'bd', stem + '-bold', stem]
    : [stem, stem + '-regular', stem + 'r'];
  for (const dir of SYSTEM_FONT_DIRS) {
    for (const name of variants) {
      for (const ext of ['.ttf', '.otf']) {
        const fp = nodePath.join(dir, name + ext);
        try {
          const buf = nodeFs.readFileSync(fp);
          return fk.create(buf);
        } catch { /* try next candidate */ }
      }
    }
  }
  return null;
}

/**
 * Simulate CSS word-wrap using fontkit glyph advance widths scaled to CSS pixels.
 * Matches browser behaviour: words wider than availWidthPx occupy their own line
 * rather than breaking mid-word.
 *
 * @param {string}  text          plain text for one paragraph (all runs joined)
 * @param {object}  font          fontkit Font object
 * @param {number}  fontSizePt    font size in points
 * @param {number}  availWidthPx  available content width in CSS pixels
 * @returns {number}  line count (≥ 1)
 */
function simulateLines(text, font, fontSizePt, availWidthPx) {
  const scale = (fontSizePt * PT_TO_PX) / font.unitsPerEm;
  function tokenPx(str) {
    const run = font.layout(str);
    let w = 0;
    for (const g of run.glyphs) w += g.advanceWidth;
    return w * scale;
  }
  const tokens = text.split(/(\s+)/); // keeps whitespace tokens in the array
  let lines = 1;
  let x     = 0;
  for (const tok of tokens) {
    if (!tok) continue;
    const tw = tokenPx(tok);
    if (/^\s/.test(tok)) {
      x += tw; // whitespace: accumulate, never force a break on its own
    } else if (x > 0 && x + tw > availWidthPx) {
      lines++;
      x = tw;  // word wraps to a new line
    } else {
      x += tw; // word fits on the current line
    }
  }
  return lines;
}

const SHRINK_MIN_PT  = 6;
const SHRINK_STEP_PT = 0.5;

/**
 * Determine the font size (pt) that fits all paragraph text within availHeightPx.
 * Reduces from the declared size in 0.5pt steps down to SHRINK_MIN_PT.
 *
 * @param {object[]} paragraphs    IR paragraph array
 * @param {number}   availWidthPx  content area width in CSS px
 * @param {number}   availHeightPx content area height in CSS px
 * @param {object}  [fontOverride] fontkit Font to use instead of tryLoadFont
 *                                 (pass null explicitly to simulate "font not found")
 * @returns {{ origPt: number, shrunkPt: number, font: object }|null}
 *   null → font unavailable or mixed sizes — caller uses CSS as-is.
 */
function measureAndShrink(paragraphs, availWidthPx, availHeightPx, fontOverride) {
  // Collect a uniform font size; bail on mixed-size shapes (can't shrink uniformly).
  let origPt = null;
  let family = 'Calibri';
  let bold   = false;
  for (const para of paragraphs) {
    for (const run of (para.runs || [])) {
      const fmt = run.formatting || {};
      if (fmt.size && typeof fmt.size === 'string' && fmt.size.endsWith('pt')) {
        const pt = parseFloat(fmt.size);
        if (!Number.isNaN(pt)) {
          if (origPt !== null && Math.abs(pt - origPt) > 0.01) return null;
          origPt = pt;
        }
      }
      if (fmt.font)              family = fmt.font;
      if (fmt.weight === 'bold') bold   = true;
    }
  }
  if (!origPt) return null;

  // Line-height ratio from the paragraph's line-spacing (unitless = ratio; fallback 1.2).
  let lhRatio = 1.2;
  for (const para of paragraphs) {
    const ls = (para.formatting || {})['line-spacing'];
    if (ls && !ls.endsWith('pt')) {
      const v = parseFloat(ls);
      if (!Number.isNaN(v)) { lhRatio = v; break; }
    }
  }

  // fontOverride=undefined → use system; fontOverride=null → caller says "not found".
  const font = fontOverride !== undefined ? fontOverride : tryLoadFont(family, bold);
  if (!font) return null;

  const texts = paragraphs.map((p) => (p.runs || []).map((r) => r.text || '').join(''));

  let curPt = origPt;
  for (;;) {
    const lineHPx    = curPt * PT_TO_PX * lhRatio;
    const totalLines = texts.reduce(
      (sum, t) => sum + simulateLines(t, font, curPt, availWidthPx), 0,
    );
    if (totalLines * lineHPx <= availHeightPx || curPt <= SHRINK_MIN_PT) break;
    curPt -= SHRINK_STEP_PT;
  }
  return { origPt, shrunkPt: curPt, font };
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

  // Browser <p> default margins (~1em top+bottom) create excessive gaps inside shapes.
  // Apply margin:0 for any paragraph that has no explicit PPTX space-before/space-after.
  const body = paragraphs.map((para) => {
    const fmt = para.formatting || {};
    if (fmt['space-before'] == null || fmt['space-after'] == null) {
      const newFmt = { ...fmt };
      if (newFmt['space-before'] == null) newFmt['space-before'] = '0pt';
      if (newFmt['space-after']  == null) newFmt['space-after']  = '0pt';
      return renderParagraph({ ...para, formatting: newFmt });
    }
    return renderParagraph(para);
  }).join('');
  const anchor = (!isArray && textOrParagraphs && textOrParagraphs.anchor) || 't';
  const ins    = (!isArray && textOrParagraphs && textOrParagraphs.insets) || null;

  // Convert EMU insets to px (1 px = 9525 EMU); fall back to PPTX defaults.
  const lPx = Math.round((ins ? ins.l : 91440) / 9525);
  const rPx = Math.round((ins ? ins.r : 91440) / 9525);
  const tPx = Math.round((ins ? ins.t : 45720) / 9525);
  const bPx = Math.round((ins ? ins.b : 45720) / 9525);

  const justify = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';

  // Single-quoted names are safe inside the double-quote-delimited style="…"
  // attribute.  Double quotes inside the attribute value would close it early.
  // Humanist fallbacks approximate Calibri's metrics on systems where it is absent.
  const SHAPE_FONT_STACK = "Calibri,'Gill Sans MT','Gill Sans','Trebuchet MS'," +
    "'Liberation Sans',Arial,sans-serif";
  const divStyle = [
    'width:100%', 'height:100%',
    'overflow:hidden', 'box-sizing:border-box',
    `padding:${tPx}px ${rPx}px ${bPx}px ${lPx}px`,
    'display:flex', 'flex-direction:column',
    `justify-content:${justify}`,
    `font-family:${SHAPE_FONT_STACK}`,
  ].join(';');

  // Upgrade spans whose font-family is the bare stack head ("Calibri") to the
  // full fallback stack.  escapeCss strips quotes, so the stack cannot go
  // through formattingToCss — a targeted regex replace is the safe path.
  // Lookahead (?=;|") matches the '; ' separator or the closing attribute '"',
  // which avoids matching longer names such as "Calibri Light".
  const bodyWithStack = body.replace(
    /font-family: Calibri(?=;|")/g,
    `font-family: ${SHAPE_FONT_STACK}`,
  );

  // Font-size shrink safety net: if the measured line count × line-height
  // exceeds the available height, reduce font size in 0.5pt steps until it
  // fits.  Falls back to the CSS-as-emitted when fontkit is unavailable or
  // the shape uses mixed font sizes.
  const availW  = wPx - lPx - rPx;
  const availH  = hPx - tPx - bPx;
  const shrink  = measureAndShrink(paragraphs, availW, availH);
  let bodyFinal = bodyWithStack;
  if (shrink && shrink.shrunkPt < shrink.origPt) {
    const origPx = Math.round(shrink.origPt  * PT_TO_PX);
    const newPx  = Math.round(shrink.shrunkPt * PT_TO_PX);
    bodyFinal = bodyWithStack.replace(
      new RegExp(`font-size: ${origPx}px`, 'g'),
      `font-size: ${newPx}px`,
    );
  }

  // foreignObject requires explicit XHTML namespace on its HTML root.
  // overflow="hidden" (SVG attribute) clips at the viewport boundary (= hPx),
  // so no inner max-height is needed — removing it avoids clipping the last
  // line's ink when Calibri's em-box slightly exceeds the computed line height.
  return (
    `<foreignObject x="0" y="0" width="${wPx}" height="${hPx}" overflow="hidden">` +
    `<div xmlns="http://www.w3.org/1999/xhtml" style="${divStyle}">` +
    `<div style="min-height:0;overflow:hidden">` +
    bodyFinal +
    `</div>` +
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

// Arrow shape builders. Coordinates are in strokeWidth units (no viewBox).
// ml = length along arrow axis, mh = half-width perpendicular (so full height = 2*mh).
function arrowMarkerShape(type, c, ml, mh) {
  switch (type) {
    case 'triangle': return `<polygon points="0 0, ${ml} ${mh}, 0 ${2*mh}" fill="${c}"/>`;
    case 'stealth':  return `<polygon points="0 0, ${ml} ${mh}, 0 ${2*mh}, ${ml*0.5} ${mh}" fill="${c}"/>`;
    case 'arrow':    return `<polyline points="0 0, ${ml} ${mh}, 0 ${2*mh}" fill="none" stroke="${c}" stroke-width="1.5"/>`;
    case 'diamond':  return `<polygon points="${ml*0.5} 0, ${ml} ${mh}, ${ml*0.5} ${2*mh}, 0 ${mh}" fill="${c}"/>`;
    case 'oval':     return `<ellipse cx="${ml*0.5}" cy="${mh}" rx="${ml*0.5}" ry="${mh}" fill="${c}"/>`;
    default: return '';
  }
}

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

  function makeMarker(end, id, atTip, orient) {
    if (!end || !end.type || end.type === 'none') return '';
    // Size in strokeWidth units: ml = length along axis, mh = half perpendicular height.
    const szMap = { sm: 1, med: 1.5, lg: 2 };
    const ml = szMap[end.length] || 3;
    const mh = szMap[end.width]  || 3;
    const shape = arrowMarkerShape(end.type, color, ml, mh);
    if (!shape) return '';
    // atTip=true  → tip (x=ml) anchored to path endpoint  → refX=ml
    // atTip=false → base (x=0) anchored to path startpoint → refX=0
    const refX = atTip ? ml : 0;
    const marker = (
      `<marker id="${id}"` +
      ` markerWidth="${ml + 1}" markerHeight="${2 * mh + 1}"` +
      ` refX="${refX}" refY="${mh}" orient="${orient}" markerUnits="strokeWidth">` +
      shape +
      `</marker>`
    );
    defs.push(marker);
    return `url(#${id})`;
  }

  // headEnd = start of the line (refX=0, orient=auto-start-reverse so arrow faces inward)
  if (headEnd && headEnd.type && headEnd.type !== 'none') {
    headAttr = makeMarker(headEnd, `mh-${safeId}`, false, 'auto-start-reverse')
      ? ` marker-start="url(#mh-${safeId})"` : '';
    // Re-check: makeMarker already pushes to defs; we just need the attr string
    if (defs.length) headAttr = ` marker-start="url(#mh-${safeId})"`;
  }
  const defsBeforeTail = defs.length;
  // tailEnd = end of the line (refX=10, orient=auto)
  if (tailEnd && tailEnd.type && tailEnd.type !== 'none') {
    makeMarker(tailEnd, `mt-${safeId}`, true, 'auto');
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
  const w = wPx;
  const h = hPx;
  switch (direction) {
    case 'left':
      return `${w},${h * 0.25} ${w * 0.35},${h * 0.25} ${w * 0.35},0 0,${h * 0.5} ${w * 0.35},${h} ${w * 0.35},${h * 0.75} ${w},${h * 0.75}`;
    case 'up':
      return `${w * 0.25},${h} ${w * 0.25},${h * 0.35} 0,${h * 0.35} ${w * 0.5},0 ${w},${h * 0.35} ${w * 0.75},${h * 0.35} ${w * 0.75},${h}`;
    case 'down':
      return `${w * 0.25},0 ${w * 0.25},${h * 0.65} 0,${h * 0.65} ${w * 0.5},${h} ${w},${h * 0.65} ${w * 0.75},${h * 0.65} ${w * 0.75},0`;
    default: // right
      return `0,${h * 0.25} ${w * 0.65},${h * 0.25} ${w * 0.65},0 ${w},${h * 0.5} ${w * 0.65},${h} ${w * 0.65},${h * 0.75} 0,${h * 0.75}`;
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
  return `<path d="M ${x1},${y1} A ${rx.toFixed(2)},${ry.toFixed(2)} 0 ${largeArc} 1 ${x2},${y2}" fill="none" stroke-linecap="butt"/>`;
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
 * Approximate OOXML cloud shape using SVG arc commands.
 * Traces 6 arcs: up the left side, across the bumpy top, down the right side,
 * then a straight bottom edge closes the silhouette.
 */
function emitCloud(wPx, hPx) {
  const p = (x, y) => `${(x * wPx).toFixed(1)},${(y * hPx).toFixed(1)}`;
  const r = (rx, ry) => `${(rx * wPx).toFixed(1)},${(ry * hPx).toFixed(1)}`;
  // Approximate OOXML cloud: 7 outward-bumping arcs (CW path, sweep=1) + curved bottom arc.
  // Bumps: left-low → left-mid → top-left → top-center → top-right → right-mid → right-low
  // Bottom: wide shallow arc closes the shape (no straight line).
  return (
    `<path d="` +
    `M ${p(0.08, 0.86)} ` +
    `A ${r(0.08, 0.12)} 0 0 1 ${p(0.04, 0.67)} ` +
    `A ${r(0.12, 0.18)} 0 0 1 ${p(0.13, 0.43)} ` +
    `A ${r(0.17, 0.25)} 0 0 1 ${p(0.34, 0.20)} ` +
    `A ${r(0.19, 0.28)} 0 0 1 ${p(0.56, 0.10)} ` +
    `A ${r(0.17, 0.25)} 0 0 1 ${p(0.76, 0.20)} ` +
    `A ${r(0.13, 0.19)} 0 0 1 ${p(0.88, 0.44)} ` +
    `A ${r(0.08, 0.12)} 0 0 1 ${p(0.92, 0.66)} ` +
    `A ${r(0.50, 0.16)} 0 0 1 ${p(0.08, 0.86)} ` +
    `Z"/>`
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
 * Database cylinder — Bezier-curve body with a visible top ellipse cap.
 * The body path uses cubic curves so the top and bottom rims are true ellipses.
 * The top cap ellipse overrides the inherited stroke with a semi-transparent
 * white rim to give a 3-D "lid" impression.
 */
function emitDatabase(wPx, hPx) {
  const rimRy = hPx * 0.18; // half-height of the top/bottom ellipse
  return (
    // Cylinder body: open at top, closed Bezier curves top and bottom
    `<path d="M 0,${rimRy}` +
    ` C 0,${-hPx * 0.06} ${wPx},${-hPx * 0.06} ${wPx},${rimRy}` +
    ` L ${wPx},${hPx * 0.82}` +
    ` C ${wPx},${hPx * 1.06} 0,${hPx * 1.06} 0,${hPx * 0.82} Z"/>` +
    // Top lid — same fill as body, distinct rim stroke
    `<ellipse cx="${wPx / 2}" cy="${rimRy}" rx="${wPx / 2}" ry="${rimRy}"` +
    ` stroke="rgba(255,255,255,0.45)" stroke-width="2"/>`
  );
}

/**
 * Callout shape: rectangle (or rounded rectangle) with a triangular tail
 * emanating from the bottom edge.
 *
 * adj1: tail-tip x in 1/100000 units of width  (default 25000 = 25%)
 * adj2: tail-tip y in 1/100000 units of height (default 200000 = 200%, i.e. below shape)
 * adj3: corner radius in 1/100000 units of min(w,h)/2 — only used when rounded=true
 *       (OOXML default 16667, matching roundRect)
 */
function emitWedgeCallout(wPx, hPx, adjustments, rounded) {
  const adjs = adjustments || [];
  const getAdj = (name, def) => {
    if (!Array.isArray(adjs)) return def;
    const entry = adjs.find((a) => a.name === name);
    return entry != null ? entry.value : def;
  };

  const adj1 = getAdj('adj1', 25000);
  const adj2 = getAdj('adj2', 200000);

  // OOXML: adj1/adj2 are offsets from the shape centre (hc, vc), not from (0,0).
  const tx  = wPx / 2 + (adj1 / 100000) * wPx;
  const ty  = hPx / 2 + (adj2 / 100000) * hPx;
  const tw  = wPx * 0.1;
  const tx1 = Math.max(0, tx - tw);
  const tx2 = Math.min(wPx, tx + tw);
  const n   = (v) => v.toFixed(1);

  if (!rounded) {
    const d = `M 0,0 L ${n(wPx)},0 L ${n(wPx)},${n(hPx)} L ${n(tx2)},${n(hPx)} L ${n(tx)},${n(ty)} L ${n(tx1)},${n(hPx)} L 0,${n(hPx)} Z`;
    return `<path d="${d}"/>`;
  }

  const adj3  = getAdj('adj3', 16667);
  const rxRaw = Math.round((adj3 / 100000) * Math.min(wPx, hPx) / 2);
  const rx    = Math.max(0, Math.min(rxRaw, Math.floor(wPx / 2), Math.floor(hPx / 2)));
  // Clamp tail base to the flat region of the bottom edge (outside corner arcs).
  const b1 = Math.max(rx, tx1);
  const b2 = Math.min(wPx - rx, tx2);

  const parts = [`M ${rx},0`, `L ${n(wPx - rx)},0`];
  if (rx > 0) parts.push(`A ${rx},${rx} 0 0 1 ${n(wPx)},${rx}`);
  parts.push(`L ${n(wPx)},${n(hPx - rx)}`);
  if (rx > 0) parts.push(`A ${rx},${rx} 0 0 1 ${n(wPx - rx)},${n(hPx)}`);
  parts.push(`L ${n(b2)},${n(hPx)}`, `L ${n(tx)},${n(ty)}`, `L ${n(b1)},${n(hPx)}`);
  if (b1 > rx) parts.push(`L ${rx},${n(hPx)}`);
  if (rx > 0) parts.push(`A ${rx},${rx} 0 0 1 0,${n(hPx - rx)}`);
  parts.push(`L 0,${rx}`);
  if (rx > 0) parts.push(`A ${rx},${rx} 0 0 1 ${rx},0`);
  parts.push('Z');
  return `<path d="${parts.join(' ')}"/>`;
}

// Map PPTX arrow preset → direction string consumed by arrowPolygon.
const ARROW_DIRECTION = {
  leftArrow:      'left',  leftArrowCallout:  'left',
  upArrow:        'up',    upArrowCallout:    'up',
  downArrow:      'down',  downArrowCallout:  'down',
  rightArrow:     'right', rightArrowCallout: 'right',
  // homePlate → approximated as right arrow; chevron now has its own case
  homePlate: 'right',
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
    (type === 'line' || type === 'connector' || type === 'arc')
      ? buildArrowMarkers(shape.stroke, shape.id || String(xPx + yPx))
      : { markerDefs: '', headAttr: '', tailAttr: '' };

  let primitive;
  let rrClipId  = null; // set in roundRect case; consumed in inner-content assembly
  let rrClipDef = '';
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
      // Build a clipPath matching the rounded rect so the foreignObject text
      // is clipped to the visible rounded boundary, not just the rectangular box.
      rrClipId  = `rrclip-${(shape.id || 'rrect').replace(/[^a-zA-Z0-9_-]/g, '-')}`;
      rrClipDef = `<defs><clipPath id="${rrClipId}">` +
        `<rect x="0" y="0" width="${wPx}" height="${hPx}" rx="${rxRaw}" ry="${rxRaw}"/>` +
        `</clipPath></defs>`;
      break;
    }

    case 'ellipse':
      primitive = `<ellipse cx="${wPx / 2}" cy="${hPx / 2}" rx="${wPx / 2}" ry="${hPx / 2}"/>`;
      break;

    case 'triangle':
      // Isosceles triangle filling the full bounding box: top-center, bottom-right, bottom-left
      primitive = `<polygon points="${wPx / 2},0 ${wPx},${hPx} 0,${hPx}"/>`;
      break;

    case 'rtTriangle':
      // Right angle at bottom-left: (0,0) top-left, (0,h) bottom-left, (w,h) bottom-right
      primitive = `<polygon points="0,0 0,${hPx} ${wPx},${hPx}"/>`;
      break;

    case 'chevron': {
      // OOXML chevron: right-facing shape with a V-notch on the left and arrow point on the right.
      // adj (default 50000 = 50%) controls the notch depth / arrowhead width.
      // OOXML formula: x1 = w*adj/100000 (notch tip x), x2 = w - x1 (arrowhead base x)
      // Points (CW): (0,0) → (x2,0) → (w,h/2) → (x2,h) → (0,h) → (x1,h/2)
      const chevAdj = (shape.adjustments && shape.adjustments[0] != null)
        ? shape.adjustments[0].value : 50000;
      const chevX1 = Math.min(wPx, hPx) * (chevAdj / 100000);  // notch depth bounded by height (matches PowerPoint)
      const chevX2 = wPx - chevX1;                // arrowhead base x
      primitive = `<polygon points="0,0 ${chevX2},0 ${wPx},${hPx / 2} ${chevX2},${hPx} 0,${hPx} ${chevX1},${hPx / 2}"/>`;
      break;
    }

    case 'pentagon': {
      // OOXML pentagon: right-pointing arrow-pentagon shape.
      // adj = width of the right point as a fraction of width (default 50000 = 50%).
      const pentAdj = (shape.adjustments && shape.adjustments[0] != null)
        ? shape.adjustments[0].value : 50000;
      const pentA = Math.min(wPx, hPx) * (pentAdj / 100000);  // arrowhead width bounded by height (matches PowerPoint)
      primitive = `<polygon points="0,0 ${wPx - pentA},0 ${wPx},${hPx / 2} ${wPx - pentA},${hPx} 0,${hPx}"/>`;
      break;
    }

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
    // curvedRightArrow gets a bespoke cubic-bezier path + arrowhead polygon.
    case 'arrow': {
      const preset = shape.preset || 'rightArrow';

      if (preset === 'curvedRightArrow') {
        // Gradient: 40% opacity at the start (top) → 100% at the arrowhead (bottom).
        // stop-color must use style="" when fill is a CSS var() to support theme colours.
        const craId = `cra-${(shape.id || 'x').replace(/[^a-zA-Z0-9_-]/g, '-')}`;
        const mkStop = (pct, op) => fill && fill.startsWith('var(')
          ? `<stop offset="${pct}%" style="stop-color:${fill};stop-opacity:${op}"/>`
          : `<stop offset="${pct}%" stop-color="${escapeHtml(fill || '#e00')}" stop-opacity="${op}"/>`;
        const craDefs =
          `<defs><linearGradient id="${craId}" x1="0" y1="0" x2="0" y2="1" gradientUnits="objectBoundingBox">` +
          mkStop(0, 0.4) + mkStop(100, 1) +
          `</linearGradient></defs>`;

        const aw = Math.min(wPx, hPx) * 0.20;
        // Curve redesigned to end pointing RIGHT (→) at (0.50w, 0.90h):
        //   start top-right → sweep left → come back right at ~90% height.
        // Stroke stops at arrowhead base (0.50w) to avoid overlapping it.
        // Arrowhead: right-pointing triangle, base at (0.50w, 0.90h).
        primitive = craDefs +
          `<path` +
          ` d="M ${wPx * 0.90} ${hPx * 0.10}` +
          ` C ${wPx * 0.12} ${hPx * 0.04}, ${wPx * -0.08} ${hPx * 0.38}, ${wPx * 0.08} ${hPx * 0.66}` +
          ` C ${wPx * 0.17} ${hPx * 0.83}, ${wPx * 0.34} ${hPx * 0.90}, ${wPx * 0.50} ${hPx * 0.90}"` +
          ` fill="none" stroke="url(#${craId})" stroke-width="${aw}" stroke-linecap="butt"/>` +
          `<polygon` +
          ` points="${wPx * 0.50},${hPx * 0.78} ${wPx * 0.95},${hPx * 0.90} ${wPx * 0.50},${hPx * 1.02}"` +
          ` fill="url(#${craId})"/>`;
        break;
      }

      const direction = ARROW_DIRECTION[preset] || 'right';
      primitive = `<polygon points="${arrowPolygon(wPx, hPx, direction)}"/>`;
      break;
    }

    // Callout — dispatch to specific geometry based on PPTX preset name.
    case 'callout': {
      const cPreset = shape.preset || '';
      if (cPreset === 'cloudCallout') {
        const cAdjs = shape.adjustments || [];
        const cGetAdj = (name, def) => {
          if (!Array.isArray(cAdjs)) return def;
          const e = cAdjs.find((a) => a.name === name);
          return e != null ? e.value : def;
        };
        // adj1/adj2 are offsets from the shape centre (hc, vc), not from (0,0).
        const tipX = wPx / 2 + (cGetAdj('adj1', -20000) / 100000) * wPx;
        // Thought-bubble tail: 3 shrinking circles that drop DOWNWARD from the
        // cloud bottom, leaning toward whichever side the callout tip is on.
        // We always go downward (ignoring tipY) so the trail looks natural
        // regardless of the adj values stored in the file.
        const leanLeft = tipX < wPx / 2;
        const baseX = leanLeft ? 0.13 * wPx : 0.87 * wPx;
        const leanDx = leanLeft ? -1 : 1;
        // circles: just below cloud body (0.87*h), then further below bounding box
        const tailCircles = [
          { cx: baseX,                    cy: 0.94 * hPx, r: 5.5 },
          { cx: baseX + leanDx * 9,       cy: 1.07 * hPx, r: 3.5 },
          { cx: baseX + leanDx * 16,      cy: 1.17 * hPx, r: 2 },
        ];
        const tail = tailCircles.map(({ cx, cy, r }) =>
          `<circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="${r}"/>`
        ).join('');
        primitive = tail + emitCloud(wPx, hPx);
      } else if (cPreset === 'wedgeRoundRectCallout') {
        primitive = emitWedgeCallout(wPx, hPx, shape.adjustments, true);
      } else if (['wedgeRectCallout', 'callout1', 'callout2', 'callout3', 'callout4'].includes(cPreset)) {
        primitive = emitWedgeCallout(wPx, hPx, shape.adjustments, false);
      } else {
        warnings.push(`callout preset "${cPreset}" not yet implemented, using rect fallback`);
        primitive = `<rect x="0" y="0" width="${wPx}" height="${hPx}" data-preset="${escapeHtml(cPreset)}"/>`;
      }
      break;
    }

    // Arc — open ellipse arc from adj1 angle to adj2 angle (CW).
    case 'arc': {
      const adjs = shape.adjustments || [];
      const getA = (name, def) => (Array.isArray(adjs) ? (adjs.find((a) => a.name === name) || {}).value : null) ?? def;
      // Arc = open elliptical segment used as a curved connector.
      // headEnd → marker-start (at path start, i.e. stAng), tailEnd → marker-end.
      // Both come from buildArrowMarkers via headAttr/tailAttr above.
      primitive = emitArc(wPx, hPx, getA('adj1', 16200000), getA('adj2', 0))
        .replace(/\/>$/, `${headAttr}${tailAttr}/>`);
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

    // Database cylinder (PPTX flowChartMagneticDisk preset).
    case 'database':
      primitive = emitDatabase(wPx, hPx);
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

  // Collect all defs (gradient + arrow markers + roundRect clip) into one ctx bucket.
  if (ctx) {
    if (gradDefs)    ctx.extraDefs = (ctx.extraDefs || '') + gradDefs;
    if (markerDefs)  ctx.extraDefs = (ctx.extraDefs || '') + markerDefs;
    if (rrClipDef)   ctx.extraDefs = (ctx.extraDefs || '') + rrClipDef;
  }

  // Pass the full text object so emitForeignObject can apply anchor + insets.
  // Handles both new IR {id,paragraphs,anchor,insets} and old IR plain array.
  // Arc shapes are geometric connectors — never render a text overlay.
  const fo = (type === 'arc') ? '' : emitForeignObject(shape.text || [], wPx, hPx);
  // For roundRect shapes wrap contents in a clipping group so text is clipped
  // to the rounded boundary, not just the rectangular foreignObject box.
  let inner;
  if (rrClipId) {
    inner = fo
      ? `\n  <g clip-path="url(#${rrClipId})">\n    ${primitive}\n    ${fo}\n  </g>\n`
      : `\n  <g clip-path="url(#${rrClipId})">\n    ${primitive}\n  </g>\n`;
  } else {
    inner = fo
      ? `\n  ${primitive}\n  ${fo}\n`
      : `\n  ${primitive}\n`;
  }

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

/**
 * For arc shapes with a headEnd arrowhead: compute the slide-space position
 * and direction of the arrowhead and return a separate SVG element at the
 * arc's natural z-index (floats above connected shapes that cover the arc body).
 */
function buildArcArrowhead(shape, xPx, yPx, wPx, hPx, rotation, strokeWidthPx, zIndex, allShapes) {
  const he = shape.stroke && shape.stroke.headEnd;
  if (!he || he.type === 'none') return '';

  const adjs = shape.adjustments || [];
  const getA = (name, def) =>
    (Array.isArray(adjs) ? (adjs.find((a) => a.name === name) || {}).value : null) ?? def;
  const adj1 = getA('adj1', 16200000);
  const adj2 = getA('adj2', 0);

  const cx = wPx / 2, cy = hPx / 2;
  const rx = wPx / 2, ry = hPx / 2;
  const a1 = (adj1 / 60000) * (Math.PI / 180);
  const a2 = (adj2 / 60000) * (Math.PI / 180);

  // Rotation matrix
  const rotRad = rotation * Math.PI / 180;
  const cr = Math.cos(rotRad), sr = Math.sin(rotRad);

  // Convert arc angle → slide-space point
  const arcPoint = (a) => ({
    x: xPx + cx + rx * Math.cos(a) * cr - ry * Math.sin(a) * sr,
    y: yPx + cy + rx * Math.cos(a) * sr + ry * Math.sin(a) * cr,
  });

  // Arc path START slide-space position
  const p1 = arcPoint(a1);
  const sx = p1.x, sy = p1.y;

  // Marker size in px
  const szMap = { sm: 1, med: 1.5, lg: 2 };
  const ml = (szMap[he.length] || 1.5) * strokeWidthPx;
  const mh = (szMap[he.width]  || 1.5) * strokeWidthPx;
  const color = colorToCss(shape.stroke.color) || '#000';

  // Default: arrowhead at the mathematical arc start, using CW tangent at a1
  const ltx0 = rx * Math.sin(a1), lty0 = -ry * Math.cos(a1);
  const tlen0 = Math.sqrt(ltx0 * ltx0 + lty0 * lty0);
  const stx0 = (ltx0 / tlen0) * cr - (lty0 / tlen0) * sr;
  const sty0 = (ltx0 / tlen0) * sr + (lty0 / tlen0) * cr;

  let bx = sx, by = sy;
  // Arrow points INTO connected shape (forward along arc at start)
  let arrowAngle = Math.atan2(sty0, stx0) * 180 / Math.PI;

  // Find the covering shape and binary-search for the exact arc exit point
  if (Array.isArray(allShapes)) {
    for (const s of allShapes) {
      if (s === shape || !s.position) continue;
      const sz = typeof s['z-index'] === 'number' ? s['z-index'] : 0;
      if (sz <= 0 || sz > zIndex) continue;
      const sRot = typeof s.rotation === 'number' ? s.rotation : 0;
      if (Math.abs(sRot) > 1 && Math.abs(sRot - 360) > 1) continue; // skip rotated shapes
      const shX = emuToPx(s.position.x) ?? 0;
      const shY = emuToPx(s.position.y) ?? 0;
      const shW = emuToPx(s.position.w ?? s.width) ?? 0;
      const shH = emuToPx(s.position.h ?? s.height) ?? 0;
      const inside = (p) => p.x >= shX && p.x <= shX + shW && p.y >= shY && p.y <= shY + shH;
      if (!inside({ x: sx, y: sy })) continue;
      if (inside(arcPoint(a2))) continue; // arc end also inside — skip

      // Binary search: 40 iterations find exit angle to sub-pixel accuracy
      let aIn = a1, aOut = a2;
      for (let i = 0; i < 40; i++) {
        const aMid = (aIn + aOut) / 2;
        if (inside(arcPoint(aMid))) aIn = aMid; else aOut = aMid;
      }
      const aBound = (aIn + aOut) / 2;
      const pb = arcPoint(aBound);
      bx = pb.x; by = pb.y;

      // CW tangent at aBound in slide space (exit direction)
      const stxB = -rx * Math.sin(aBound) * cr - ry * Math.cos(aBound) * sr;
      const styB = -rx * Math.sin(aBound) * sr + ry * Math.cos(aBound) * cr;
      // Arrowhead points INTO shape = opposite of exit direction
      arrowAngle = Math.atan2(-styB, -stxB) * 180 / Math.PI;
      break;
    }
  }

  // Polygon: (0,0)-(ml,mh)-(0,2mh); ref point (0,mh) placed at boundary point
  const ar = arrowAngle * Math.PI / 180;
  const tx = bx + mh * Math.sin(ar);
  const ty = by - mh * Math.cos(ar);
  const pts = `0 0, ${ml.toFixed(2)} ${mh.toFixed(2)}, 0 ${(2 * mh).toFixed(2)}`;
  const xfm = `translate(${tx.toFixed(2)}, ${ty.toFixed(2)}) rotate(${arrowAngle.toFixed(2)})`;

  const style = [
    'position:absolute', 'left:0', 'top:0', 'width:100%', 'height:100%',
    'overflow:visible', 'pointer-events:none', `z-index:${zIndex}`,
  ].join(';');

  return `<svg style="${style}" xmlns="http://www.w3.org/2000/svg">` +
    `<polygon points="${pts}" fill="${color}" transform="${xfm}"/>` +
    `</svg>`;
}

function renderShape(shape, allShapes) {
  const ctx = { warnings: [], extraDefs: '' };
  const g = emitShape(shape, ctx);
  if (!g) return '';

  const rawZIndex = typeof shape['z-index'] === 'number' ? shape['z-index'] : 0;
  // Lines, connectors, and open arcs are drawn behind filled shapes so their
  // endpoints are clipped by the shapes they connect to.
  // Arc arrowheads are emitted as a separate high-z SVG placed at the shape boundary.
  const isWire = shape.type === 'connector' || shape.type === 'line' ||
    (shape.type === 'arc' && shape.fill && shape.fill.type === 'none');
  const zIndex = isWire ? Math.min(rawZIndex, 0) : rawZIndex;
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

  let arcArrowSvg = '';
  if (shape.type === 'arc' && shape.stroke && shape.stroke.headEnd &&
      shape.stroke.headEnd.type !== 'none') {
    const _xPx = emuToPx(shape.position && shape.position.x) ?? 0;
    const _yPx = emuToPx(shape.position && shape.position.y) ?? 0;
    const _wPx = emuToPx(shape.width ?? (shape.position && shape.position.w)) ?? 0;
    const _hPx = emuToPx(shape.height ?? (shape.position && shape.position.h)) ?? 0;
    const _rawRot = typeof shape.rotation === 'number' ? shape.rotation : 0;
    const _rotation = _rawRot > 360 ? _rawRot / 60000 : _rawRot;
    const _sw = strokeAttrs(shape.stroke).widthPx || 4;
    arcArrowSvg = buildArcArrowhead(
      shape, _xPx, _yPx, _wPx, _hPx, _rotation, _sw, rawZIndex, allShapes
    );
  }

  return (
    `<svg style="${style}" xmlns="http://www.w3.org/2000/svg">` +
    (ctx.extraDefs || '') +
    g +
    `</svg>` +
    arcArrowSvg
  );
}

module.exports = { emitShape, renderShape, simulateLines, measureAndShrink };