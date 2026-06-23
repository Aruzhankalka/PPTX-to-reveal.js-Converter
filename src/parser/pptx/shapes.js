'use strict';

/**
 * FR-10 Shape parser — stub-first dispatcher.
 *
 * Reads p:sp and p:cxnSp nodes from an spTree, resolves geometry/fill/stroke
 * to IR Shape objects, and returns them with a warnings array.
 *
 * Design constraints:
 *   - Geometry stored in EMU (integer). The generator converts to CSS.
 *   - rotation stored in native PPTX rot units (1/60000 of a degree). The
 *     generator converts to CSS degrees.
 *   - Theme colors kept as { space:'theme', ref:'accentN' } — NOT baked hex.
 *   - Every preset MUST result in a shape; unsupported presets emit
 *     type:'unknown' + warning so no shape is ever silently dropped.
 *   - The switch lists every mapped type so a collaborating group can add
 *     a new branch without touching existing branches (purely additive).
 *
 * Reuses the xfrm XML access pattern from text.js but reads EMU directly
 * instead of calling emuToPx(), keeping a single geometry path.
 */

const { asArray } = require('./xml');
const { shapeToTextBlock } = require('./text');

// ---------------------------------------------------------------------------
// Preset geometry → IR type mapping
// ---------------------------------------------------------------------------

// Map OOXML prstGeom @prst → IR shape type.
// Each entry is one switch branch — add new cases here, nowhere else.
const PRESET_TO_TYPE = {
  // Rectangles
  rect:          'rect',
  snip1Rect:     'rect',
  snip2SameRect: 'rect',
  snip2DiagRect: 'rect',
  snipRoundRect: 'rect',

  // Rounded rectangles
  roundRect: 'roundRect',

  // Ellipse / circle
  ellipse: 'ellipse',

  // Lines
  line:              'line',
  straightConnector1:'line',

  // Arrows
  rightArrow:       'arrow',
  leftArrow:        'arrow',
  upArrow:          'arrow',
  downArrow:        'arrow',
  leftRightArrow:   'arrow',
  upDownArrow:      'arrow',
  bentArrow:        'arrow',
  uturnArrow:       'arrow',
  circularArrow:    'arrow',
  curvedRightArrow: 'arrow',
  curvedLeftArrow:  'arrow',
  curvedUpArrow:    'arrow',
  curvedDownArrow:  'arrow',
  stripedRightArrow:'arrow',
  notchedRightArrow:'arrow',
  homePlate:        'arrow',
  chevron:          'arrow',
  rightArrowCallout:'arrow',
  leftArrowCallout: 'arrow',
  upArrowCallout:   'arrow',
  downArrowCallout: 'arrow',

  // Polygons
  triangle: 'triangle',
  hexagon:  'hexagon',
  octagon:  'octagon',

  // Arc — open ellipse arc, angles from adj1/adj2 adjustments.
  arc: 'arc',

  // Stars — N-pointed star polygon; star count from preset name, inner ratio from adj.
  star4: 'star', star5: 'star', star6: 'star', star7: 'star', star8: 'star',
  star10: 'star', star12: 'star', star16: 'star', star24: 'star', star32: 'star',

  // Cloud — approximate bezier cloud outline.
  cloud: 'cloud',

  // Flowchart cylinder/disk.
  flowChartMagneticDisk: 'flowchartDisk',

  // Connectors (cxnSp handled separately — see parseCxnShape)
  bentConnector2:   'connector',
  bentConnector3:   'connector',
  bentConnector4:   'connector',
  bentConnector5:   'connector',
  curvedConnector2: 'connector',
  curvedConnector3: 'connector',
  curvedConnector4: 'connector',
  curvedConnector5: 'connector',
  elbow:            'connector',

  // Callouts
  wedgeRectCallout:      'callout',
  wedgeRoundRectCallout: 'callout',
  wedgeEllipseCallout:   'callout',
  cloudCallout:          'callout',
  borderCallout1:        'callout',
  borderCallout2:        'callout',
  borderCallout3:        'callout',
  accentCallout1:        'callout',
  accentCallout2:        'callout',
  accentCallout3:        'callout',
  callout1:              'callout',
  callout2:              'callout',
  callout3:              'callout',
};

// ---------------------------------------------------------------------------
// OOXML scheme-color slot → IR theme ref
// ---------------------------------------------------------------------------

// Maps <a:schemeClr val="..."> to the IR shapeColor.ref enum values.
// Aliases (tx1/bg1 etc.) are resolved here; the IR never contains them.
const SCHEME_TO_REF = {
  accent1:  'accent1',
  accent2:  'accent2',
  accent3:  'accent3',
  accent4:  'accent4',
  accent5:  'accent5',
  accent6:  'accent6',
  dk1:      'text1',
  dk2:      'text2',
  lt1:      'bg1',
  lt2:      'bg2',
  tx1:      'text1',  // display-text alias for dk1
  tx2:      'text2',  // display-text alias for dk2
  bg1:      'bg1',    // background alias for lt1
  bg2:      'bg2',    // background alias for lt2
  hlink:    'link',
  folHlink: 'linkVisited',
};

// ---------------------------------------------------------------------------
// Geometry extraction — reuses the same xfrm XML access pattern as text.js
// but preserves EMU integers instead of converting to px.
// ---------------------------------------------------------------------------

/**
 * Extract position/size/rotation/flip from a <p:spPr> or <p:spPr>-like node.
 *
 * @param {object|null} spPr - parsed <p:spPr> node
 * @returns {{ position: {x,y,w,h}, rotation: number, flipH: boolean, flipV: boolean }}
 */
function extractXfrm(spPr) {
  const xfrm = spPr && spPr['a:xfrm'];
  const off  = (xfrm && xfrm['a:off']) || {};
  const ext  = (xfrm && xfrm['a:ext']) || {};

  const x = Number(off['@_x']) || 0;
  const y = Number(off['@_y']) || 0;
  const w = Number(ext['@_cx']) || 0;
  const h = Number(ext['@_cy']) || 0;

  // rotation: raw PPTX rot units (1/60000 of a degree), 0 if absent.
  // The generator converts to CSS degrees — NOT done here.
  const rotation = (xfrm && xfrm['@_rot'] != null)
    ? (Number(xfrm['@_rot']) || 0)
    : 0;

  const flipH = xfrm ? (xfrm['@_flipH'] === '1' || xfrm['@_flipH'] === true) : false;
  const flipV = xfrm ? (xfrm['@_flipV'] === '1' || xfrm['@_flipV'] === true) : false;

  return { position: { x, y, w, h }, rotation, flipH, flipV };
}

// ---------------------------------------------------------------------------
// Color resolution
// ---------------------------------------------------------------------------

/**
 * Resolve an OOXML fill/stroke color child node to an IR shapeColor object.
 *
 * Supports:
 *   <a:srgbClr val="RRGGBB"/>        → { space:'srgb', hex:'RRGGBB' }
 *   <a:sysClr lastClr="RRGGBB"/>     → { space:'srgb', hex:'RRGGBB' }
 *   <a:schemeClr val="accent1"/>     → { space:'theme', ref:'accent1' }
 *
 * Returns null when the node is absent or unresolvable.
 */
/**
 * Extract an <a:alpha val="..."/> modifier from a parsed color child node.
 * Returns the alpha as an integer 0–100, or null when the modifier is absent.
 * PPTX val units are 1/1000 of a percent (50000 → 50 %).
 */
function extractAlpha(colorNode) {
  const alphaEl = colorNode && colorNode['a:alpha'];
  if (!alphaEl || alphaEl['@_val'] == null) return null;
  const pct = Math.round(Number(alphaEl['@_val']) / 1000);
  return Math.max(0, Math.min(100, pct));
}

function resolveColorNode(node) {
  if (!node) return null;

  const srgb = node['a:srgbClr'];
  if (srgb && srgb['@_val']) {
    const color = { space: 'srgb', hex: String(srgb['@_val']).toUpperCase() };
    const alpha = extractAlpha(srgb);
    if (alpha != null && alpha < 100) color.alpha = alpha;
    return color;
  }

  const sys = node['a:sysClr'];
  if (sys && sys['@_lastClr']) {
    const color = { space: 'srgb', hex: String(sys['@_lastClr']).toUpperCase() };
    const alpha = extractAlpha(sys);
    if (alpha != null && alpha < 100) color.alpha = alpha;
    return color;
  }

  const scheme = node['a:schemeClr'];
  if (scheme && scheme['@_val']) {
    const raw = String(scheme['@_val']);
    const ref = SCHEME_TO_REF[raw];
    const color = { space: 'theme', ref: ref || 'text1' };
    const alpha = extractAlpha(scheme);
    if (alpha != null && alpha < 100) color.alpha = alpha;
    return color;
  }

  return null;
}

// ---------------------------------------------------------------------------
// Fill resolution
// ---------------------------------------------------------------------------

/**
 * Resolve <a:gradFill> to an IR gradient fill, or null + warning if < 2 stops.
 *
 * Reads <a:gsLst>/<a:gs pos="..."> children; position is 0–100000.
 * Determines kind from <a:lin> (linear) or <a:path> (radial).
 *
 * @param {object}   gradFill - parsed <a:gradFill> node
 * @param {string[]} warnings - mutable array; push when falling back to none
 * @returns {{ type:'gradient', kind, angle?, stops } | null}
 */
function resolveGradientFill(gradFill, warnings) {
  const gsLst = gradFill && gradFill['a:gsLst'];
  if (!gsLst) {
    warnings.push('gradFill: missing gsLst, falling back to none');
    return null;
  }

  const stops = [];
  for (const gs of asArray(gsLst['a:gs'])) {
    const pos   = gs['@_pos'] != null ? Number(gs['@_pos']) : null;
    if (pos === null) continue;
    const color = resolveColorNode(gs);
    if (!color) continue;
    stops.push({ pos, color });
  }

  if (stops.length < 2) {
    warnings.push(`gradFill: only ${stops.length} resolvable stop(s), falling back to none`);
    return null;
  }

  stops.sort((a, b) => a.pos - b.pos);

  const lin = gradFill['a:lin'];
  if (lin) {
    const angle = lin['@_ang'] != null ? Number(lin['@_ang']) : 0;
    return { type: 'gradient', kind: 'linear', angle, stops };
  }

  const path = gradFill['a:path'];
  if (path) {
    return { type: 'gradient', kind: 'radial', stops };
  }

  // No direction element — default to linear with angle 0
  return { type: 'gradient', kind: 'linear', angle: 0, stops };
}

/**
 * Resolve <p:spPr> fill sub-nodes to an IR shapeFill object.
 *
 * Handled:
 *   <a:noFill/>       → { type:'none' }
 *   <a:solidFill>...  → { type:'solid', color: Color }
 *   <a:gradFill>...   → { type:'gradient', kind, angle?, stops } or none if < 2 stops
 *   pattern/group fill → { type:'none' }
 */
function resolveFill(spPr, warnings = []) {
  if (!spPr) return { type: 'none' };

  if (spPr['a:noFill']) return { type: 'none' };

  const solidFill = spPr['a:solidFill'];
  if (solidFill) {
    const color = resolveColorNode(solidFill);
    if (color) return { type: 'solid', color };
    return { type: 'none' };
  }

  const gradFill = spPr['a:gradFill'];
  if (gradFill) {
    return resolveGradientFill(gradFill, warnings) || { type: 'none' };
  }

  return { type: 'none' };
}

// ---------------------------------------------------------------------------
// Stroke resolution
// ---------------------------------------------------------------------------

/**
 * Resolve an <a:headEnd> or <a:tailEnd> node to an IR arrowEnd object.
 * Returns undefined when the arrow type is absent or 'none'.
 *
 * @param {object|undefined} endNode - parsed end-marker node
 * @returns {{ type, width?, length? } | undefined}
 */
function resolveArrowEnd(endNode) {
  if (!endNode) return undefined;
  const type = endNode['@_type'] || 'none';
  if (type === 'none') return undefined;
  const result = { type };
  const w   = endNode['@_w'];
  const len = endNode['@_len'];
  if (w)   result.width  = w;
  if (len) result.length = len;
  return result;
}

/**
 * Resolve <p:spPr><a:ln> to an IR shapeStroke object.
 *
 * Default PPTX line width is 12700 EMU (1 point).
 * Captures optional headEnd / tailEnd arrowhead markers.
 */
function resolveStroke(spPr) {
  if (!spPr) return { type: 'none' };

  const ln = spPr['a:ln'];
  if (!ln) return { type: 'none' };
  if (ln['a:noFill']) return { type: 'none' };

  const solidFill = ln['a:solidFill'];
  if (solidFill) {
    const color = resolveColorNode(solidFill);
    if (color) {
      const widthEmu = ln['@_w'] != null ? (Number(ln['@_w']) || 12700) : 12700;
      const stroke = { type: 'solid', color, widthEmu };
      const headEnd = resolveArrowEnd(ln['a:headEnd']);
      const tailEnd = resolveArrowEnd(ln['a:tailEnd']);
      if (headEnd) stroke.headEnd = headEnd;
      if (tailEnd) stroke.tailEnd = tailEnd;
      return stroke;
    }
  }

  return { type: 'none' };
}

/**
 * Resolve <p:spPr><a:effectLst> to an IR effects object, or undefined if absent.
 *
 * Captures <a:outerShdw> and <a:innerShdw> (outer takes priority).
 * Alpha is extracted from the color child's <a:alpha val> (1/1000 percent units).
 *
 * @param {object|undefined} spPr - parsed <p:spPr> node
 * @returns {{ shadow: {...} } | undefined}
 */
function resolveEffects(spPr) {
  if (!spPr) return undefined;
  const effectLst = spPr['a:effectLst'];
  if (!effectLst) return undefined;

  const effects = {};

  for (const [mode, tag] of [['outer', 'a:outerShdw'], ['inner', 'a:innerShdw']]) {
    const shdwNode = effectLst[tag];
    if (!shdwNode) continue;

    const blurEmu       = shdwNode['@_blurRad'] != null ? Number(shdwNode['@_blurRad']) : 0;
    const distanceEmu   = shdwNode['@_dist']    != null ? Number(shdwNode['@_dist'])    : 0;
    const directionAngle = shdwNode['@_dir']    != null ? Number(shdwNode['@_dir'])     : 0;

    const color = resolveColorNode(shdwNode);
    if (!color) continue;

    // Alpha lives as a child modifier of the color element, e.g.
    // <a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr>
    // PPTX alpha val is in 1/1000 percent units: 50000 → 50 %
    let alphaPct = 100;
    for (const colorTag of ['a:srgbClr', 'a:schemeClr', 'a:sysClr']) {
      const colorChild = shdwNode[colorTag];
      if (!colorChild) continue;
      const alphaNode = colorChild['a:alpha'];
      if (alphaNode && alphaNode['@_val'] != null) {
        alphaPct = Math.max(0, Math.min(100, Math.round(Number(alphaNode['@_val']) / 1000)));
      }
      break;
    }

    effects.shadow = { mode, color, blurEmu, distanceEmu, directionAngle, alphaPct };
    break; // outer shadow takes priority; stop after first hit
  }

  return Object.keys(effects).length > 0 ? effects : undefined;
}

// ---------------------------------------------------------------------------
// p:style inheritance — fill/stroke/font from the shape's theme style reference
// ---------------------------------------------------------------------------

/**
 * Extract the CSS color string from <p:style><a:fontRef> for use as the
 * fallback text color of shape text runs.
 *
 * <a:fontRef idx="minor|major|none"> carries the shape's default font slot
 * and a color child that defines the text color inherited by all runs that
 * carry no explicit <a:solidFill> in their <a:rPr>.
 *
 * The returned string uses the same format as run.formatting.color in text.js
 * so the SVG foreignObject renderer can apply it directly.
 *
 * @param {object|undefined} pStyle - parsed <p:style> node
 * @returns {string|null} CSS color string or null when absent/unresolvable
 */
const TEXT_SCHEME_ALIAS = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };

function fontRefColor(pStyle) {
  if (!pStyle) return null;
  const fontRef = pStyle['a:fontRef'];
  if (!fontRef) return null;
  // idx="none" means no font reference; major/minor are valid slot names.
  if (fontRef['@_idx'] === 'none') return null;

  const srgb = fontRef['a:srgbClr'];
  if (srgb && srgb['@_val']) return '#' + String(srgb['@_val']).toUpperCase();

  const sys = fontRef['a:sysClr'];
  if (sys && sys['@_lastClr']) return '#' + String(sys['@_lastClr']).toUpperCase();

  const scheme = fontRef['a:schemeClr'];
  if (scheme && scheme['@_val']) {
    const raw = String(scheme['@_val']);
    return `var(--theme-${TEXT_SCHEME_ALIAS[raw] || raw})`;
  }

  return null;
}

/**
 * Return true when <p:spPr> carries any explicit fill node, meaning the fill
 * is self-contained and no style inheritance is needed.
 */
function hasExplicitFillNode(spPr) {
  return !!(
    spPr['a:noFill'] ||
    spPr['a:solidFill'] ||
    spPr['a:gradFill'] ||
    spPr['a:pattFill'] ||
    spPr['a:blipFill'] ||
    spPr['a:grpFill']
  );
}

/**
 * Return true when <p:spPr> carries an explicit <a:ln> node, meaning the
 * stroke is self-contained and no style inheritance is needed.
 */
function hasExplicitStrokeNode(spPr) {
  return !!spPr['a:ln'];
}

/**
 * Resolve the fill from a shape's <p:style><a:fillRef> node.
 *
 * OOXML §20.1.4.2.10: fillRef idx=0 means no fill; idx≥1 references the
 * theme's fill effects list.  The color child of <a:fillRef> is the actual
 * tint applied to that fill slot and is what we use directly.
 *
 * Returns null when the style carries no fill information.
 */
function resolveStyleFill(pStyle) {
  if (!pStyle) return null;
  const fillRef = pStyle['a:fillRef'];
  if (!fillRef) return null;

  const idx = Number(fillRef['@_idx']);
  if (idx === 0) return { type: 'none' }; // explicit "no fill" in style

  const color = resolveColorNode(fillRef);
  if (!color) return null;
  return { type: 'solid', color };
}

/**
 * Resolve the stroke from a shape's <p:style><a:lnRef> node.
 *
 * OOXML §20.1.4.2.19: lnRef idx=0 means no line; idx≥1 references the
 * theme's line effects list.  Width comes from the theme list (not directly
 * available here), so we default to 12700 EMU (1 pt) — sufficient for
 * visibility; a future pass can refine it using the theme's lnStyleLst.
 *
 * Returns null when the style carries no line information.
 */
function resolveStyleStroke(pStyle) {
  if (!pStyle) return null;
  const lnRef = pStyle['a:lnRef'];
  if (!lnRef) return null;

  const idx = Number(lnRef['@_idx']);
  if (idx === 0) return { type: 'none' }; // explicit "no line" in style

  const color = resolveColorNode(lnRef);
  if (!color) return null;
  return { type: 'solid', color, widthEmu: 12700 };
}

// ---------------------------------------------------------------------------
// Adjustments (roundRect corner radius, etc.)
// ---------------------------------------------------------------------------

// Tags that carry path commands in document order (fast-xml-parser breaks ordering
// across different tag names, so each type is collected separately).
const CUSTOM_GEO_OP_TAGS = [
  ['a:moveTo',     'moveTo',     1],
  ['a:lnTo',       'lnTo',       1],
  ['a:cubicBezTo', 'cubicBezTo', 3],
  ['a:quadBezTo',  'quadBezTo',  2],
  ['a:arcTo',      'arcTo',      0],
  ['a:close',      'close',      0],
];

/**
 * Extract shape adjustments from <a:prstGeom><a:avLst>.
 * Returns an array of { name, value } or undefined when none are present.
 * Non-val formulas (e.g. expressions) are skipped with a warning.
 */
function extractAdjustments(prstGeom, warnings = []) {
  if (!prstGeom) return undefined;
  const avLst = prstGeom['a:avLst'];
  if (!avLst) return undefined;

  const result = [];
  for (const gd of asArray(avLst['a:gd'])) {
    const name  = gd['@_name'];
    const fmla  = gd['@_fmla'];
    if (!name || !fmla) continue;
    const match = String(fmla).match(/^val\s+(-?\d+)$/);
    if (match) {
      result.push({ name, value: Number(match[1]) });
    } else {
      warnings.push(`adjustment "${name}" has non-val formula "${fmla}", skipped`);
    }
  }

  return result.length > 0 ? result : undefined;
}

/**
 * Extract <a:custGeom> path data from <p:spPr> into an IR customGeometry object.
 *
 * Each <a:path w h> element's commands are collected per tag type.
 * Note: fast-xml-parser does not preserve order across different tag names, so
 * commands of different types (e.g. interleaved moveTo+lnTo) lose their relative
 * ordering. Commands of the same type retain their original sequence.
 *
 * Returns undefined when no custGeom or no paths are present.
 */
function extractCustomGeometry(spPr) {
  if (!spPr) return undefined;
  const custGeom = spPr['a:custGeom'];
  if (!custGeom) return undefined;
  const pathLst = custGeom['a:pathLst'];
  if (!pathLst) return undefined;

  let topW = 0;
  let topH = 0;
  const paths = [];

  for (const pathEl of asArray(pathLst['a:path'])) {
    const pw = pathEl['@_w'] != null ? Number(pathEl['@_w']) : 0;
    const ph = pathEl['@_h'] != null ? Number(pathEl['@_h']) : 0;
    if (!topW) topW = pw;
    if (!topH) topH = ph;

    const commands = [];
    for (const [xmlTag, op, ptCount] of CUSTOM_GEO_OP_TAGS) {
      for (const cmdNode of asArray(pathEl[xmlTag])) {
        const cmd = { op };
        if (ptCount > 0) {
          const ptList = asArray(cmdNode['a:pt']);
          cmd.pts = ptList.slice(0, ptCount).map((pt) => ({
            x: Number(pt['@_x']) || 0,
            y: Number(pt['@_y']) || 0,
          }));
        }
        commands.push(cmd);
      }
    }
    if (commands.length > 0) paths.push({ commands });
  }

  if (paths.length === 0) return undefined;
  return { w: topW, h: topH, paths };
}

// ---------------------------------------------------------------------------
// Embedded text extraction
// ---------------------------------------------------------------------------

/**
 * Extract embedded text from a shape's <p:txBody> as an IR TextBlock.
 * Strips position/dimensions (they live on the enclosing shape) and returns
 * id + paragraphs + bodyPr metadata needed by the SVG renderer.
 *
 * anchor:  vertical alignment ('t' | 'ctr' | 'b') — default 't'
 * insets:  body padding in EMU; PPTX defaults are lIns=rIns=91440, tIns=bIns=45720
 */
function extractEmbeddedText(pSp, idx, txStyles) {
  // Use the master txStyles so sizeFromTxStyles picks p:otherStyle (18pt) for
  // non-placeholder drawing-object text instead of the body-text default (22pt).
  const block = shapeToTextBlock(pSp, idx, txStyles);
  if (!block) return null;

  const bodyPr = pSp['p:txBody'] && pSp['p:txBody']['a:bodyPr'];
  const anchor = (bodyPr && bodyPr['@_anchor']) || 't';

  // PPTX default insets (EMU): left=right=91440, top=bottom=45720
  const lIns = bodyPr && bodyPr['@_lIns'] != null ? Number(bodyPr['@_lIns']) : 91440;
  const rIns = bodyPr && bodyPr['@_rIns'] != null ? Number(bodyPr['@_rIns']) : 91440;
  const tIns = bodyPr && bodyPr['@_tIns'] != null ? Number(bodyPr['@_tIns']) : 45720;
  const bIns = bodyPr && bodyPr['@_bIns'] != null ? Number(bodyPr['@_bIns']) : 45720;

  const result = {
    id:         block.id,
    paragraphs: block.paragraphs,
    anchor,
    insets: { l: lIns, r: rIns, t: tIns, b: bIns },
  };
  if (block.autoFit != null) result.autoFit = block.autoFit;
  return result;
}

// ---------------------------------------------------------------------------
// Shape type resolution (stub-first dispatcher)
// ---------------------------------------------------------------------------

/**
 * Map a prstGeom @prst attribute to an IR type string.
 * Returns null when the preset is absent (custom geometry), which the
 * caller converts to 'unknown'.
 */
function mapPreset(prst) {
  if (!prst) return null;
  return PRESET_TO_TYPE[prst] || null;
}

// ---------------------------------------------------------------------------
// p:sp parser
// ---------------------------------------------------------------------------

/**
 * Parse a single <p:sp> node that is NOT a text placeholder.
 * Text placeholders (those with <p:nvSpPr><p:nvPr><p:ph>) are handled by
 * text.js and must not appear in the shapes output.
 *
 * @param {object}   pSp      - parsed <p:sp> node
 * @param {number}   idx      - 0-based counter for generating stable ids
 * @param {object}   txStyles - master txStyles (may be null)
 * @param {string[]} warnings - mutable array; push human-readable strings
 * @returns {object|null} IR Shape, or null when the node has a placeholder
 */
function parseSp(pSp, idx, txStyles, warnings) {
  const nvSpPr = pSp['p:nvSpPr'];
  const ph = nvSpPr && nvSpPr['p:nvPr'] && nvSpPr['p:nvPr']['p:ph'];
  if (ph) return null; // text placeholder — handled by text.js

  const spPr    = pSp['p:spPr'] || {};
  const prstGeom = spPr['a:prstGeom'];
  const prst     = prstGeom ? prstGeom['@_prst'] : null;

  const { position, rotation, flipH, flipV } = extractXfrm(spPr);
  const pStyle = pSp['p:style'];
  const fill   = hasExplicitFillNode(spPr)
    ? resolveFill(spPr, warnings)
    : (resolveStyleFill(pStyle) || { type: 'none' });
  const stroke = hasExplicitStrokeNode(spPr)
    ? resolveStroke(spPr)
    : (resolveStyleStroke(pStyle) || { type: 'none' });

  const irType = mapPreset(prst);
  let type;
  if (irType) {
    type = irType;
  } else {
    // Preserve the original PPTX preset name so the generator can attempt
    // approximate rendering (e.g. "hexagon", "star7"). Custom geometry with
    // no prst attribute falls back to the sentinel "unknown".
    type = prst || 'unknown';
    warnings.push(`shape preset "${type}" not yet supported by generator`);
  }

  const shape = {
    id:       `shp-${idx}`,
    type,
    position,
    rotation,
    flipH,
    flipV,
    fill,
    stroke,
    z: 0, // overridden by slide.js via getSpTreeChildOrder
  };

  // Preserve the original PPTX preset name so the generator can do
  // direction-aware rendering (e.g. distinguish rightArrow from leftArrow).
  if (prst) shape.preset = prst;

  if (!irType) shape.supported = false;

  const adj = extractAdjustments(prstGeom, warnings);
  if (adj) shape.adjustments = adj;

  const effects = resolveEffects(spPr);
  if (effects) shape.effects = effects;

  const custGeo = extractCustomGeometry(spPr);
  if (custGeo) shape.customGeometry = custGeo;

  // Connectors and polyline/polygon need a vertex list; for p:sp shapes these
  // come from custGeom (not yet parsed — emit empty points so the schema
  // constraint is satisfied and the generator can skip gracefully).
  if (type === 'connector' || type === 'polyline' || type === 'polygon') {
    shape.points = [];
  }

  const text = extractEmbeddedText(pSp, idx, txStyles);
  if (text) {
    // Apply p:style fontRef color as fallback for runs that carry no explicit color.
    // This is the shape's inherited text color (e.g. white text on a colored rect).
    const frc = fontRefColor(pStyle);
    if (frc) {
      for (const para of text.paragraphs || []) {
        for (const run of para.runs || []) {
          if (run.text && (!run.formatting || !run.formatting.color)) {
            if (!run.formatting) run.formatting = {};
            run.formatting.color = frc;
          }
        }
      }
    }
    shape.text = text;
  }

  return shape;
}

// ---------------------------------------------------------------------------
// p:cxnSp parser (connection shapes)
// ---------------------------------------------------------------------------

/**
 * Parse a <p:cxnSp> (connection shape) node into an IR Shape with
 * type:'connector'.
 */
function parseCxnSp(pCxnSp, idx, warnings) {
  const spPr = pCxnSp['p:spPr'] || {};

  const { position, rotation, flipH, flipV } = extractXfrm(spPr);
  const pStyle = pCxnSp['p:style'];
  const fill   = hasExplicitFillNode(spPr)
    ? resolveFill(spPr, warnings)
    : (resolveStyleFill(pStyle) || { type: 'none' });
  const stroke = hasExplicitStrokeNode(spPr)
    ? resolveStroke(spPr)
    : (resolveStyleStroke(pStyle) || { type: 'none' });

  const prstGeom = spPr['a:prstGeom'];
  const prst     = prstGeom ? prstGeom['@_prst'] : null;

  // cxnSp nodes are always connectors in the IR regardless of prst.
  // If prst is not in the connector list we still emit connector, not unknown,
  // because the element type (p:cxnSp) is authoritative.
  const irType = (prst && PRESET_TO_TYPE[prst] === 'connector')
    ? 'connector'
    : 'connector'; // always connector

  if (prst && !PRESET_TO_TYPE[prst]) {
    warnings.push(`connector preset "${prst}" not yet supported`);
  }

  const shape = {
    id:       `cxn-${idx}`,
    type:     irType,
    position,
    rotation,
    flipH,
    flipV,
    fill,
    stroke,
    points:   [], // endpoints from spTree not yet resolved
    z: 0,
  };

  if (prst) shape.preset = prst;

  const adj = extractAdjustments(prstGeom, warnings);
  if (adj) shape.adjustments = adj;

  const effects = resolveEffects(spPr);
  if (effects) shape.effects = effects;

  return shape;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse all shape elements from an spTree into IR Shape objects.
 *
 * Processes p:sp (non-placeholder) and p:cxnSp children at all levels of
 * nesting, including inside p:grpSp group containers. Shapes that are text
 * placeholders (have <p:ph>) are skipped — text.js handles them separately.
 *
 * @param {object}   spTree   - parsed <p:spTree> node
 * @param {object}   txStyles - master txStyles passed through to text extraction
 * @param {string[]} warnings - mutable array; warnings are appended here
 * @returns {object[]} array of IR Shape objects (never null, may be empty)
 */
function parseShapes(spTree, txStyles, warnings) {
  if (!spTree) return [];

  const shapes = [];
  let spIdx  = 0;
  let cxnIdx = 0;

  // Recurse into p:grpSp so shapes inside groups are not silently dropped.
  // spIdx and cxnIdx are shared across all levels so IDs remain unique.
  function walkTree(tree) {
    for (const pSp of asArray(tree['p:sp'])) {
      const shape = parseSp(pSp, spIdx, txStyles, warnings);
      if (shape) {
        shapes.push(shape);
        spIdx++;
      }
    }
    for (const pCxnSp of asArray(tree['p:cxnSp'])) {
      shapes.push(parseCxnSp(pCxnSp, cxnIdx++, warnings));
    }
    for (const pGrpSp of asArray(tree['p:grpSp'])) {
      walkTree(pGrpSp);
    }
  }

  walkTree(spTree);
  return shapes;
}

/**
 * Collect layout/master placeholder shapes that carry explicit fills as
 * background shape objects.
 *
 * In OOXML, a layout's <p:sp> with <p:ph> normally defines a template slot
 * for slide content (text, media). parseSp() correctly skips those so they
 * don't appear as duplicate graphical shapes on the slide. However, many
 * templates apply explicit <a:solidFill> (or gradient fill) to those same
 * placeholder nodes to create colored background areas. Without this function
 * those colors are silently dropped and shapes appear invisible.
 *
 * This function is ONLY meant to be called on layout/master spTrees (not
 * slide spTrees). It returns bare rect shapes — text is intentionally absent
 * because slide.js already deletes text from inherited shapes.
 *
 * @param {object}   spTree   - parsed <p:spTree> from a layout or master XML
 * @param {string[]} warnings - mutable array; warnings are appended here
 * @returns {object[]} IR Shape objects (type:'rect', no text field)
 */
function parsePlaceholderBackgrounds(spTree, warnings) {
  if (!spTree) return [];

  const shapes = [];
  let idx = 0;

  for (const pSp of asArray(spTree['p:sp'])) {
    const ph = pSp['p:nvSpPr']
      && pSp['p:nvSpPr']['p:nvPr']
      && pSp['p:nvSpPr']['p:nvPr']['p:ph'];
    if (!ph) continue; // non-placeholder shapes handled by parseShapes

    const spPr   = pSp['p:spPr'] || {};
    const pStyle = pSp['p:style'];

    const fill = hasExplicitFillNode(spPr)
      ? resolveFill(spPr, warnings)
      : (resolveStyleFill(pStyle) || { type: 'none' });

    if (fill.type === 'none') continue; // no visible fill → nothing to contribute

    const { position, rotation, flipH, flipV } = extractXfrm(spPr);
    const stroke = hasExplicitStrokeNode(spPr)
      ? resolveStroke(spPr)
      : (resolveStyleStroke(pStyle) || { type: 'none' });

    shapes.push({
      id:       `ph-bg-${idx++}`,
      type:     'rect',
      position,
      rotation,
      flipH,
      flipV,
      fill,
      stroke,
      z: 0, // overridden by slide.js
    });
  }

  return shapes;
}

module.exports = {
  parseShapes,
  parsePlaceholderBackgrounds,
  extractXfrm,
  resolveColorNode,
  resolveFill,
  resolveGradientFill,
  resolveStroke,
  resolveArrowEnd,
  resolveEffects,
  extractAdjustments,
  extractCustomGeometry,
};
