'use strict';

const { emuToPx, pptxRotationToDegrees } = require('./units');
const { SCHEME_ALIAS } = require('./color');

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
  homePlate:        'pentagon',
  rightArrowCallout:'arrow',
  leftArrowCallout: 'arrow',
  upArrowCallout:   'arrow',
  downArrowCallout: 'arrow',

  // Polygons
  triangle:   'triangle',
  rtTriangle: 'rtTriangle',
  hexagon:    'hexagon',
  octagon:    'octagon',
  pentagon:   'pentagon',
  chevron:    'chevron',

  // Arc — open ellipse arc, angles from adj1/adj2 adjustments.
  arc: 'arc',

  // Stars — N-pointed star polygon; star count from preset name, inner ratio from adj.
  star4: 'star', star5: 'star', star6: 'star', star7: 'star', star8: 'star',
  star10: 'star', star12: 'star', star16: 'star', star24: 'star', star32: 'star',

  // Cloud — approximate bezier cloud outline.
  cloud: 'cloud',

  // Flowchart cylinder/disk.
  flowChartMagneticDisk: 'database',

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

// Raw <a:schemeClr val="..."> aliases that don't match a theme.colors slot key
// directly (theme.js only stores the canonical dk1/lt1/dk2/lt2 slots).
const RAW_SCHEME_TO_THEME_SLOT = {
  tx1: 'dk1',
  tx2: 'dk2',
  bg1: 'lt1',
  bg2: 'lt2',
};

// ---------------------------------------------------------------------------
// Geometry extraction — reuses the same xfrm XML access pattern as text.js
// but preserves EMU integers instead of converting to px.
// ---------------------------------------------------------------------------

/**
 * Extract position/size/rotation/flip from a <p:spPr> or <p:spPr>-like node.
 *
 * Named distinctly from layouts.js's extractXfrm: that one takes a full <p:sp>
 * node and returns px/degrees; this one takes <p:spPr> directly and returns
 * raw EMU/raw PPTX rotation units (converted later by the generator). Same
 * name, incompatible contracts — kept separate to avoid a silent unit-mismatch
 * bug if the wrong one is ever imported into the wrong file.
 *
 * @param {object|null} spPr - parsed <p:spPr> node
 * @returns {{ position: {x,y,w,h}, rotation: number, flipH: boolean, flipV: boolean }}
 */
function extractShapeXfrm(spPr) {
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
// Group transform composition (FR-10 groups[])
//
// In OOXML, a <p:grpSp>'s children are NOT positioned in slide coordinates.
// Their <a:xfrm> off/ext live in the group's own "child coordinate space",
// defined by the group's chOff/chExt — which the group's own off/ext (its
// box on the slide, or on its parent group) is scaled+translated onto. If a
// group is resized or moved as a whole after its children were authored,
// chOff/chExt and off/ext diverge, and reading a child's raw off/ext (the
// old behavior here) places it in the wrong spot. Composing a transform per
// nesting level — scale+translate from chOff/chExt into off/ext, then rotate
// the group's box (with its now-correctly-placed children) as one rigid body
// around its own center — and chaining that through nested groups keeps
// every descendant's reported position correct in absolute slide EMU.
//
// Known scope limit: group-level flipH/flipV is not composed into children
// (real OOXML mirrors child position+orientation too). Groups that are both
// non-uniformly scaled AND rotated can shear in true PPTX rendering, which a
// simple rotated bounding box can't reproduce. Both are rare in practice and
// left as a documented approximation rather than blocking this feature.
// ---------------------------------------------------------------------------

const IDENTITY_TRANSFORM = {
  scaleX: 1, scaleY: 1, rotationUnits: 0,
  mapPoint: (x, y) => ({ x, y }),
};

/**
 * Rotate point (x,y) around (cx,cy) by rotUnits (PPTX 1/60000-degree units).
 * Screen/PPTX coordinates are y-down with clockwise-positive rotation, which
 * the standard (un-negated) rotation matrix already matches.
 */
function rotatePoint(x, y, cx, cy, rotUnits) {
  if (!rotUnits) return { x, y };
  const rad = (rotUnits / 60000) * (Math.PI / 180);
  const dx = x - cx;
  const dy = y - cy;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  return { x: cx + dx * cos - dy * sin, y: cy + dx * sin + dy * cos };
}

/**
 * Extract a group's own off/ext/chOff/chExt/rot from <p:grpSpPr><a:xfrm>.
 * Returns null when the group carries no xfrm (rare/malformed) so the caller
 * can fall back to flattening its children without a coordinate frame.
 */
function extractGroupXfrm(pGrpSp) {
  const xfrm = pGrpSp['p:grpSpPr'] && pGrpSp['p:grpSpPr']['a:xfrm'];
  if (!xfrm) return null;
  const off   = xfrm['a:off']   || {};
  const ext   = xfrm['a:ext']   || {};
  const chOff = xfrm['a:chOff'] || {};
  const chExt = xfrm['a:chExt'] || {};
  return {
    offX: Number(off['@_x']) || 0,
    offY: Number(off['@_y']) || 0,
    extW: Number(ext['@_cx']) || 0,
    extH: Number(ext['@_cy']) || 0,
    chOffX: Number(chOff['@_x']) || 0,
    chOffY: Number(chOff['@_y']) || 0,
    chExtW: Number(chExt['@_cx']) || 0,
    chExtH: Number(chExt['@_cy']) || 0,
    rotUnits: xfrm['@_rot'] != null ? (Number(xfrm['@_rot']) || 0) : 0,
  };
}

/**
 * Compose a parent transform with one group level's own xfrm into a new
 * transform mapping points from THIS group's child coordinate space all the
 * way to absolute slide EMU.
 */
function composeGroupTransform(parent, g) {
  // chExt missing or zero (e.g. malformed/test fixtures) → treat as 1:1,
  // i.e. the child coordinate space equals the group's own box unscaled.
  const safeChExtW = g.chExtW || g.extW || 1;
  const safeChExtH = g.chExtH || g.extH || 1;
  const localScaleX = g.extW / safeChExtW;
  const localScaleY = g.extH / safeChExtH;
  const centerX = g.offX + g.extW / 2;
  const centerY = g.offY + g.extH / 2;

  function mapPoint(x, y) {
    const localX = g.offX + (x - g.chOffX) * localScaleX;
    const localY = g.offY + (y - g.chOffY) * localScaleY;
    const rotated = rotatePoint(localX, localY, centerX, centerY, g.rotUnits);
    return parent.mapPoint(rotated.x, rotated.y);
  }

  return {
    scaleX: parent.scaleX * localScaleX,
    scaleY: parent.scaleY * localScaleY,
    rotationUnits: parent.rotationUnits + g.rotUnits,
    mapPoint,
  };
}

/**
 * Map a local (x,y,w,h,rotation) box through an accumulated group transform
 * into absolute slide-space EMU. Identity transform short-circuits to the
 * exact input (no rounding) so non-grouped shapes are byte-for-byte unchanged.
 */
function mapBoxThroughTransform(x, y, w, h, rotUnits, transform) {
  if (transform === IDENTITY_TRANSFORM) {
    return { position: { x, y, w, h }, rotation: rotUnits };
  }
  const cx = x + w / 2;
  const cy = y + h / 2;
  const abs = transform.mapPoint(cx, cy);
  const absW = w * transform.scaleX;
  const absH = h * transform.scaleY;
  return {
    position: {
      x: Math.round(abs.x - absW / 2),
      y: Math.round(abs.y - absH / 2),
      w: Math.round(absW),
      h: Math.round(absH),
    },
    rotation: Math.round(rotUnits + transform.rotationUnits),
  };
}

/**
 * Apply an accumulated transform to a shape's already-extracted local
 * geometry (extractShapeXfrm's return shape).
 */
function applyGroupTransform(geo, transform) {
  if (transform === IDENTITY_TRANSFORM) return geo;
  const { position, rotation } = mapBoxThroughTransform(
    geo.position.x, geo.position.y, geo.position.w, geo.position.h,
    geo.rotation, transform,
  );
  return { position, rotation, flipH: geo.flipH, flipV: geo.flipV };
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

/**
 * @param {object} node - parsed color-bearing node (e.g. <a:solidFill>)
 * @param {Record<string,string>|null} [themeColors] - theme.colors dict (slot → #RRGGBB),
 *   used to bake tint/shade/lumMod/lumOff modifiers on a direct <a:schemeClr> into a
 *   concrete hex. Without it (or without modifiers), theme colors stay structured
 *   { space:'theme', ref } so the generator emits var(--theme-X).
 */
function resolveColorNode(node, themeColors) {
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

    const hasLightnessMods = scheme['a:tint'] || scheme['a:shade'] || scheme['a:lumMod'] || scheme['a:lumOff'];
    if (hasLightnessMods && themeColors) {
      const slot = RAW_SCHEME_TO_THEME_SLOT[raw] || raw;
      const baseHex = themeColors[slot] && String(themeColors[slot]).replace('#', '');
      if (baseHex && baseHex.length === 6) {
        const color = { space: 'srgb', hex: applyColorModifiers(baseHex, scheme).toUpperCase() };
        const alpha = extractAlpha(scheme);
        if (alpha != null && alpha < 100) color.alpha = alpha;
        return color;
      }
    }

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
 * @param {Record<string,string>|null} [themeColors] - see resolveColorNode
 * @returns {{ type:'gradient', kind, angle?, stops } | null}
 */
function resolveGradientFill(gradFill, warnings, themeColors) {
  const gsLst = gradFill && gradFill['a:gsLst'];
  if (!gsLst) {
    warnings.push('gradFill: missing gsLst, falling back to none');
    return null;
  }

  const stops = [];
  for (const gs of asArray(gsLst['a:gs'])) {
    const pos   = gs['@_pos'] != null ? Number(gs['@_pos']) : null;
    if (pos === null) continue;
    const color = resolveColorNode(gs, themeColors);
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
 *
 * @param {object|undefined} spPr
 * @param {string[]} [warnings]
 * @param {Record<string,string>|null} [themeColors] - see resolveColorNode
 */
function resolveFill(spPr, warnings = [], themeColors) {
  if (!spPr) return { type: 'none' };

  if (spPr['a:noFill']) return { type: 'none' };

  const solidFill = spPr['a:solidFill'];
  if (solidFill) {
    const c = resolveColorNode(solidFill, themeColors);
    const color = shapeColorToFlatCss(c);
    if (color) {
      const alpha = shapeColorAlpha(c);
      return alpha != null ? { type: 'solid', color, alpha } : { type: 'solid', color };
    }
    return { type: 'none' };
  }

  const gradFill = spPr['a:gradFill'];
  if (gradFill) {
    // Gradient stops keep structured shapeColor for alpha precision in the
    // SVG renderer; the top-level fill.color field is absent for gradients
    // (no single representative color exists).
    return resolveGradientFill(gradFill, warnings, themeColors) || { type: 'none' };
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
 * Resolve <p:spPr><a:ln> to an IR stroke object.
 *
 * spec fields: type (none|solid), color (flat CSS string), width (px),
 *   style (solid|dashed|pointed), plus headEnd/tailEnd arrowheads (extension).
 * Default PPTX line width is 12700 EMU ≈ 1.33 px.
 *
 * @param {object|undefined} spPr
 * @param {Record<string,string>|null} [themeColors] - see resolveColorNode
 */
function resolveStroke(spPr, themeColors) {
  if (!spPr) return { type: 'none' };

  const ln = spPr['a:ln'];
  if (!ln) return { type: 'none' };
  if (ln['a:noFill']) return { type: 'none' };

  const solidFill = ln['a:solidFill'];
  if (solidFill) {
    const color = shapeColorToFlatCss(resolveColorNode(solidFill, themeColors));
    if (color) {
      const widthEmu = ln['@_w'] != null ? (Number(ln['@_w']) || 12700) : 12700;
      const width = parseFloat((widthEmu / 9525).toFixed(2));
      const style = resolveStrokeDash(ln);
      const stroke = { type: 'solid', color, width, style };
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
 * Extract an effect node's own <a:alpha val> from its color child (1/1000 percent
 * units, e.g. <a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr> → 50).
 * Defaults to 100 when the color child carries no alpha modifier.
 */
function extractEffectAlpha(effectNode) {
  for (const colorTag of ['a:srgbClr', 'a:schemeClr', 'a:sysClr']) {
    const colorChild = effectNode[colorTag];
    if (!colorChild) continue;
    const alphaNode = colorChild['a:alpha'];
    if (alphaNode && alphaNode['@_val'] != null) {
      return Math.max(0, Math.min(100, Math.round(Number(alphaNode['@_val']) / 1000)));
    }
    break;
  }
  return 100;
}

/**
 * Resolve <p:spPr><a:effectLst> to an IR effects object, or undefined if absent.
 *
 * Captures <a:outerShdw>/<a:innerShdw> (outer takes priority), <a:glow>, and
 * <a:softEdge>. Alpha is extracted from the color child's <a:alpha val>
 * (1/1000 percent units).
 *
 * @param {object|undefined} spPr - parsed <p:spPr> node
 * @param {Record<string,string>|null} [themeColors] - see resolveColorNode
 * @returns {{ shadow?: {...}, glow?: {...}, softEdge?: {...} } | undefined}
 */
function resolveEffects(spPr, themeColors) {
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

    const color = resolveColorNode(shdwNode, themeColors);
    if (!color) continue;

    effects.shadow = { mode, color, blurEmu, distanceEmu, directionAngle, alphaPct: extractEffectAlpha(shdwNode) };
    break; // outer shadow takes priority; stop after first hit
  }

  const glowNode = effectLst['a:glow'];
  if (glowNode) {
    const color = resolveColorNode(glowNode, themeColors);
    if (color) {
      const radiusEmu = glowNode['@_rad'] != null ? Number(glowNode['@_rad']) : 0;
      effects.glow = { color, radiusEmu, alphaPct: extractEffectAlpha(glowNode) };
    }
  }

  const softEdgeNode = effectLst['a:softEdge'];
  if (softEdgeNode) {
    const radiusEmu = softEdgeNode['@_rad'] != null ? Number(softEdgeNode['@_rad']) : 0;
    effects.softEdge = { radiusEmu };
  }

  return Object.keys(effects).length > 0 ? effects : undefined;
}

// ---------------------------------------------------------------------------
// p:style inheritance — fill/stroke/font/effects from the shape's theme style reference
// ---------------------------------------------------------------------------

/**
 * Apply OOXML lightness modifiers (tint/shade/lumMod/lumOff) to a 6-char hex
 * color string.  Approximated in RGB space — sufficient for perceived fidelity.
 * Returns a new 6-char hex string.
 */
function applyColorModifiers(hex, mod) {
  let r = parseInt(hex.slice(0, 2), 16);
  let g = parseInt(hex.slice(2, 4), 16);
  let b = parseInt(hex.slice(4, 6), 16);
  // tint: blend toward white (R + (255-R)*t, etc.)
  if (mod['a:tint']) {
    const t = Number(mod['a:tint']['@_val']) / 100000;
    r = Math.round(r + (255 - r) * t);
    g = Math.round(g + (255 - g) * t);
    b = Math.round(b + (255 - b) * t);
  }
  // shade: darken (R * s, etc.)
  if (mod['a:shade']) {
    const s = Number(mod['a:shade']['@_val']) / 100000;
    r = Math.round(r * s);
    g = Math.round(g * s);
    b = Math.round(b * s);
  }
  // lumMod then lumOff: L' = L*lumMod + lumOff (approximated per channel)
  if (mod['a:lumMod']) {
    const m = Number(mod['a:lumMod']['@_val']) / 100000;
    r = Math.round(r * m);
    g = Math.round(g * m);
    b = Math.round(b * m);
  }
  if (mod['a:lumOff']) {
    const off = Number(mod['a:lumOff']['@_val']) / 100000;
    r = Math.round(r + (255 - r) * off);
    g = Math.round(g + (255 - g) * off);
    b = Math.round(b + (255 - b) * off);
  }
  const clamp = (c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, '0');
  return clamp(r) + clamp(g) + clamp(b);
}

/**
 * Deep-clone a parsed fill/effect node, replacing every <a:schemeClr val="phClr">
 * with the actual color child from the fillRef/effectRef element.
 *
 * When themeColors is provided and the stop has tint/shade/lumMod/lumOff modifiers,
 * the modifiers are applied to the resolved hex and the stop is emitted as
 * <a:srgbClr> so resolveColorNode picks up the correct computed color.
 */
function deepSubstitutePhClr(node, refNode, themeColors) {
  if (!node || typeof node !== 'object') return node;
  if (Array.isArray(node)) return node.map(n => deepSubstitutePhClr(n, refNode, themeColors));

  const result = {};
  for (const [key, val] of Object.entries(node)) {
    if (key === 'a:schemeClr' && val && val['@_val'] === 'phClr') {
      if (refNode['a:schemeClr']) {
        const refColorName = refNode['a:schemeClr']['@_val'];
        const baseHex = themeColors && themeColors[refColorName] && themeColors[refColorName].replace('#', '');
        const hasModifiers = val['a:tint'] || val['a:shade'] || val['a:lumMod'] || val['a:lumOff'];
        if (baseHex && baseHex.length === 6 && hasModifiers) {
          // Compute the tinted/shaded hex directly so resolveColorNode sees the real color
          result['a:srgbClr'] = { '@_val': applyColorModifiers(baseHex, val) };
        } else {
          result['a:schemeClr'] = { ...val, '@_val': refColorName };
        }
      } else if (refNode['a:srgbClr']) {
        result['a:srgbClr'] = { '@_val': refNode['a:srgbClr']['@_val'] };
      } else if (refNode['a:sysClr'] && refNode['a:sysClr']['@_lastClr']) {
        result['a:srgbClr'] = { '@_val': refNode['a:sysClr']['@_lastClr'] };
      } else {
        result[key] = deepSubstitutePhClr(val, refNode, themeColors);
      }
    } else {
      result[key] = deepSubstitutePhClr(val, refNode, themeColors);
    }
  }
  return result;
}

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
    return `var(--theme-${SCHEME_ALIAS[raw] || raw})`;
  }

  return null;
}

// ---------------------------------------------------------------------------
// Flat CSS color helper + spec-type vocabulary
// ---------------------------------------------------------------------------

/**
 * Convert a structured shapeColor ({space,hex|ref,alpha?}) to a flat CSS
 * string — the format used by spec fields fill.color, stroke.color, etc.
 * Semi-transparent sRGB colors use rgba(); theme-color alpha is dropped
 * (CSS custom properties can't carry per-use opacity directly).
 */
function shapeColorToFlatCss(c) {
  if (!c) return null;
  if (c.space === 'srgb' && c.hex) {
    // Encode alpha separately so we can propagate it alongside the hex color.
    // The generator reads fill.colorAlpha to set SVG fill-opacity (spec-compliant)
    // rather than rgba() in the fill attribute (non-standard per SVG 1.1).
    return `#${c.hex}`;
  }
  if (c.space === 'theme') {
    return `var(--theme-${c.ref})`;
  }
  return null;
}

/** Extract alpha (0–1) from a shapeColor, or null when fully opaque. */
function shapeColorAlpha(c) {
  if (!c || c.alpha == null) return null;
  const a = c.alpha / 100;
  return a < 1 ? a : null;
}

// Spec's type vocabulary — maps internal IR type to the closed enum in the
// professor's spec. Types not in the spec's 14-item list map to 'custom'.
// Note: spec writes 'ellipsis' (likely a typo for 'ellipse') — matched literally.
const SPEC_TYPE_MAP = {
  rect:       'rectangle',  roundRect:  'rectangle',
  ellipse:    'ellipsis',   triangle:   'triangle',
  rtTriangle: 'triangle',   line:       'line',
  arrow:      'arrow',      connector:  'connector',
  polyline:   'polyline',   polygon:    'polygon',
  callout:    'callout',    star:       'star',
  cloud:      'cloud',      database:   'database',
  chevron:    'chevron',    pentagon:   'custom',
  hexagon:    'custom',     octagon:    'custom',
  arc:        'custom',     unknown:    'custom',
};

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
 * Read <a:prstDash> from a <a:ln> node and return the spec stroke style.
 * spec enum: solid | dashed | pointed | none
 *   PPTX dash → pointed: 'dot', 'sysDot'
 *   PPTX dash → dashed:  everything else ('dash', 'dashDot', 'lgDash', ...)
 *   absent or 'solid'/'solidEdit' → solid
 */
function resolveStrokeDash(ln) {
  const prstDash = ln && ln['a:prstDash'];
  if (!prstDash) return 'solid';
  const val = prstDash['@_val'] || '';
  if (!val || val === 'solid' || val === 'solidEdit') return 'solid';
  if (val === 'dot' || val === 'sysDot') return 'pointed';
  return 'dashed';
}

/**
 * Resolve the fill from <p:style><a:fillRef>, using the theme's fillStyleLst
 * to recover the actual gradient when idx≥2.
 *
 * OOXML §20.1.4.2.10: fillRef idx=0 → no fill; idx≥1 → fillStyleLst[idx-1].
 * Gradient stops that use <a:schemeClr val="phClr"/> are substituted with the
 * fillRef's own color child before being resolved.
 *
 * Falls back to a solid fill using the fillRef color when fmtScheme is absent.
 *
 * @param {object|null} pStyle    - parsed <p:style> node
 * @param {object|null} fmtScheme - theme format scheme (theme.fmtScheme)
 */
function resolveStyleFill(pStyle, fmtScheme) {
  if (!pStyle) return null;
  const fillRef = pStyle['a:fillRef'];
  if (!fillRef) return null;

  const idx = Number(fillRef['@_idx']);
  if (idx === 0) return { type: 'none' };

  // Look up the actual fill entry from the theme's fillStyleLst
  const themeColors = (fmtScheme && fmtScheme.colors) || null;
  if (fmtScheme && fmtScheme.fillStyleLst && fmtScheme.fillStyleLst[idx - 1]) {
    const themeFillEntry = fmtScheme.fillStyleLst[idx - 1];
    const substituted = deepSubstitutePhClr(themeFillEntry, fillRef, themeColors);
    const resolved = resolveFill(substituted, [], themeColors);
    if (resolved && resolved.type !== 'none') return resolved;
  }

  // Fallback: solid fill from fillRef color directly
  const c = resolveColorNode(fillRef, themeColors);
  const color = shapeColorToFlatCss(c);
  if (!color) return null;
  const alpha = shapeColorAlpha(c);
  return alpha != null ? { type: 'solid', color, alpha } : { type: 'solid', color };
}

/**
 * Resolve the stroke from <p:style><a:lnRef>, reading actual line width and
 * dash style from the theme's lnStyleLst when available.
 *
 * OOXML §20.1.4.2.19: lnRef idx=0 → no line; idx≥1 → lnStyleLst[idx-1].
 * The lnRef color child gives the stroke color; width and dash come from the
 * theme entry.
 *
 * @param {object|null} pStyle    - parsed <p:style> node
 * @param {object|null} fmtScheme - theme format scheme (theme.fmtScheme)
 */
function resolveStyleStroke(pStyle, fmtScheme) {
  if (!pStyle) return null;
  const lnRef = pStyle['a:lnRef'];
  if (!lnRef) return null;

  const idx = Number(lnRef['@_idx']);
  if (idx === 0) return { type: 'none' };

  const themeColors = (fmtScheme && fmtScheme.colors) || null;
  const color = shapeColorToFlatCss(resolveColorNode(lnRef, themeColors));
  if (!color) return null;

  let widthEmu = 12700; // OOXML default (1 pt)
  let style = 'solid';
  if (fmtScheme && fmtScheme.lnStyleLst && fmtScheme.lnStyleLst[idx - 1]) {
    const lnNode = fmtScheme.lnStyleLst[idx - 1]['a:ln'];
    if (lnNode) {
      if (lnNode['@_w'] != null) widthEmu = Number(lnNode['@_w']) || 12700;
      style = resolveStrokeDash(lnNode);
    }
  }

  const width = parseFloat((widthEmu / 9525).toFixed(2));
  return { type: 'solid', color, width, style };
}

/**
 * Resolve drop-shadow / glow effects from <p:style><a:effectRef> using the
 * theme's effectStyleLst.
 *
 * OOXML §20.1.4.2.8: effectRef idx=0 → no effect; idx≥1 → effectStyleLst[idx-1].
 * Returns undefined when no effect is found (matches resolveEffects contract).
 *
 * @param {object|null} pStyle    - parsed <p:style> node
 * @param {object|null} fmtScheme - theme format scheme (theme.fmtScheme)
 */
function resolveStyleEffects(pStyle, fmtScheme) {
  if (!pStyle) return undefined;
  const effectRef = pStyle['a:effectRef'];
  if (!effectRef) return undefined;

  const idx = Number(effectRef['@_idx']);
  if (idx === 0) return undefined;

  if (!fmtScheme || !fmtScheme.effectStyleLst || !fmtScheme.effectStyleLst[idx - 1]) {
    return undefined;
  }

  const effectStyleContent = fmtScheme.effectStyleLst[idx - 1];
  if (!effectStyleContent['a:effectLst']) return undefined;

  const themeColors = fmtScheme.colors || null;
  return resolveEffects({ 'a:effectLst': effectStyleContent['a:effectLst'] }, themeColors);
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
 * @param {object}   [transform] - accumulated ancestor group transform
 *   (IDENTITY_TRANSFORM when not nested inside a <p:grpSp>)
 * @param {object|null} [fmtScheme] - theme format scheme for style fill/stroke/effect lookup
 * @returns {object|null} IR Shape, or null when the node has a placeholder
 */
function parseSp(pSp, idx, txStyles, warnings, transform = IDENTITY_TRANSFORM, fmtScheme = null) {
  const nvSpPr = pSp['p:nvSpPr'];
  const ph = nvSpPr && nvSpPr['p:nvPr'] && nvSpPr['p:nvPr']['p:ph'];
  if (ph) return null; // text placeholder — handled by text.js
  const pptxId = nvSpPr && nvSpPr['p:cNvPr'] && Number(nvSpPr['p:cNvPr']['@_id']);

  const spPr    = pSp['p:spPr'] || {};
  const prstGeom = spPr['a:prstGeom'];
  const prst     = prstGeom ? prstGeom['@_prst'] : null;

  // Geometry: apply group transform in EMU, then convert to px for the IR.
  const rawGeo = applyGroupTransform(extractShapeXfrm(spPr), transform);
  const position = {
    x: emuToPx(rawGeo.position.x) ?? 0,
    y: emuToPx(rawGeo.position.y) ?? 0,
  };
  const width    = emuToPx(rawGeo.position.w) ?? 0;
  const height   = emuToPx(rawGeo.position.h) ?? 0;
  const rotation = pptxRotationToDegrees(rawGeo.rotation);
  const { flipH, flipV } = rawGeo;

  const pStyle = pSp['p:style'];
  const themeColors = (fmtScheme && fmtScheme.colors) || null;
  const fill   = hasExplicitFillNode(spPr)
    ? resolveFill(spPr, warnings, themeColors)
    : (resolveStyleFill(pStyle, fmtScheme) || { type: 'none' });
  const stroke = hasExplicitStrokeNode(spPr)
    ? resolveStroke(spPr, themeColors)
    : (resolveStyleStroke(pStyle, fmtScheme) || { type: 'none' });

  // Internal rendering type → spec vocabulary type (+ subtype for generator dispatch).
  const irType = mapPreset(prst);
  const subtype = irType || prst || 'unknown';
  const type    = SPEC_TYPE_MAP[subtype] || 'custom';

  const shape = {
    id:       `shp-${idx}`,
    type,      // spec vocabulary: rectangle|triangle|ellipsis|line|...
    subtype,   // internal rendering type for generator dispatch (rect|roundRect|ellipse|...)
    position,
    width,
    height,
    rotation,
    flipH,
    flipV,
    fill,
    stroke,
    'z-index': 0,
  };

  // Raw PPTX preset for direction-aware rendering (e.g. rightArrow vs leftArrow).
  if (prst) shape.preset = prst;
  if (pptxId) shape._pptxId = pptxId; // slide.js uses this to resolve animation targetIds

  if (!irType) shape.supported = false;
  if (!irType) warnings.push(`shape preset "${subtype}" not yet supported by generator`);

  const adj = extractAdjustments(prstGeom, warnings);
  if (adj) {
    shape.adjustments = adj; // structured array for generator
    shape.config = Object.fromEntries(adj.map((a) => [a.name, a.value])); // spec flat bag
  }

  const effects = resolveEffects(spPr, themeColors) || resolveStyleEffects(pStyle, fmtScheme);
  if (effects) shape.effects = effects;

  const custGeo = extractCustomGeometry(spPr);
  if (custGeo) shape.customGeometry = custGeo;

  // Points constraint on subtype (internal), not spec type.
  if (subtype === 'connector' || subtype === 'polyline' || subtype === 'polygon') {
    shape.points = [];
  }

  const text = extractEmbeddedText(pSp, idx, txStyles);
  if (text) {
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
    // spec: paragraphs[] directly on shape; anchor/insets as separate fields.
    shape.paragraphs = text.paragraphs;
    if (text.anchor) shape['text-anchor'] = text.anchor;
    if (text.insets) shape['text-insets'] = text.insets;
    if (text.autoFit != null) shape.autoFit = text.autoFit;
  }

  return shape;
}

// ---------------------------------------------------------------------------
// p:cxnSp parser (connection shapes)
// ---------------------------------------------------------------------------

/**
 * Parse a <p:cxnSp> (connection shape) node into an IR Shape with
 * type:'connector'.
 *
 * @param {object}   pCxnSp   - parsed <p:cxnSp> node
 * @param {number}   idx      - 0-based counter for generating stable ids
 * @param {string[]} warnings - mutable array; push human-readable strings
 * @param {object}   [transform] - accumulated ancestor group transform
 * @param {object|null} [fmtScheme] - theme format scheme for style fill/stroke lookup
 */
function parseCxnSp(pCxnSp, idx, warnings, transform = IDENTITY_TRANSFORM, fmtScheme = null) {
  const spPr   = pCxnSp['p:spPr'] || {};
  const nvCxnSpPr = pCxnSp['p:nvCxnSpPr'];
  const cxnPptxId = nvCxnSpPr && nvCxnSpPr['p:cNvPr'] && Number(nvCxnSpPr['p:cNvPr']['@_id']);

  const rawGeo = applyGroupTransform(extractShapeXfrm(spPr), transform);
  const position = {
    x: emuToPx(rawGeo.position.x) ?? 0,
    y: emuToPx(rawGeo.position.y) ?? 0,
  };
  const width    = emuToPx(rawGeo.position.w) ?? 0;
  const height   = emuToPx(rawGeo.position.h) ?? 0;
  const rotation = pptxRotationToDegrees(rawGeo.rotation);
  const { flipH, flipV } = rawGeo;

  const pStyle = pCxnSp['p:style'];
  const themeColors = (fmtScheme && fmtScheme.colors) || null;
  const fill   = hasExplicitFillNode(spPr)
    ? resolveFill(spPr, warnings, themeColors)
    : (resolveStyleFill(pStyle, fmtScheme) || { type: 'none' });
  const stroke = hasExplicitStrokeNode(spPr)
    ? resolveStroke(spPr, themeColors)
    : (resolveStyleStroke(pStyle, fmtScheme) || { type: 'none' });

  const prstGeom = spPr['a:prstGeom'];
  const prst     = prstGeom ? prstGeom['@_prst'] : null;

  if (prst && !PRESET_TO_TYPE[prst]) {
    warnings.push(`connector preset "${prst}" not yet supported`);
  }

  const shape = {
    id:       `cxn-${idx}`,
    type:     'connector',   // spec vocabulary
    subtype:  'connector',   // internal type (always connector for cxnSp)
    position,
    width,
    height,
    rotation,
    flipH,
    flipV,
    fill,
    stroke,
    points:   [],
    'z-index': 0,
  };

  if (prst) shape.preset = prst;
  if (cxnPptxId) shape._pptxId = cxnPptxId;

  const adj = extractAdjustments(prstGeom, warnings);
  if (adj) {
    shape.adjustments = adj;
    shape.config = Object.fromEntries(adj.map((a) => [a.name, a.value]));
  }

  const effects = resolveEffects(spPr, themeColors) || resolveStyleEffects(pStyle, fmtScheme);
  if (effects) shape.effects = effects;

  return shape;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse all shape elements from an spTree into IR Shape objects, plus one IR
 * `group` entry per <p:grpSp> (FR-10 groups[]).
 *
 * Processes p:sp (non-placeholder) and p:cxnSp children at all levels of
 * nesting, including inside p:grpSp group containers. Shapes that are text
 * placeholders (have <p:ph>) are skipped — text.js handles them separately.
 *
 * Each descendant shape's position/rotation is corrected through the chain
 * of ancestor group transforms (see "Group transform composition" above) so
 * it lands in absolute slide EMU, not raw group-local coordinates.
 *
 * @param {object}   spTree   - parsed <p:spTree> node
 * @param {object}   txStyles - master txStyles passed through to text extraction
 * @param {string[]} warnings - mutable array; warnings are appended here
 * @returns {{ shapes: object[], groups: object[], topLevelGroupsByIdx: object[], pictures: object[] }}
 *   shapes   - flat list of every shape at any nesting depth (never null, may be empty)
 *   groups   - one entry per <p:grpSp> at any nesting depth, in discovery order
 *   topLevelGroupsByIdx - groups that are direct children of spTree, in document
 *     order, for the z-index walk in slide.js (mirrors p:grpSp ordinal position)
 *   pictures - one { pPic, transform, elementsOut } descriptor per <p:pic> at any
 *     nesting depth. media.js still owns id/rel/bundle-path assignment (it needs
 *     slideRels), so callers convert each via pictureToMedia(pPic, ..., transform)
 *     and push the resulting media.id into elementsOut to register it as a
 *     member of its owning group (elementsOut is a throwaway array for
 *     top-level, non-grouped pictures).
 */
function parseShapes(spTree, txStyles, warnings, fmtScheme = null) {
  if (!spTree) return { shapes: [], groups: [], topLevelGroupsByIdx: [], pictures: [] };

  const shapes = [];
  const groups = [];
  const topLevelGroupsByIdx = [];
  const pictures = [];
  let spIdx  = 0;
  let cxnIdx = 0;
  let grpIdx = 0;

  // spIdx/cxnIdx/grpIdx are shared across all nesting levels so ids remain
  // unique. elementsOut is the elements[] array of the group currently being
  // built (or a throwaway array at the spTree root, which has no group to
  // record into) — each child shape/group/picture discovered at this level
  // is recorded there (pictures via the caller, once media.js assigns an id).
  function walkTree(tree, transform, elementsOut) {
    for (const pSp of asArray(tree['p:sp'])) {
      const shape = parseSp(pSp, spIdx, txStyles, warnings, transform, fmtScheme);
      if (shape) {
        shapes.push(shape);
        elementsOut.push(shape.id);
        spIdx++;
      }
    }
    for (const pCxnSp of asArray(tree['p:cxnSp'])) {
      const shape = parseCxnSp(pCxnSp, cxnIdx++, warnings, transform, fmtScheme);
      shapes.push(shape);
      elementsOut.push(shape.id);
    }
    for (const pPic of asArray(tree['p:pic'])) {
      pictures.push({ pPic, transform, elementsOut });
    }
    for (const pGrpSp of asArray(tree['p:grpSp'])) {
      const g = extractGroupXfrm(pGrpSp);
      if (!g) {
        // No xfrm — no coordinate frame to anchor a group entry to. Flatten
        // its children under the current transform rather than emit a
        // meaningless box (matches the old Sprint-1 flatten behavior).
        warnings.push('group with no xfrm; children flattened without a group entry');
        walkTree(pGrpSp, transform, elementsOut);
        continue;
      }

      const rawGrpGeo = mapBoxThroughTransform(
        g.offX, g.offY, g.extW, g.extH, g.rotUnits, transform,
      );
      const grpPosition = {
        x: emuToPx(rawGrpGeo.position.x) ?? 0,
        y: emuToPx(rawGrpGeo.position.y) ?? 0,
      };
      const grpWidth    = emuToPx(rawGrpGeo.position.w) ?? 0;
      const grpHeight   = emuToPx(rawGrpGeo.position.h) ?? 0;
      const grpRotation = pptxRotationToDegrees(rawGrpGeo.rotation);
      const grpNvPr = pGrpSp['p:nvGrpSpPr'];
      const grpPptxId = grpNvPr && grpNvPr['p:cNvPr'] && Number(grpNvPr['p:cNvPr']['@_id']);
      const group = {
        id: `grp-${grpIdx++}`,
        elements: [],
        position: grpPosition,
        width: grpWidth,
        height: grpHeight,
        rotation: grpRotation,
        'z-index': 0, // overridden by slide.js via getSpTreeOrder/fallback walk
      };
      if (grpPptxId) group._pptxId = grpPptxId;
      groups.push(group);
      if (transform === IDENTITY_TRANSFORM) topLevelGroupsByIdx.push(group);
      elementsOut.push(group.id);

      const childTransform = composeGroupTransform(transform, g);
      walkTree(pGrpSp, childTransform, group.elements);
    }
  }

  walkTree(spTree, IDENTITY_TRANSFORM, []);
  return { shapes, groups, topLevelGroupsByIdx, pictures };
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
function parsePlaceholderBackgrounds(spTree, warnings, fmtScheme = null) {
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
    const themeColors = (fmtScheme && fmtScheme.colors) || null;

    const fill = hasExplicitFillNode(spPr)
      ? resolveFill(spPr, warnings, themeColors)
      : (resolveStyleFill(pStyle, fmtScheme) || { type: 'none' });

    if (fill.type === 'none') continue;

    const rawGeo = extractShapeXfrm(spPr);
    const position = { x: emuToPx(rawGeo.position.x) ?? 0, y: emuToPx(rawGeo.position.y) ?? 0 };
    const width    = emuToPx(rawGeo.position.w) ?? 0;
    const height   = emuToPx(rawGeo.position.h) ?? 0;
    const rotation = pptxRotationToDegrees(rawGeo.rotation);
    const { flipH, flipV } = rawGeo;
    const stroke = hasExplicitStrokeNode(spPr)
      ? resolveStroke(spPr, themeColors)
      : (resolveStyleStroke(pStyle, fmtScheme) || { type: 'none' });

    shapes.push({
      id:      `ph-bg-${idx++}`,
      type:    'rectangle', subtype: 'rect',
      position, width, height, rotation, flipH, flipV, fill, stroke,
      'z-index': 0,
    });
  }

  return shapes;
}

module.exports = {
  parseShapes,
  parsePlaceholderBackgrounds,
  extractShapeXfrm,
  resolveColorNode,
  resolveFill,
  resolveGradientFill,
  resolveStroke,
  resolveArrowEnd,
  resolveEffects,
  extractAdjustments,
  extractCustomGeometry,
  IDENTITY_TRANSFORM,
  mapBoxThroughTransform,
};
