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
function resolveColorNode(node) {
  if (!node) return null;

  const srgb = node['a:srgbClr'];
  if (srgb && srgb['@_val']) {
    return { space: 'srgb', hex: String(srgb['@_val']).toUpperCase() };
  }

  const sys = node['a:sysClr'];
  if (sys && sys['@_lastClr']) {
    return { space: 'srgb', hex: String(sys['@_lastClr']).toUpperCase() };
  }

  const scheme = node['a:schemeClr'];
  if (scheme && scheme['@_val']) {
    const raw = String(scheme['@_val']);
    const ref = SCHEME_TO_REF[raw];
    if (ref) return { space: 'theme', ref };
    // Unknown scheme slot — map to text1 so the shape still renders.
    return { space: 'theme', ref: 'text1' };
  }

  return null;
}

// ---------------------------------------------------------------------------
// Fill resolution
// ---------------------------------------------------------------------------

/**
 * Resolve <p:spPr> fill sub-nodes to an IR shapeFill object.
 *
 * Handled:
 *   <a:noFill/>       → { type:'none' }
 *   <a:solidFill>...  → { type:'solid', color: Color }
 *   gradient/pattern/group fill → { type:'none' } (stub; warning via caller)
 */
function resolveFill(spPr) {
  if (!spPr) return { type: 'none' };

  if (spPr['a:noFill']) return { type: 'none' };

  const solidFill = spPr['a:solidFill'];
  if (solidFill) {
    const color = resolveColorNode(solidFill);
    if (color) return { type: 'solid', color };
    return { type: 'none' };
  }

  // gradient / pattern / group fill — not yet implemented, emit none
  return { type: 'none' };
}

// ---------------------------------------------------------------------------
// Stroke resolution
// ---------------------------------------------------------------------------

/**
 * Resolve <p:spPr><a:ln> to an IR shapeStroke object.
 *
 * Default PPTX line width is 12700 EMU (1 point).
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
      return { type: 'solid', color, widthEmu };
    }
  }

  return { type: 'none' };
}

// ---------------------------------------------------------------------------
// Adjustments (roundRect corner radius, etc.)
// ---------------------------------------------------------------------------

/**
 * Extract shape adjustments from <a:prstGeom><a:avLst>.
 * Returns an object or undefined when none are present.
 */
function extractAdjustments(prstGeom) {
  if (!prstGeom) return undefined;
  const avLst = prstGeom['a:avLst'];
  if (!avLst) return undefined;

  const result = {};
  for (const gd of asArray(avLst['a:gd'])) {
    const name  = gd['@_name'];
    const fmla  = gd['@_fmla'];
    if (!name || !fmla) continue;
    // fmla is "val <integer>" where the integer is in 1/100000 units
    const match = String(fmla).match(/^val\s+(-?\d+)$/);
    if (match) result[name] = Number(match[1]);
  }

  return Object.keys(result).length > 0 ? result : undefined;
}

// ---------------------------------------------------------------------------
// Embedded text extraction
// ---------------------------------------------------------------------------

/**
 * Extract embedded text from a shape's <p:txBody> as an IR TextBlock.
 * Strips position/dimensions (they live on the enclosing shape) and returns
 * only id + paragraphs, or null when there is no body.
 */
function extractEmbeddedText(pSp, idx, txStyles) {
  const block = shapeToTextBlock(pSp, idx, txStyles);
  if (!block) return null;
  // Keep only the text-content fields; geometry is on the shape.
  const text = { id: block.id, paragraphs: block.paragraphs };
  return text;
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
  const fill   = resolveFill(spPr);
  const stroke = resolveStroke(spPr);

  const irType = mapPreset(prst);
  let type;
  if (irType) {
    type = irType;
  } else {
    type = 'unknown';
    const label = prst || 'custom geometry';
    warnings.push(`shape type "${label}" not yet supported`);
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

  const adj = extractAdjustments(prstGeom);
  if (adj) shape.adjustments = adj;

  // Connectors and polyline/polygon need a vertex list; for p:sp shapes these
  // come from custGeom (not yet parsed — emit empty points so the schema
  // constraint is satisfied and the generator can skip gracefully).
  if (type === 'connector' || type === 'polyline' || type === 'polygon') {
    shape.points = [];
  }

  const text = extractEmbeddedText(pSp, idx, txStyles);
  if (text) shape.text = text;

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
  const fill   = resolveFill(spPr);
  const stroke = resolveStroke(spPr);

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

  const adj = extractAdjustments(prstGeom);
  if (adj) shape.adjustments = adj;

  return shape;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Parse all shape elements from an spTree into IR Shape objects.
 *
 * Processes p:sp (non-placeholder) and p:cxnSp children.
 * Shapes that are text placeholders (have <p:ph>) are skipped — text.js
 * handles them separately so there is no duplication.
 *
 * @param {object}   spTree   - parsed <p:spTree> node
 * @param {object}   txStyles - master txStyles passed through to text extraction
 * @param {string[]} warnings - mutable array; warnings are appended here
 * @returns {object[]} array of IR Shape objects (never null, may be empty)
 */
function parseShapes(spTree, txStyles, warnings) {
  if (!spTree) return [];

  const shapes  = [];
  let spIdx     = 0;
  let cxnIdx    = 0;

  for (const pSp of asArray(spTree['p:sp'])) {
    const shape = parseSp(pSp, spIdx, txStyles, warnings);
    if (shape) {
      shapes.push(shape);
      spIdx++;
    }
  }

  for (const pCxnSp of asArray(spTree['p:cxnSp'])) {
    shapes.push(parseCxnSp(pCxnSp, cxnIdx++, warnings));
  }

  return shapes;
}

module.exports = { parseShapes, extractXfrm, resolveColorNode, resolveFill, resolveStroke };
