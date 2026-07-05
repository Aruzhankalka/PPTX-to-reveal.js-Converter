'use strict';

const { validate, validateTargetIds } = require('../src/ir/validator');
const { parseShapes, parsePlaceholderBackgrounds, extractShapeXfrm, resolveColorNode, resolveFill, resolveGradientFill, resolveStroke, resolveArrowEnd, resolveEffects, extractAdjustments, extractCustomGeometry } = require('../src/parser/pptx/shapes');
const { parseAnimations } = require('../src/parser/pptx/anim');
const fixture = require('./fixtures/ir-shapes-anim.sample.json');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function cloneFixture() {
  return JSON.parse(JSON.stringify(fixture));
}

/** Build a minimal spTree with one p:sp using the given override fields. */
function buildSpTree(spOverrides = {}) {
  return {
    'p:sp': [buildSp(spOverrides)],
  };
}

/** Build a p:sp node from explicit field values. */
function buildSp({
  prst = 'rect',
  x = 914400, y = 685800, cx = 3200400, cy = 1371600,
  rot = null,
  flipH = null,
  flipV = null,
  solidFillHex = null,
  solidFillScheme = null,
  noFill = false,
  lineWidth = null,
  hasPh = false,
  txBody = null,
} = {}) {
  const off = { '@_x': String(x), '@_y': String(y) };
  const ext = { '@_cx': String(cx), '@_cy': String(cy) };
  const xfrm = { 'a:off': off, 'a:ext': ext };
  if (rot !== null) xfrm['@_rot'] = String(rot);
  if (flipH !== null) xfrm['@_flipH'] = flipH ? '1' : '0';
  if (flipV !== null) xfrm['@_flipV'] = flipV ? '1' : '0';

  let fillNode = {};
  if (noFill) {
    fillNode = { 'a:noFill': {} };
  } else if (solidFillHex) {
    fillNode = { 'a:solidFill': { 'a:srgbClr': { '@_val': solidFillHex } } };
  } else if (solidFillScheme) {
    fillNode = { 'a:solidFill': { 'a:schemeClr': { '@_val': solidFillScheme } } };
  }

  const spPr = { 'a:xfrm': xfrm, ...fillNode };
  if (prst) {
    spPr['a:prstGeom'] = { '@_prst': prst };
  }
  if (lineWidth) {
    spPr['a:ln'] = {
      '@_w': String(lineWidth),
      'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
    };
  }

  const sp = { 'p:spPr': spPr };
  if (hasPh) {
    sp['p:nvSpPr'] = { 'p:nvPr': { 'p:ph': { '@_type': 'body' } } };
  }
  if (txBody) {
    sp['p:txBody'] = txBody;
  }

  return sp;
}

/** Build a minimal slide object with a p:timing tree containing one effect. */
function buildSlideWithAnim({
  presetClass = 'entr',
  presetID = '21',
  nodeType = 'clickEffect',
  spid = '3',
  dur = '500',
  delay = '0',
  noTiming = false,
  noTarget = false,
} = {}) {
  if (noTiming) return {};

  const tgtEl = noTarget
    ? {}
    : { 'p:spTgt': { '@_spid': spid } };

  const tweenNode = {
    'p:cBhvr': {
      'p:cTn': { '@_id': '10', '@_dur': '1' },
      'p:tgtEl': tgtEl,
      'p:attrNameLst': { 'p:attrName': 'style.visibility' },
    },
  };

  const effectCTn = {
    '@_id': '4',
    '@_presetID': presetID,
    '@_presetClass': presetClass,
    '@_nodeType': nodeType,
    '@_dur': dur,
    'p:stCondLst': { 'p:cond': { '@_delay': delay } },
    'p:childTnLst': { 'p:set': tweenNode },
  };

  const clickCTn = {
    '@_id': '3',
    '@_nodeType': 'click',
    'p:childTnLst': { 'p:par': { 'p:cTn': effectCTn } },
  };

  const mainSeqCTn = {
    '@_id': '2',
    '@_nodeType': 'mainSeq',
    'p:childTnLst': { 'p:par': { 'p:cTn': clickCTn } },
  };

  const rootCTn = {
    '@_id': '1',
    '@_nodeType': 'tmRoot',
    'p:childTnLst': { 'p:seq': { 'p:cTn': mainSeqCTn } },
  };

  return {
    'p:timing': {
      'p:tnLst': { 'p:par': { 'p:cTn': rootCTn } },
    },
  };
}

// ===========================================================================
// 1. Schema — fixture and negative cases
// ===========================================================================

describe('schema — fixture validates', () => {
  test('fixture is accepted by the schema + targetId check', () => {
    const { valid, errors } = validate(fixture);
    expect(errors).toBeNull();
    expect(valid).toBe(true);
  });
});

describe('schema — shape required fields', () => {
  test('rejects a Shape missing position', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0].position;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a Shape missing id', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0].id;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a Shape missing z-index', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0]['z-index'];
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('accepts any spec vocabulary type', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].type = 'custom'; // PPTX-specific name goes in subtype
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });

  test('rejects a Shape with a non-string type', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].type = 42;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a position with missing y', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0].position.y;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });
});

describe('schema — conditional: polyline/polygon/connector require points', () => {
  test('rejects a polyline without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline'; shape.subtype = 'polyline';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a polygon without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polygon'; shape.subtype = 'polygon';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a connector without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'connector'; shape.subtype = 'connector';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('accepts a polyline with an empty points array', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline'; shape.subtype = 'polyline'; shape.points = [];
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });

  test('accepts a polyline with populated points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline'; shape.subtype = 'polyline';
    shape.points = [{ x: 0, y: 0 }, { x: 100, y: 80 }];
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });
});

describe('schema — fill and stroke color types', () => {
  test('accepts theme-color fill as flat CSS var string', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid', color: 'var(--theme-accent2)',
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects fill.color as non-string (e.g. number)', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid', color: 42,
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects fill.color as boolean', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid', color: true,
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('accepts fill type:none', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = { type: 'none' };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects solid fill missing color', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = { type: 'solid' };
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — animation fields', () => {
  test('rejects an Animation missing a spec-required field (effect)', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.animations[0].effect;
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects an Animation with invalid trigger', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0].trigger = 'onHover';
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects an Animation with invalid effect-detail.class', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0]['effect-detail'].class = 'dance';
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — targetId cross-reference', () => {
  test('rejects an Animation whose targetId has no matching element', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0].targetId = 'nonexistent-element';
    const { valid, errors } = validate(doc);
    expect(valid).toBe(false);
    expect(errors[0].message).toContain('nonexistent-element');
  });

  test('accepts animations when all targetIds resolve', () => {
    expect(validate(fixture).valid).toBe(true);
  });

  test('validateTargetIds returns empty array for valid doc', () => {
    expect(validateTargetIds(fixture)).toHaveLength(0);
  });

  test('validateTargetIds detects missing reference', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0].targetId = 'ghost-id';
    const errors = validateTargetIds(doc);
    expect(errors.length).toBeGreaterThan(0);
    expect(errors[0].instancePath).toContain('animations/0/targetId');
  });
});

// ===========================================================================
// 2. shapes.js unit tests
// ===========================================================================

describe('extractShapeXfrm', () => {
  test('reads x/y/w/h as raw EMU integers', () => {
    const spPr = {
      'a:xfrm': {
        'a:off': { '@_x': '914400', '@_y': '685800' },
        'a:ext': { '@_cx': '3200400', '@_cy': '1371600' },
      },
    };
    const { position } = extractShapeXfrm(spPr);
    expect(position).toEqual({ x: 914400, y: 685800, w: 3200400, h: 1371600 });
  });

  test('rotation is stored in raw PPTX units (NOT divided by 60000)', () => {
    const spPr = {
      'a:xfrm': {
        '@_rot': '5400000',
        'a:off': { '@_x': '0', '@_y': '0' },
        'a:ext': { '@_cx': '0', '@_cy': '0' },
      },
    };
    const { rotation } = extractShapeXfrm(spPr);
    expect(rotation).toBe(5400000); // 90 degrees in PPTX units
  });

  test('missing xfrm returns zeroed position and rotation 0', () => {
    const { position, rotation } = extractShapeXfrm(null);
    expect(position).toEqual({ x: 0, y: 0, w: 0, h: 0 });
    expect(rotation).toBe(0);
  });

  test('flipH and flipV are parsed as booleans', () => {
    const spPr = {
      'a:xfrm': {
        '@_flipH': '1',
        '@_flipV': '0',
        'a:off': { '@_x': '0', '@_y': '0' },
        'a:ext': { '@_cx': '0', '@_cy': '0' },
      },
    };
    const { flipH, flipV } = extractShapeXfrm(spPr);
    expect(flipH).toBe(true);
    expect(flipV).toBe(false);
  });
});

describe('resolveColorNode', () => {
  test('srgbClr maps to {space:srgb, hex} uppercase', () => {
    const node = { 'a:srgbClr': { '@_val': '4472c4' } };
    expect(resolveColorNode(node)).toEqual({ space: 'srgb', hex: '4472C4' });
  });

  test('sysClr lastClr maps to {space:srgb, hex}', () => {
    const node = { 'a:sysClr': { '@_lastClr': '000000' } };
    expect(resolveColorNode(node)).toEqual({ space: 'srgb', hex: '000000' });
  });

  test('schemeClr accent1 maps to {space:theme, ref:accent1}', () => {
    const node = { 'a:schemeClr': { '@_val': 'accent1' } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'accent1' });
  });

  test('schemeClr dk1 maps to {space:theme, ref:text1}', () => {
    const node = { 'a:schemeClr': { '@_val': 'dk1' } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'text1' });
  });

  test('schemeClr tx1 alias maps to text1', () => {
    const node = { 'a:schemeClr': { '@_val': 'tx1' } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'text1' });
  });

  test('schemeClr hlink maps to link', () => {
    const node = { 'a:schemeClr': { '@_val': 'hlink' } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'link' });
  });

  test('unknown scheme name falls back to text1', () => {
    const node = { 'a:schemeClr': { '@_val': 'neonPink' } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'text1' });
  });

  test('null input returns null', () => {
    expect(resolveColorNode(null)).toBeNull();
  });

  test('node without recognized child returns null', () => {
    expect(resolveColorNode({})).toBeNull();
  });
});

describe('resolveFill', () => {
  test('noFill → {type:none}', () => {
    expect(resolveFill({ 'a:noFill': {} })).toEqual({ type: 'none' });
  });

  test('solidFill with srgbClr → {type:solid, color: flat CSS hex}', () => {
    const spPr = { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } };
    expect(resolveFill(spPr)).toEqual({ type: 'solid', color: '#FF0000' });
  });

  test('solidFill with schemeClr → {type:solid, color: flat CSS var}', () => {
    const spPr = { 'a:solidFill': { 'a:schemeClr': { '@_val': 'accent2' } } };
    expect(resolveFill(spPr)).toEqual({ type: 'solid', color: 'var(--theme-accent2)' });
  });

  test('gradFill with empty gsLst (0 stops) → {type:none} + warning', () => {
    const warnings = [];
    const spPr = { 'a:gradFill': { 'a:gsLst': {} } };
    expect(resolveFill(spPr, warnings)).toEqual({ type: 'none' });
    expect(warnings.length).toBeGreaterThan(0);
    expect(warnings[0]).toContain('stop');
  });

  test('gradFill with 2 sRGB stops → {type:gradient, kind:linear, stops}', () => {
    const spPr = {
      'a:gradFill': {
        'a:gsLst': {
          'a:gs': [
            { '@_pos': '0',      'a:srgbClr': { '@_val': 'FF0000' } },
            { '@_pos': '100000', 'a:srgbClr': { '@_val': '0000FF' } },
          ],
        },
        'a:lin': { '@_ang': '0' },
      },
    };
    const result = resolveFill(spPr, []);
    expect(result.type).toBe('gradient');
    expect(result.kind).toBe('linear');
    expect(result.stops).toHaveLength(2);
  });

  test('null spPr → {type:none}', () => {
    expect(resolveFill(null)).toEqual({ type: 'none' });
  });
});

describe('resolveStroke', () => {
  test('no ln node → {type:none}', () => {
    expect(resolveStroke({})).toEqual({ type: 'none' });
  });

  test('ln with noFill → {type:none}', () => {
    expect(resolveStroke({ 'a:ln': { 'a:noFill': {} } })).toEqual({ type: 'none' });
  });

  test('ln with solidFill and explicit width → {type:solid, color (flat), width (px), style}', () => {
    const spPr = {
      'a:ln': {
        '@_w': '25400',
        'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
      },
    };
    const stroke = resolveStroke(spPr);
    expect(stroke.type).toBe('solid');
    expect(stroke.color).toBe('#000000');
    expect(stroke.width).toBeCloseTo(25400 / 9525, 1);
    expect(stroke.style).toBe('solid');
  });

  test('ln with solidFill and no width defaults to 12700 EMU (≈1.33 px)', () => {
    const spPr = {
      'a:ln': {
        'a:solidFill': { 'a:srgbClr': { '@_val': 'FFFFFF' } },
      },
    };
    const stroke = resolveStroke(spPr);
    expect(stroke.type).toBe('solid');
    expect(stroke.width).toBeCloseTo(12700 / 9525, 1);
    expect(stroke.style).toBe('solid');
  });
});

describe('parseShapes — shape type dispatch', () => {
  test('rect preset maps to spec type:rectangle (subtype:rect)', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('rectangle');
    expect(shapes[0].subtype).toBe('rect');
  });

  test('roundRect preset maps to type:rectangle (subtype:roundRect)', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'roundRect' }), null, []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('rectangle');
    expect(shapes[0].subtype).toBe('roundRect');
  });

  test('ellipse preset maps to spec type:ellipsis (subtype:ellipse)', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'ellipse' }), null, []);
    expect(shapes[0].type).toBe('ellipsis');
    expect(shapes[0].subtype).toBe('ellipse');
  });

  test('line preset maps to type:line', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'line' }), null, []);
    expect(shapes[0].type).toBe('line');
  });

  test('rightArrow maps to type:arrow', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rightArrow' }), null, []);
    expect(shapes[0].type).toBe('arrow');
  });

  test('wedgeRectCallout maps to type:callout', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'wedgeRectCallout' }), null, []);
    expect(shapes[0].type).toBe('callout');
  });

  test('bentConnector3 maps to type:connector', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'bentConnector3' }), null, []);
    expect(shapes[0].type).toBe('connector');
  });
});

describe('parseShapes — unsupported preset maps to type:custom + subtype=PPTX name', () => {
  test('unrecognized preset gets spec type:custom (PPTX name in subtype)', () => {
    const warnings = [];
    const { shapes } = parseShapes(buildSpTree({ prst: 'star50' }), null, warnings);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('custom');
    expect(shapes[0].subtype).toBe('star50');
  });

  test('unrecognized preset sets supported:false', () => {
    const warnings = [];
    const { shapes } = parseShapes(buildSpTree({ prst: 'star50' }), null, warnings);
    expect(shapes[0].supported).toBe(false);
  });

  test('unsupported preset pushes a warning containing the preset name', () => {
    const warnings = [];
    parseShapes(buildSpTree({ prst: 'star50' }), null, warnings);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('star50');
  });

  test('shape is NOT dropped for unsupported preset', () => {
    const warnings = [];
    const { shapes } = parseShapes(buildSpTree({ prst: 'cloudCallout2000' }), null, warnings);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('custom');
    expect(shapes[0].subtype).toBe('cloudCallout2000');
  });

  test('recognized preset does NOT set supported:false', () => {
    const warnings = [];
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect' }), null, warnings);
    expect(shapes[0].type).toBe('rectangle');
    expect(shapes[0].supported).toBeUndefined();
  });

  test('missing prst (custom geometry) emits type:custom, subtype:unknown + supported:false', () => {
    const spTree = buildSpTree({ prst: null });
    delete spTree['p:sp'][0]['p:spPr']['a:prstGeom'];
    const warnings = [];
    const { shapes } = parseShapes(spTree, null, warnings);
    expect(shapes[0].type).toBe('custom');
    expect(shapes[0].subtype).toBe('unknown');
    expect(shapes[0].supported).toBe(false);
    expect(warnings).toHaveLength(1);
  });
});

describe('parseShapes — geometry fields', () => {
  test('position is in CSS px, width/height as separate top-level fields', () => {
    const { shapes } = parseShapes(
      buildSpTree({ prst: 'rect', x: 914400, y: 685800, cx: 3200400, cy: 1371600 }),
      null, [],
    );
    expect(shapes[0].position).toEqual({ x: 96, y: 72 }); // 914400/9525=96, 685800/9525=72
    expect(shapes[0].width).toBe(336);   // 3200400/9525=336
    expect(shapes[0].height).toBe(144);  // 1371600/9525=144
  });

  test('rotation stored in degrees (not raw PPTX rot units)', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect', rot: 3000000 }), null, []);
    expect(shapes[0].rotation).toBeCloseTo(50, 1); // 3000000/60000=50 degrees
  });

  test('flipH=true is preserved', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect', flipH: true }), null, []);
    expect(shapes[0].flipH).toBe(true);
  });

  test('roundRect with avLst adj emits adjustments array [{name, value}]', () => {
    const spTree = buildSpTree({ prst: 'roundRect' });
    spTree['p:sp'][0]['p:spPr']['a:prstGeom']['a:avLst'] = {
      'a:gd': { '@_name': 'adj', '@_fmla': 'val 16667' },
    };
    const { shapes } = parseShapes(spTree, null, []);
    expect(Array.isArray(shapes[0].adjustments)).toBe(true);
    expect(shapes[0].adjustments).toEqual([{ name: 'adj', value: 16667 }]);
  });
});

describe('parseShapes — fill and stroke', () => {
  test('solidFill hex emits fill.color as flat CSS hex string', () => {
    const { shapes } = parseShapes(
      buildSpTree({ prst: 'rect', solidFillHex: '4472C4' }), null, [],
    );
    expect(shapes[0].fill).toEqual({ type: 'solid', color: '#4472C4' });
  });

  test('solidFill scheme emits fill.color as flat CSS var string', () => {
    const { shapes } = parseShapes(
      buildSpTree({ prst: 'rect', solidFillScheme: 'accent3' }), null, [],
    );
    expect(shapes[0].fill).toEqual({ type: 'solid', color: 'var(--theme-accent3)' });
  });

  test('noFill emits fill.type:none', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect', noFill: true }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'none' });
  });

  test('stroke with explicit width stores width in px and style', () => {
    const { shapes } = parseShapes(
      buildSpTree({ prst: 'rect', lineWidth: 25400 }), null, [],
    );
    expect(shapes[0].stroke.type).toBe('solid');
    expect(shapes[0].stroke.width).toBeCloseTo(25400 / 9525, 1);
    expect(shapes[0].stroke.style).toBe('solid');
  });
});

describe('parseShapes — p:style fill/stroke inheritance', () => {
  function buildSpWithStyle({ fillRefIdx = 1, fillScheme = null, fillHex = null, lnRefIdx = 1, lnScheme = null } = {}) {
    const spPr = {
      'a:xfrm': {
        'a:off': { '@_x': '914400', '@_y': '685800' },
        'a:ext': { '@_cx': '3200400', '@_cy': '1371600' },
      },
      'a:prstGeom': { '@_prst': 'rect' },
      // deliberately no fill or ln node — forces style fallback
    };

    const fillColor = fillScheme
      ? { 'a:schemeClr': { '@_val': fillScheme } }
      : fillHex
        ? { 'a:srgbClr': { '@_val': fillHex } }
        : {};

    const lnColor = lnScheme
      ? { 'a:schemeClr': { '@_val': lnScheme } }
      : {};

    const pStyle = {
      'a:fillRef': { '@_idx': String(fillRefIdx), ...fillColor },
      'a:lnRef':   { '@_idx': String(lnRefIdx),   ...lnColor },
    };

    return { 'p:sp': [{ 'p:spPr': spPr, 'p:style': pStyle }] };
  }

  test('fillRef with scheme color produces solid fill with flat CSS var', () => {
    const { shapes } = parseShapes(buildSpWithStyle({ fillScheme: 'accent2' }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'solid', color: 'var(--theme-accent2)' });
  });

  test('fillRef with srgb color produces solid fill with flat CSS hex', () => {
    const { shapes } = parseShapes(buildSpWithStyle({ fillHex: 'FF0000' }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'solid', color: '#FF0000' });
  });

  test('fillRef idx=0 produces fill.type:none', () => {
    const { shapes } = parseShapes(buildSpWithStyle({ fillRefIdx: 0, fillScheme: 'accent1' }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'none' });
  });

  test('lnRef with scheme color produces solid stroke with flat CSS var', () => {
    const { shapes } = parseShapes(buildSpWithStyle({ lnScheme: 'accent1' }), null, []);
    expect(shapes[0].stroke.type).toBe('solid');
    expect(shapes[0].stroke.color).toBe('var(--theme-accent1)');
    expect(shapes[0].stroke.style).toBe('solid');
    expect(shapes[0].stroke.width).toBeCloseTo(12700 / 9525, 1);
  });

  test('lnRef idx=0 produces stroke.type:none', () => {
    const { shapes } = parseShapes(buildSpWithStyle({ lnRefIdx: 0, lnScheme: 'accent1' }), null, []);
    expect(shapes[0].stroke).toEqual({ type: 'none' });
  });

  test('explicit spPr solidFill overrides fillRef color', () => {
    const { shapes } = parseShapes(
      buildSpTree({ prst: 'rect', solidFillHex: 'AABBCC' }), null, [],
    );
    expect(shapes[0].fill).toEqual({ type: 'solid', color: '#AABBCC' });
  });

  test('explicit spPr noFill overrides fillRef — remains none', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect', noFill: true }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'none' });
  });

  test('no p:style and no spPr fill produces fill.type:none', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'none' });
  });
});

describe('parseShapes — placeholder shapes are skipped', () => {
  test('placeholder p:sp is not included in shapes output', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect', hasPh: true }), null, []);
    expect(shapes).toHaveLength(0);
  });
});

describe('parseShapes — connector and points', () => {
  test('connector type emits an empty points array', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'bentConnector3' }), null, []);
    expect(Array.isArray(shapes[0].points)).toBe(true);
  });

  test('null/empty spTree returns empty shapes array', () => {
    expect(parseShapes(null, null, []).shapes).toEqual([]);
    expect(parseShapes({}, null, []).shapes).toEqual([]);
  });
});

describe('parseShapes — z-index placeholder', () => {
  test('z-index is 0 (slide.js assigns final value)', () => {
    const { shapes } = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(shapes[0]['z-index']).toBe(0);
  });
});

describe('parseShapes — groups[]', () => {
  function buildSpTreeWithGroup({ childPrests = ['rect'], groupX = 0, groupY = 0, groupW = 5000000, groupH = 5000000 } = {}) {
    const children = childPrests.map((prst) => buildSp({ prst }));
    return {
      'p:grpSp': [{
        'p:grpSpPr': {
          'a:xfrm': {
            'a:off': { '@_x': String(groupX), '@_y': String(groupY) },
            'a:ext': { '@_cx': String(groupW), '@_cy': String(groupH) },
          },
        },
        'p:sp': children,
      }],
    };
  }

  test('a p:grpSp produces one entry in groups[]', () => {
    const { groups } = parseShapes(buildSpTreeWithGroup(), null, []);
    expect(groups).toHaveLength(1);
  });

  test('group gets a stable id prefixed grp-', () => {
    const { groups } = parseShapes(buildSpTreeWithGroup(), null, []);
    expect(groups[0].id).toMatch(/^grp-\d+$/);
  });

  test('group.elements lists the ids of child shapes', () => {
    const { shapes, groups } = parseShapes(buildSpTreeWithGroup({ childPrests: ['rect', 'ellipse'] }), null, []);
    expect(shapes).toHaveLength(2);
    expect(groups[0].elements).toEqual([shapes[0].id, shapes[1].id]);
  });

  test('group.position is in CSS px with separate width/height', () => {
    const { groups } = parseShapes(buildSpTreeWithGroup({ groupX: 914400, groupY: 685800, groupW: 3200400, groupH: 1371600 }), null, []);
    expect(groups[0].position).toEqual({ x: 96, y: 72 }); // 914400/9525=96, 685800/9525=72
    expect(groups[0].width).toBe(336);   // 3200400/9525=336
    expect(groups[0].height).toBe(144);  // 1371600/9525=144
  });

  test('nested groups: outer.elements includes inner group id', () => {
    const innerGrpSp = [{
      'p:grpSpPr': { 'a:xfrm': { 'a:off': { '@_x': '0', '@_y': '0' }, 'a:ext': { '@_cx': '1000000', '@_cy': '1000000' } } },
      'p:sp': [buildSp({ prst: 'rect' })],
    }];
    const spTree = {
      'p:grpSp': [{
        'p:grpSpPr': { 'a:xfrm': { 'a:off': { '@_x': '0', '@_y': '0' }, 'a:ext': { '@_cx': '5000000', '@_cy': '5000000' } } },
        'p:grpSp': innerGrpSp,
      }],
    };
    const { groups } = parseShapes(spTree, null, []);
    expect(groups).toHaveLength(2);
    const outerGroup = groups.find((g) => g.elements.some((id) => id.startsWith('grp-')));
    expect(outerGroup).toBeDefined();
  });

  test('no groups in spTree yields empty groups array', () => {
    const { groups } = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(groups).toEqual([]);
  });

  test('topLevelGroupsByIdx[0] matches first top-level group', () => {
    const { groups, topLevelGroupsByIdx } = parseShapes(buildSpTreeWithGroup(), null, []);
    expect(topLevelGroupsByIdx).toHaveLength(1);
    expect(topLevelGroupsByIdx[0]).toBe(groups[0]);
  });
});

// ===========================================================================
// 3. anim.js unit tests
// ===========================================================================

describe('parseAnimations — basic', () => {
  test('missing timing returns empty animations array', () => {
    const { animations, warnings } = parseAnimations({});
    expect(animations).toEqual([]);
    expect(warnings).toEqual([]);
  });

  test('null slide returns empty animations array', () => {
    const { animations } = parseAnimations(null);
    expect(animations).toEqual([]);
  });

  test('p:timing without tnLst returns empty array', () => {
    const { animations } = parseAnimations({ 'p:timing': {} });
    expect(animations).toEqual([]);
  });
});

describe('parseAnimations — mappable effect', () => {
  test('known presetClass entr + presetID 21 → supported:true, class:entrance', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations).toHaveLength(1);
    expect(animations[0].supported).toBe(true);
    // effect is now a plain string (spec); class lives in effect-detail
    expect(animations[0]['effect-detail'].class).toBe('entrance');
  });

  test('presetID 21 maps to effect "fade" (plain string per spec)', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].effect).toBe('fade');
    expect(animations[0]['effect-detail'].preset).toBe('fade');
  });

  test('timing delayMs and durationMs are extracted', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '3', dur: '750', delay: '200' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].timing.durationMs).toBe(750);
    expect(animations[0].timing.delayMs).toBe(200);
  });

  test('targetId includes the spid', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '99' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].targetId).toContain('99');
  });

  test('nodeType clickEffect → trigger:onClick', () => {
    const sld = buildSlideWithAnim({ nodeType: 'clickEffect', presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].trigger).toBe('onClick');
  });

  test('nodeType withEffect → trigger:withPrevious', () => {
    const sld = buildSlideWithAnim({ nodeType: 'withEffect', presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].trigger).toBe('withPrevious');
  });

  test('nodeType afterEffect → trigger:afterPrevious', () => {
    const sld = buildSlideWithAnim({ nodeType: 'afterEffect', presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].trigger).toBe('afterPrevious');
  });
});

describe('parseAnimations — unmappable effect', () => {
  test('unknown presetClass → supported:false', () => {
    const sld = buildSlideWithAnim({ presetClass: 'dance', presetID: '21', spid: '3' });
    const { animations, warnings } = parseAnimations(sld);
    expect(animations[0].supported).toBe(false);
    expect(warnings).toHaveLength(1);
  });

  test('unknown presetID → supported:false + warning', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '9999', spid: '3' });
    const { animations, warnings } = parseAnimations(sld);
    expect(animations[0].supported).toBe(false);
    expect(warnings.length).toBeGreaterThan(0);
  });

  test('no target element → supported:false + warning mentioning target', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', noTarget: true });
    const { animations, warnings } = parseAnimations(sld);
    expect(animations[0].supported).toBe(false);
    expect(warnings[0]).toContain('target');
  });

  test('external warnings array is populated', () => {
    const sld = buildSlideWithAnim({ presetClass: 'boom', presetID: '0', spid: '3' });
    const externalWarnings = [];
    parseAnimations(sld, externalWarnings);
    expect(externalWarnings.length).toBeGreaterThan(0);
  });
});

describe('parsePlaceholderBackgrounds', () => {
  function buildLayoutSpTree({ fillHex = null, fillScheme = null, noFill = false, noPh = false } = {}) {
    const spPr = {
      'a:xfrm': {
        'a:off': { '@_x': '914400', '@_y': '685800' },
        'a:ext': { '@_cx': '3200400', '@_cy': '1371600' },
      },
    };
    if (noFill) {
      spPr['a:noFill'] = {};
    } else if (fillHex) {
      spPr['a:solidFill'] = { 'a:srgbClr': { '@_val': fillHex } };
    } else if (fillScheme) {
      spPr['a:solidFill'] = { 'a:schemeClr': { '@_val': fillScheme } };
    }

    const sp = { 'p:spPr': spPr };
    if (!noPh) {
      sp['p:nvSpPr'] = { 'p:nvPr': { 'p:ph': { '@_type': 'body' } } };
    }
    return { 'p:sp': [sp] };
  }

  test('placeholder with solidFill sRGB produces a rectangle background shape with flat CSS color', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ fillHex: '92D050' }), []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('rectangle');
    expect(shapes[0].fill).toEqual({ type: 'solid', color: '#92D050' });
  });

  test('placeholder with solidFill scheme produces fill with flat CSS var', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ fillScheme: 'accent3' }), []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].fill).toEqual({ type: 'solid', color: 'var(--theme-accent3)' });
  });

  test('placeholder with noFill is NOT included (fill.type:none → skip)', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ noFill: true }), []);
    expect(shapes).toHaveLength(0);
  });

  test('placeholder with no fill node is NOT included', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree(), []);
    expect(shapes).toHaveLength(0);
  });

  test('non-placeholder sp is ignored (handled by parseShapes)', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ fillHex: 'FF0000', noPh: true }), []);
    expect(shapes).toHaveLength(0);
  });

  test('null spTree returns empty array', () => {
    expect(parsePlaceholderBackgrounds(null, [])).toEqual([]);
  });

  test('id has expected ph-bg prefix', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ fillHex: 'AABBCC' }), []);
    expect(shapes[0].id).toMatch(/^ph-bg-/);
  });

  test('position is extracted in CSS px with separate width/height', () => {
    const shapes = parsePlaceholderBackgrounds(buildLayoutSpTree({ fillHex: '0BDA7E' }), []);
    expect(shapes[0].position).toEqual({ x: 96, y: 72 }); // 914400/9525=96, 685800/9525=72
    expect(shapes[0].width).toBe(336);   // 3200400/9525=336
    expect(shapes[0].height).toBe(144);  // 1371600/9525=144
  });
});

// ===========================================================================
// 4. New shape-effect unit tests
// ===========================================================================

// ---------------------------------------------------------------------------
// Helper: build a minimal <a:gradFill> node
// ---------------------------------------------------------------------------
function makeGradFill({ stopDefs = [], hasLin = false, angle = 0, hasPath = false } = {}) {
  const gsList = stopDefs.map(([pos, hex, scheme]) => {
    const gs = { '@_pos': String(pos) };
    if (hex)    gs['a:srgbClr']  = { '@_val': hex };
    if (scheme) gs['a:schemeClr'] = { '@_val': scheme };
    return gs;
  });
  const result = { 'a:gsLst': {} };
  if (gsList.length > 0) result['a:gsLst']['a:gs'] = gsList.length === 1 ? gsList[0] : gsList;
  if (hasLin)  result['a:lin']  = { '@_ang': String(angle) };
  if (hasPath) result['a:path'] = { '@_path': 'circle' };
  return result;
}

describe('resolveGradientFill', () => {
  test('2 sRGB stops + lin → gradient linear with stops', () => {
    const gradFill = makeGradFill({
      stopDefs: [[0, 'FF0000'], [100000, '0000FF']],
      hasLin: true, angle: 5400000,
    });
    const w = [];
    const result = resolveGradientFill(gradFill, w);
    expect(result).toMatchObject({ type: 'gradient', kind: 'linear', angle: 5400000 });
    expect(result.stops).toHaveLength(2);
    expect(result.stops[0]).toEqual({ pos: 0,      color: { space: 'srgb', hex: 'FF0000' } });
    expect(result.stops[1]).toEqual({ pos: 100000, color: { space: 'srgb', hex: '0000FF' } });
    expect(w).toHaveLength(0);
  });

  test('2 stops + path element → kind:radial (no angle)', () => {
    const gradFill = makeGradFill({
      stopDefs: [[0, 'FF0000'], [100000, '0000FF']],
      hasPath: true,
    });
    const result = resolveGradientFill(gradFill, []);
    expect(result).toMatchObject({ type: 'gradient', kind: 'radial' });
    expect(result.angle).toBeUndefined();
  });

  test('stops are sorted by position ascending', () => {
    const gradFill = makeGradFill({
      stopDefs: [[100000, 'FFFFFF'], [0, '000000']],
      hasLin: true,
    });
    const result = resolveGradientFill(gradFill, []);
    expect(result.stops[0].pos).toBe(0);
    expect(result.stops[1].pos).toBe(100000);
  });

  test('schemeClr in stop resolves to theme ref', () => {
    const gradFill = makeGradFill({
      stopDefs: [[0, null, 'accent1'], [100000, 'FFFFFF']],
      hasLin: true,
    });
    const result = resolveGradientFill(gradFill, []);
    expect(result.stops[0].color).toEqual({ space: 'theme', ref: 'accent1' });
  });

  test('missing gsLst → null + warning containing "gsLst"', () => {
    const w = [];
    const result = resolveGradientFill({ 'a:gsLst': undefined }, w);
    expect(result).toBeNull();
    expect(w[0]).toContain('gsLst');
  });

  test('only 1 resolvable stop → null + warning containing "stop"', () => {
    const gradFill = makeGradFill({ stopDefs: [[0, 'FF0000']] });
    const w = [];
    const result = resolveGradientFill(gradFill, w);
    expect(result).toBeNull();
    expect(w[0]).toContain('stop');
  });

  test('no lin or path defaults to linear angle 0', () => {
    const gradFill = makeGradFill({ stopDefs: [[0, 'FF0000'], [100000, '000000']] });
    const result = resolveGradientFill(gradFill, []);
    expect(result).toMatchObject({ type: 'gradient', kind: 'linear', angle: 0 });
  });
});

describe('resolveArrowEnd', () => {
  test('undefined → undefined', () => {
    expect(resolveArrowEnd(undefined)).toBeUndefined();
  });

  test('type:none → undefined', () => {
    expect(resolveArrowEnd({ '@_type': 'none' })).toBeUndefined();
  });

  test('absent type attribute defaults to none → undefined', () => {
    expect(resolveArrowEnd({})).toBeUndefined();
  });

  test('triangle with width and length', () => {
    expect(resolveArrowEnd({ '@_type': 'triangle', '@_w': 'med', '@_len': 'lg' }))
      .toEqual({ type: 'triangle', width: 'med', length: 'lg' });
  });

  test('arrow without size attributes → only type field', () => {
    expect(resolveArrowEnd({ '@_type': 'arrow' })).toEqual({ type: 'arrow' });
  });

  test('stealth with sm/sm sizes', () => {
    expect(resolveArrowEnd({ '@_type': 'stealth', '@_w': 'sm', '@_len': 'sm' }))
      .toEqual({ type: 'stealth', width: 'sm', length: 'sm' });
  });
});

describe('resolveStroke — arrowheads', () => {
  test('headEnd and tailEnd are captured in the stroke', () => {
    const spPr = {
      'a:ln': {
        '@_w': '12700',
        'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
        'a:headEnd': { '@_type': 'triangle', '@_w': 'med', '@_len': 'med' },
        'a:tailEnd':  { '@_type': 'arrow',    '@_w': 'lg',  '@_len': 'lg'  },
      },
    };
    const stroke = resolveStroke(spPr);
    expect(stroke.type).toBe('solid');
    expect(stroke.headEnd).toEqual({ type: 'triangle', width: 'med', length: 'med' });
    expect(stroke.tailEnd).toEqual({ type: 'arrow', width: 'lg', length: 'lg' });
  });

  test('no end markers → headEnd and tailEnd absent from output', () => {
    const spPr = { 'a:ln': { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } } };
    const stroke = resolveStroke(spPr);
    expect(stroke.headEnd).toBeUndefined();
    expect(stroke.tailEnd).toBeUndefined();
  });

  test('type:none end markers are omitted (not emitted as {type:"none"})', () => {
    const spPr = {
      'a:ln': {
        'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } },
        'a:headEnd': { '@_type': 'none' },
        'a:tailEnd': { '@_type': 'none' },
      },
    };
    const stroke = resolveStroke(spPr);
    expect(stroke.headEnd).toBeUndefined();
    expect(stroke.tailEnd).toBeUndefined();
  });
});

describe('resolveEffects', () => {
  function makeShadowSpPr({
    mode = 'outer', blurRad = '50800', dist = '38100',
    dir = '2700000', colorHex = '000000', alphaVal = '50000',
  } = {}) {
    const tag = mode === 'outer' ? 'a:outerShdw' : 'a:innerShdw';
    const colorNode = { '@_val': colorHex };
    if (alphaVal != null) colorNode['a:alpha'] = { '@_val': alphaVal };
    return {
      'a:effectLst': {
        [tag]: { '@_blurRad': blurRad, '@_dist': dist, '@_dir': dir, 'a:srgbClr': colorNode },
      },
    };
  }

  test('outer shadow extracts all fields', () => {
    const effects = resolveEffects(makeShadowSpPr());
    expect(effects).toBeDefined();
    expect(effects.shadow).toMatchObject({
      mode: 'outer', blurEmu: 50800, distanceEmu: 38100, directionAngle: 2700000, alphaPct: 50,
    });
    // resolveColorNode now propagates <a:alpha> onto the color object
    expect(effects.shadow.color).toEqual({ space: 'srgb', hex: '000000', alpha: 50 });
  });

  test('inner shadow sets mode:inner', () => {
    const effects = resolveEffects(makeShadowSpPr({ mode: 'inner' }));
    expect(effects.shadow.mode).toBe('inner');
  });

  test('outer shadow takes priority over inner when both present', () => {
    const spPr = {
      'a:effectLst': {
        'a:outerShdw': { '@_blurRad': '50800', '@_dist': '38100', '@_dir': '0', 'a:srgbClr': { '@_val': 'FF0000' } },
        'a:innerShdw': { '@_blurRad': '25400', '@_dist': '19050', '@_dir': '0', 'a:srgbClr': { '@_val': '0000FF' } },
      },
    };
    const effects = resolveEffects(spPr);
    expect(effects.shadow.mode).toBe('outer');
    expect(effects.shadow.color.hex).toBe('FF0000');
  });

  test('alpha val 100000 → alphaPct 100', () => {
    expect(resolveEffects(makeShadowSpPr({ alphaVal: '100000' })).shadow.alphaPct).toBe(100);
  });

  test('alpha val 0 → alphaPct 0', () => {
    expect(resolveEffects(makeShadowSpPr({ alphaVal: '0' })).shadow.alphaPct).toBe(0);
  });

  test('no effectLst → undefined', () => {
    expect(resolveEffects({})).toBeUndefined();
  });

  test('null spPr → undefined', () => {
    expect(resolveEffects(null)).toBeUndefined();
  });

  test('effectLst with no recognized shadow key → undefined', () => {
    expect(resolveEffects({ 'a:effectLst': { 'a:glow': {} } })).toBeUndefined();
  });

  test('glow extracts color, radius, and alpha', () => {
    const spPr = {
      'a:effectLst': {
        'a:glow': { '@_rad': '139700', 'a:srgbClr': { '@_val': 'FFC000', 'a:alpha': { '@_val': '40000' } } },
      },
    };
    const effects = resolveEffects(spPr);
    expect(effects.glow).toEqual({
      color: { space: 'srgb', hex: 'FFC000', alpha: 40 },
      radiusEmu: 139700,
      alphaPct: 40,
    });
  });

  test('glow with unresolvable color → no glow effect', () => {
    const spPr = { 'a:effectLst': { 'a:glow': { '@_rad': '139700' } } };
    expect(resolveEffects(spPr)).toBeUndefined();
  });

  test('softEdge extracts radius', () => {
    const spPr = { 'a:effectLst': { 'a:softEdge': { '@_rad': '127000' } } };
    expect(resolveEffects(spPr).softEdge).toEqual({ radiusEmu: 127000 });
  });

  test('shadow, glow, and softEdge can coexist', () => {
    const spPr = {
      'a:effectLst': {
        'a:outerShdw': { '@_blurRad': '50800', '@_dist': '38100', '@_dir': '0', 'a:srgbClr': { '@_val': '000000' } },
        'a:glow': { '@_rad': '139700', 'a:srgbClr': { '@_val': 'FFC000' } },
        'a:softEdge': { '@_rad': '127000' },
      },
    };
    const effects = resolveEffects(spPr);
    expect(effects.shadow).toBeDefined();
    expect(effects.glow).toBeDefined();
    expect(effects.softEdge).toEqual({ radiusEmu: 127000 });
  });
});

describe('resolveColorNode — lumMod/lumOff baking', () => {
  const themeColors = { accent1: '#4472C4', dk1: '#000000' };

  test('schemeClr with lumMod/lumOff and themeColors bakes to srgb hex', () => {
    const node = {
      'a:schemeClr': { '@_val': 'accent1', 'a:lumMod': { '@_val': '60000' }, 'a:lumOff': { '@_val': '40000' } },
    };
    const result = resolveColorNode(node, themeColors);
    expect(result.space).toBe('srgb');
    expect(result.hex).toMatch(/^[0-9A-F]{6}$/);
  });

  test('schemeClr with lumMod but no themeColors falls back to theme ref', () => {
    const node = { 'a:schemeClr': { '@_val': 'accent1', 'a:lumMod': { '@_val': '60000' } } };
    expect(resolveColorNode(node)).toEqual({ space: 'theme', ref: 'accent1' });
  });

  test('schemeClr without lumMod/tint/shade stays a theme ref even with themeColors passed', () => {
    const node = { 'a:schemeClr': { '@_val': 'accent1' } };
    expect(resolveColorNode(node, themeColors)).toEqual({ space: 'theme', ref: 'accent1' });
  });

  test('tx1/bg1 aliases resolve through RAW_SCHEME_TO_THEME_SLOT to dk1/lt1', () => {
    const node = { 'a:schemeClr': { '@_val': 'tx1', 'a:shade': { '@_val': '50000' } } };
    const result = resolveColorNode(node, themeColors);
    expect(result).toEqual({ space: 'srgb', hex: '000000' });
  });

  test('unresolvable theme slot with modifiers falls back to theme ref', () => {
    const node = { 'a:schemeClr': { '@_val': 'accent2', 'a:lumMod': { '@_val': '60000' } } };
    expect(resolveColorNode(node, themeColors)).toEqual({ space: 'theme', ref: 'accent2' });
  });
});

describe('extractAdjustments', () => {
  test('single adj → [{name, value}]', () => {
    const pg = { 'a:avLst': { 'a:gd': { '@_name': 'adj', '@_fmla': 'val 16667' } } };
    expect(extractAdjustments(pg, [])).toEqual([{ name: 'adj', value: 16667 }]);
  });

  test('multiple adjs → array with all entries', () => {
    const pg = {
      'a:avLst': {
        'a:gd': [
          { '@_name': 'adj1', '@_fmla': 'val 10000' },
          { '@_name': 'adj2', '@_fmla': 'val 20000' },
        ],
      },
    };
    expect(extractAdjustments(pg, [])).toEqual([
      { name: 'adj1', value: 10000 },
      { name: 'adj2', value: 20000 },
    ]);
  });

  test('non-val formula pushes warning and is skipped', () => {
    const pg = {
      'a:avLst': {
        'a:gd': [
          { '@_name': 'adj1', '@_fmla': 'val 10000' },
          { '@_name': 'adj2', '@_fmla': '*/ adj1 2 3' },
        ],
      },
    };
    const w = [];
    const result = extractAdjustments(pg, w);
    expect(result).toHaveLength(1);
    expect(result[0].name).toBe('adj1');
    expect(w).toHaveLength(1);
    expect(w[0]).toContain('adj2');
  });

  test('negative value is preserved', () => {
    const pg = { 'a:avLst': { 'a:gd': { '@_name': 'adj', '@_fmla': 'val -5000' } } };
    expect(extractAdjustments(pg, [])[0].value).toBe(-5000);
  });

  test('no avLst → undefined', () => {
    expect(extractAdjustments({ 'a:avLst': undefined }, [])).toBeUndefined();
  });

  test('null prstGeom → undefined', () => {
    expect(extractAdjustments(null, [])).toBeUndefined();
  });

  test('empty avLst (no a:gd) → undefined', () => {
    expect(extractAdjustments({ 'a:avLst': {} }, [])).toBeUndefined();
  });
});

describe('extractCustomGeometry', () => {
  test('moveTo + close → commands with correct ops', () => {
    const spPr = {
      'a:custGeom': {
        'a:pathLst': {
          'a:path': {
            '@_w': '1524000', '@_h': '1524000',
            'a:moveTo': { 'a:pt': { '@_x': '762000', '@_y': '0' } },
            'a:close': {},
          },
        },
      },
    };
    const result = extractCustomGeometry(spPr);
    expect(result).toBeDefined();
    expect(result.w).toBe(1524000);
    expect(result.h).toBe(1524000);
    const cmds = result.paths[0].commands;
    const moveTo = cmds.find((c) => c.op === 'moveTo');
    expect(moveTo).toBeDefined();
    expect(moveTo.pts).toEqual([{ x: 762000, y: 0 }]);
    expect(cmds.find((c) => c.op === 'close')).toBeDefined();
  });

  test('lnTo captures its single point', () => {
    const spPr = {
      'a:custGeom': {
        'a:pathLst': {
          'a:path': {
            '@_w': '100', '@_h': '100',
            'a:moveTo': { 'a:pt': { '@_x': '0', '@_y': '0' } },
            'a:lnTo':   { 'a:pt': { '@_x': '100', '@_y': '100' } },
          },
        },
      },
    };
    const result = extractCustomGeometry(spPr);
    const lnTo = result.paths[0].commands.find((c) => c.op === 'lnTo');
    expect(lnTo.pts).toEqual([{ x: 100, y: 100 }]);
  });

  test('cubicBezTo captures 3 pts', () => {
    const spPr = {
      'a:custGeom': {
        'a:pathLst': {
          'a:path': {
            '@_w': '100', '@_h': '100',
            'a:moveTo': { 'a:pt': { '@_x': '0', '@_y': '0' } },
            'a:cubicBezTo': {
              'a:pt': [
                { '@_x': '25', '@_y': '0' },
                { '@_x': '75', '@_y': '100' },
                { '@_x': '100', '@_y': '100' },
              ],
            },
          },
        },
      },
    };
    const result = extractCustomGeometry(spPr);
    const bez = result.paths[0].commands.find((c) => c.op === 'cubicBezTo');
    expect(bez.pts).toHaveLength(3);
    expect(bez.pts[2]).toEqual({ x: 100, y: 100 });
  });

  test('arcTo produces command without pts property', () => {
    const spPr = {
      'a:custGeom': {
        'a:pathLst': {
          'a:path': {
            '@_w': '100', '@_h': '100',
            'a:arcTo': { '@_wR': '50', '@_hR': '50', '@_stAng': '0', '@_swAng': '5400000' },
          },
        },
      },
    };
    const result = extractCustomGeometry(spPr);
    const arc = result.paths[0].commands.find((c) => c.op === 'arcTo');
    expect(arc).toBeDefined();
    expect(arc.pts).toBeUndefined();
  });

  test('multiple paths are all collected', () => {
    const spPr = {
      'a:custGeom': {
        'a:pathLst': {
          'a:path': [
            { '@_w': '100', '@_h': '100', 'a:moveTo': { 'a:pt': { '@_x': '0', '@_y': '0' } } },
            { '@_w': '100', '@_h': '100', 'a:moveTo': { 'a:pt': { '@_x': '50', '@_y': '50' } } },
          ],
        },
      },
    };
    expect(extractCustomGeometry(spPr).paths).toHaveLength(2);
  });

  test('no custGeom → undefined', () => {
    expect(extractCustomGeometry({ 'a:prstGeom': { '@_prst': 'rect' } })).toBeUndefined();
  });

  test('null spPr → undefined', () => {
    expect(extractCustomGeometry(null)).toBeUndefined();
  });
});

// ===========================================================================
// 5. Schema — new type validation
// ===========================================================================

describe('schema — gradient fill', () => {
  test('accepts valid 2-stop linear gradient', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'gradient', kind: 'linear', angle: 5400000,
      stops: [
        { pos: 0,      color: { space: 'theme', ref: 'accent1' } },
        { pos: 100000, color: { space: 'srgb',  hex: 'FFFFFF'  } },
      ],
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('accepts radial gradient (no angle)', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'gradient', kind: 'radial',
      stops: [
        { pos: 0,      color: { space: 'srgb', hex: 'FF0000' } },
        { pos: 100000, color: { space: 'srgb', hex: '0000FF' } },
      ],
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects gradient with only 1 stop (minItems: 2)', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'gradient', kind: 'linear', angle: 0,
      stops: [{ pos: 0, color: { space: 'srgb', hex: 'FF0000' } }],
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects gradient missing kind', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'gradient',
      stops: [
        { pos: 0,      color: { space: 'srgb', hex: 'FF0000' } },
        { pos: 100000, color: { space: 'srgb', hex: '0000FF' } },
      ],
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects gradient stop with pos > 100000', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'gradient', kind: 'linear', angle: 0,
      stops: [
        { pos: 0,      color: { space: 'srgb', hex: 'FF0000' } },
        { pos: 200000, color: { space: 'srgb', hex: '0000FF' } },
      ],
    };
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — shadow effects', () => {
  test('accepts valid outer shadow', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].effects = {
      shadow: {
        mode: 'outer', color: { space: 'srgb', hex: '000000' },
        blurEmu: 50800, distanceEmu: 38100, directionAngle: 2700000, alphaPct: 50,
      },
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('accepts inner shadow', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].effects = {
      shadow: {
        mode: 'inner', color: { space: 'theme', ref: 'text1' },
        blurEmu: 0, distanceEmu: 0, directionAngle: 0, alphaPct: 75,
      },
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects shadow missing alphaPct', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].effects = {
      shadow: {
        mode: 'outer', color: { space: 'srgb', hex: '000000' },
        blurEmu: 50800, distanceEmu: 38100, directionAngle: 0,
      },
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects shadow with alphaPct > 100', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].effects = {
      shadow: {
        mode: 'outer', color: { space: 'srgb', hex: '000000' },
        blurEmu: 0, distanceEmu: 0, directionAngle: 0, alphaPct: 150,
      },
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects shadow with invalid mode string', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].effects = {
      shadow: {
        mode: 'blurred', color: { space: 'srgb', hex: '000000' },
        blurEmu: 0, distanceEmu: 0, directionAngle: 0, alphaPct: 50,
      },
    };
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — arrowhead stroke', () => {
  test('accepts solid stroke with headEnd and tailEnd', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].stroke = {
      type: 'solid', color: '#000000', width: 1.33, style: 'solid',
      headEnd: { type: 'triangle', width: 'med', length: 'med' },
      tailEnd:  { type: 'arrow',   width: 'lg',  length: 'lg'  },
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('accepts stroke without arrowheads', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].stroke = {
      type: 'solid', color: '#000000', width: 1.33, style: 'solid',
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects headEnd with type outside enum', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].stroke = {
      type: 'solid', color: '#000000', width: 1.33, style: 'solid',
      headEnd: { type: 'bigarrow' },
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects headEnd with width outside enum', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].stroke = {
      type: 'solid', color: '#000000', width: 1.33, style: 'solid',
      headEnd: { type: 'triangle', width: 'huge' },
    };
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — structured adjustments', () => {
  test('accepts adjustments as array of {name, value}', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].adjustments = [{ name: 'adj', value: 16667 }];
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects adjustments as plain object (old format)', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].adjustments = { adj: 16667 };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects adjustment item missing name', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].adjustments = [{ value: 16667 }];
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects adjustment item missing value', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].adjustments = [{ name: 'adj' }];
    expect(validate(doc).valid).toBe(false);
  });
});

describe('schema — customGeometry', () => {
  test('accepts valid triangle customGeometry', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].customGeometry = {
      w: 1524000, h: 1524000,
      paths: [{
        commands: [
          { op: 'moveTo', pts: [{ x: 0, y: 0 }] },
          { op: 'lnTo',   pts: [{ x: 100, y: 100 }] },
          { op: 'close' },
        ],
      }],
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('accepts close command with no pts field', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].customGeometry = {
      w: 100, h: 100,
      paths: [{ commands: [{ op: 'close' }] }],
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects command with unknown op', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].customGeometry = {
      w: 100, h: 100,
      paths: [{ commands: [{ op: 'bezierTo', pts: [{ x: 0, y: 0 }] }] }],
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects customGeometry missing w', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].customGeometry = {
      h: 100,
      paths: [{ commands: [{ op: 'close' }] }],
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects customGeometry missing paths', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].customGeometry = { w: 100, h: 100 };
    expect(validate(doc).valid).toBe(false);
  });
});

describe('parseAnimations — IR contract shape', () => {
  test('every Animation has spec fields and internal extensions', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '5' });
    const { animations } = parseAnimations(sld);
    const a = animations[0];
    // Spec fields
    expect(typeof a.id).toBe('string');
    expect(typeof a.sequence).toBe('number');
    expect(typeof a.effect).toBe('string');       // plain string per spec
    expect(typeof a.speed).toBe('number');        // seconds per spec
    expect(typeof a['effect-options']).toBe('object');
    // Internal extensions
    expect(typeof a.targetId).toBe('string');
    expect(['onClick', 'withPrevious', 'afterPrevious']).toContain(a.trigger);
    expect(typeof a['effect-detail'].class).toBe('string');
    expect(typeof a['effect-detail'].preset).toBe('string');
    expect(typeof a.timing.delayMs).toBe('number');
    expect(typeof a.timing.durationMs).toBe('number');
    expect(typeof a.supported).toBe('boolean');
  });

  test('speed is duration in seconds (durationMs / 1000)', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '3', dur: '750' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].speed).toBeCloseTo(0.75, 3);
    expect(animations[0].timing.durationMs).toBe(750);
  });
});
