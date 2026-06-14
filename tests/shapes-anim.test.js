'use strict';

const { validate, validateTargetIds } = require('../src/ir/validator');
const { parseShapes, extractXfrm, resolveColorNode, resolveFill, resolveStroke } = require('../src/parser/pptx/shapes');
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

  test('rejects a Shape missing z', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0].z;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('accepts any string as shape type (open enum — PPTX preset names preserved)', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].type = 'hexagon';
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });

  test('rejects a Shape with a non-string type', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].type = 42;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a position with missing w', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.shapes[0].position.w;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });
});

describe('schema — conditional: polyline/polygon/connector require points', () => {
  test('rejects a polyline without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a polygon without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polygon';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('rejects a connector without points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'connector';
    delete shape.points;
    const { valid } = validate(doc);
    expect(valid).toBe(false);
  });

  test('accepts a polyline with an empty points array', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline';
    shape.points = [];
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });

  test('accepts a polyline with populated points', () => {
    const doc = cloneFixture();
    const shape = doc.slideset.slides[0].contents.shapes[0];
    shape.type = 'polyline';
    shape.points = [{ x: 0, y: 0 }, { x: 914400, y: 685800 }];
    const { valid } = validate(doc);
    expect(valid).toBe(true);
  });
});

describe('schema — fill and stroke color types', () => {
  test('accepts theme-color fill', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid',
      color: { space: 'theme', ref: 'accent2' },
    };
    expect(validate(doc).valid).toBe(true);
  });

  test('rejects fill.color with unknown theme ref', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid',
      color: { space: 'theme', ref: 'neonPink' },
    };
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects fill.color with hex that is too short', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.shapes[0].fill = {
      type: 'solid',
      color: { space: 'srgb', hex: '4472' },
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
  test('rejects an Animation missing required fields', () => {
    const doc = cloneFixture();
    delete doc.slideset.slides[0].contents.animations[0].trigger;
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects an Animation with invalid trigger', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0].trigger = 'onHover';
    expect(validate(doc).valid).toBe(false);
  });

  test('rejects an Animation with invalid effect.class', () => {
    const doc = cloneFixture();
    doc.slideset.slides[0].contents.animations[0].effect.class = 'dance';
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

describe('extractXfrm', () => {
  test('reads x/y/w/h as raw EMU integers', () => {
    const spPr = {
      'a:xfrm': {
        'a:off': { '@_x': '914400', '@_y': '685800' },
        'a:ext': { '@_cx': '3200400', '@_cy': '1371600' },
      },
    };
    const { position } = extractXfrm(spPr);
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
    const { rotation } = extractXfrm(spPr);
    expect(rotation).toBe(5400000); // 90 degrees in PPTX units
  });

  test('missing xfrm returns zeroed position and rotation 0', () => {
    const { position, rotation } = extractXfrm(null);
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
    const { flipH, flipV } = extractXfrm(spPr);
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

  test('solidFill with srgbClr → {type:solid, color:{space:srgb,...}}', () => {
    const spPr = { 'a:solidFill': { 'a:srgbClr': { '@_val': 'FF0000' } } };
    expect(resolveFill(spPr)).toEqual({
      type: 'solid',
      color: { space: 'srgb', hex: 'FF0000' },
    });
  });

  test('solidFill with schemeClr → {type:solid, color:{space:theme,...}}', () => {
    const spPr = { 'a:solidFill': { 'a:schemeClr': { '@_val': 'accent2' } } };
    expect(resolveFill(spPr)).toEqual({
      type: 'solid',
      color: { space: 'theme', ref: 'accent2' },
    });
  });

  test('gradient fill (not yet supported) → {type:none}', () => {
    const spPr = { 'a:gradFill': { 'a:gsLst': {} } };
    expect(resolveFill(spPr)).toEqual({ type: 'none' });
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

  test('ln with solidFill and explicit width → {type:solid, color, widthEmu}', () => {
    const spPr = {
      'a:ln': {
        '@_w': '25400',
        'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
      },
    };
    expect(resolveStroke(spPr)).toEqual({
      type: 'solid',
      color: { space: 'srgb', hex: '000000' },
      widthEmu: 25400,
    });
  });

  test('ln with solidFill and no width defaults to 12700 EMU', () => {
    const spPr = {
      'a:ln': {
        'a:solidFill': { 'a:srgbClr': { '@_val': 'FFFFFF' } },
      },
    };
    const stroke = resolveStroke(spPr);
    expect(stroke.type).toBe('solid');
    expect(stroke.widthEmu).toBe(12700);
  });
});

describe('parseShapes — shape type dispatch', () => {
  test('rect preset maps to type:rect', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('rect');
  });

  test('roundRect preset maps fully (type:roundRect)', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'roundRect' }), null, []);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('roundRect');
  });

  test('ellipse preset maps to type:ellipse', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'ellipse' }), null, []);
    expect(shapes[0].type).toBe('ellipse');
  });

  test('line preset maps to type:line', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'line' }), null, []);
    expect(shapes[0].type).toBe('line');
  });

  test('rightArrow maps to type:arrow', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rightArrow' }), null, []);
    expect(shapes[0].type).toBe('arrow');
  });

  test('wedgeRectCallout maps to type:callout', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'wedgeRectCallout' }), null, []);
    expect(shapes[0].type).toBe('callout');
  });

  test('bentConnector3 maps to type:connector', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'bentConnector3' }), null, []);
    expect(shapes[0].type).toBe('connector');
  });
});

describe('parseShapes — unsupported preset preserves original PPTX name', () => {
  test('unrecognized preset keeps its PPTX name as type (not "unknown")', () => {
    const warnings = [];
    const shapes = parseShapes(buildSpTree({ prst: 'star5' }), null, warnings);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('star5');
  });

  test('unrecognized preset sets supported:false', () => {
    const warnings = [];
    const shapes = parseShapes(buildSpTree({ prst: 'star5' }), null, warnings);
    expect(shapes[0].supported).toBe(false);
  });

  test('unsupported preset pushes a warning containing the preset name', () => {
    const warnings = [];
    parseShapes(buildSpTree({ prst: 'star5' }), null, warnings);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('star5');
  });

  test('shape is NOT dropped for unsupported preset', () => {
    const warnings = [];
    const shapes = parseShapes(buildSpTree({ prst: 'cloudCallout2000' }), null, warnings);
    expect(shapes).toHaveLength(1);
    expect(shapes[0].type).toBe('cloudCallout2000');
  });

  test('recognized preset does NOT set supported:false', () => {
    const warnings = [];
    const shapes = parseShapes(buildSpTree({ prst: 'rect' }), null, warnings);
    expect(shapes[0].type).toBe('rect');
    expect(shapes[0].supported).toBeUndefined();
  });

  test('missing prst (custom geometry) emits type:unknown + supported:false + warning', () => {
    const spTree = buildSpTree({ prst: null });
    delete spTree['p:sp'][0]['p:spPr']['a:prstGeom'];
    const warnings = [];
    const shapes = parseShapes(spTree, null, warnings);
    expect(shapes[0].type).toBe('unknown');
    expect(shapes[0].supported).toBe(false);
    expect(warnings).toHaveLength(1);
  });
});

describe('parseShapes — geometry fields', () => {
  test('position is in EMU (not converted to pixels)', () => {
    const shapes = parseShapes(
      buildSpTree({ prst: 'rect', x: 914400, y: 685800, cx: 3200400, cy: 1371600 }),
      null, [],
    );
    expect(shapes[0].position).toEqual({ x: 914400, y: 685800, w: 3200400, h: 1371600 });
  });

  test('rotation stored as raw PPTX rot units', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect', rot: 3000000 }), null, []);
    expect(shapes[0].rotation).toBe(3000000);
  });

  test('flipH=true is preserved', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect', flipH: true }), null, []);
    expect(shapes[0].flipH).toBe(true);
  });

  test('roundRect with avLst adj emits adjustments object', () => {
    const spTree = buildSpTree({ prst: 'roundRect' });
    spTree['p:sp'][0]['p:spPr']['a:prstGeom']['a:avLst'] = {
      'a:gd': { '@_name': 'adj', '@_fmla': 'val 16667' },
    };
    const shapes = parseShapes(spTree, null, []);
    expect(shapes[0].adjustments).toBeDefined();
    expect(shapes[0].adjustments.adj).toBe(16667);
  });
});

describe('parseShapes — fill and stroke', () => {
  test('solidFill hex emits fill.color as srgb', () => {
    const shapes = parseShapes(
      buildSpTree({ prst: 'rect', solidFillHex: '4472C4' }),
      null, [],
    );
    expect(shapes[0].fill).toEqual({ type: 'solid', color: { space: 'srgb', hex: '4472C4' } });
  });

  test('solidFill scheme emits fill.color as theme ref', () => {
    const shapes = parseShapes(
      buildSpTree({ prst: 'rect', solidFillScheme: 'accent3' }),
      null, [],
    );
    expect(shapes[0].fill).toEqual({ type: 'solid', color: { space: 'theme', ref: 'accent3' } });
  });

  test('noFill emits fill.type:none', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect', noFill: true }), null, []);
    expect(shapes[0].fill).toEqual({ type: 'none' });
  });

  test('stroke with explicit width stores widthEmu', () => {
    const shapes = parseShapes(
      buildSpTree({ prst: 'rect', lineWidth: 25400 }),
      null, [],
    );
    expect(shapes[0].stroke.type).toBe('solid');
    expect(shapes[0].stroke.widthEmu).toBe(25400);
  });
});

describe('parseShapes — placeholder shapes are skipped', () => {
  test('placeholder p:sp is not included in shapes output', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect', hasPh: true }), null, []);
    expect(shapes).toHaveLength(0);
  });
});

describe('parseShapes — connector and points', () => {
  test('connector type emits an empty points array', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'bentConnector3' }), null, []);
    expect(Array.isArray(shapes[0].points)).toBe(true);
  });

  test('null/empty spTree returns empty array', () => {
    expect(parseShapes(null, null, [])).toEqual([]);
    expect(parseShapes({}, null, [])).toEqual([]);
  });
});

describe('parseShapes — z placeholder', () => {
  test('z is 0 (slide.js assigns final value)', () => {
    const shapes = parseShapes(buildSpTree({ prst: 'rect' }), null, []);
    expect(shapes[0].z).toBe(0);
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
    expect(animations[0].effect.class).toBe('entrance');
  });

  test('presetID 21 maps to preset "fade"', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '3' });
    const { animations } = parseAnimations(sld);
    expect(animations[0].effect.preset).toBe('fade');
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

describe('parseAnimations — IR contract shape', () => {
  test('every Animation has required fields', () => {
    const sld = buildSlideWithAnim({ presetClass: 'entr', presetID: '21', spid: '5' });
    const { animations } = parseAnimations(sld);
    const a = animations[0];
    expect(typeof a.id).toBe('string');
    expect(typeof a.targetId).toBe('string');
    expect(['onClick', 'withPrevious', 'afterPrevious']).toContain(a.trigger);
    expect(typeof a.order).toBe('number');
    expect(typeof a.effect.class).toBe('string');
    expect(typeof a.effect.preset).toBe('string');
    expect(typeof a.timing.delayMs).toBe('number');
    expect(typeof a.timing.durationMs).toBe('number');
    expect(typeof a.supported).toBe('boolean');
  });
});
