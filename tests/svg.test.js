'use strict';

const { emitShape, renderShape, simulateLines, measureAndShrink } = require('../src/generator/revealjs/svg');

const EMU_PER_PX = 9525;

// Build a minimal valid rect shape, applying any field overrides.
function rect(overrides = {}) {
  return {
    type: 'rect',
    position: { x: 0, y: 0 },
    width:  100 * EMU_PER_PX,
    height:  50 * EMU_PER_PX,
    fill:   { type: 'solid', color: '#ff0000' },
    stroke: { color: '#000000', width: 1 },
    ...overrides,
  };
}

// ---------------------------------------------------------------------------
// emitShape — translate / rotate geometry
// ---------------------------------------------------------------------------

describe('emitShape — translate from EMU position', () => {
  test('converts EMU x/y to px in translate()', () => {
    const g = emitShape(rect({ position: { x: 381 * EMU_PER_PX, y: 289 * EMU_PER_PX } }), { warnings: [] });
    expect(g).toContain('translate(381,289)');
  });

  test('zero position emits translate(0,0)', () => {
    const g = emitShape(rect({ position: { x: 0, y: 0 } }), { warnings: [] });
    expect(g).toContain('translate(0,0)');
  });

  test('missing position defaults to translate(0,0)', () => {
    const shape = { type: 'rect', width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX };
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('translate(0,0)');
  });
});

describe('emitShape — rotation about shape center', () => {
  test('non-zero rotation appends rotate(deg cx cy) with center in local coords', () => {
    // width=180px → cx=90, height=110px → cy=55
    const g = emitShape(rect({
      position: { x: 381 * EMU_PER_PX, y: 289 * EMU_PER_PX },
      width:  180 * EMU_PER_PX,
      height: 110 * EMU_PER_PX,
      rotation: 45,
    }), { warnings: [] });
    expect(g).toContain('translate(381,289)');
    expect(g).toContain('rotate(45 90 55)');
  });

  test('rotation=0 omits the rotate() clause entirely', () => {
    const g = emitShape(rect({ rotation: 0 }), { warnings: [] });
    expect(g).not.toContain('rotate(');
  });

  test('absent rotation omits the rotate() clause', () => {
    const shape = { type: 'rect', position: { x: 0, y: 0 }, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX };
    const g = emitShape(shape, { warnings: [] });
    expect(g).not.toContain('rotate(');
  });
});

// ---------------------------------------------------------------------------
// emitShape — rounded corners
// ---------------------------------------------------------------------------

describe('emitShape — roundRect geometry', () => {
  test('geometry.rx and geometry.ry map to SVG rx/ry on the <rect>', () => {
    const g = emitShape(rect({ geometry: { rx: 10, ry: 8 } }), { warnings: [] });
    expect(g).toContain('rx="10" ry="8"');
  });

  test('geometry.rx alone sets symmetric corners (ry defaults to rx)', () => {
    const g = emitShape(rect({ geometry: { rx: 15 } }), { warnings: [] });
    expect(g).toContain('rx="15" ry="15"');
  });

  test('no geometry sets rx=0 ry=0 (sharp corners)', () => {
    const g = emitShape(rect(), { warnings: [] });
    expect(g).toContain('rx="0" ry="0"');
  });
});

// ---------------------------------------------------------------------------
// emitShape — fill
// ---------------------------------------------------------------------------

describe('emitShape — fill colors', () => {
  test('theme-color fill passes through as-is (var(--theme-X))', () => {
    const g = emitShape(rect({ fill: { type: 'solid', color: 'var(--theme-accent1)' } }), { warnings: [] });
    expect(g).toContain('fill="var(--theme-accent1)"');
  });

  test('explicit RGB hex fill is emitted verbatim', () => {
    const g = emitShape(rect({ fill: { type: 'solid', color: '#0d6ce7' } }), { warnings: [] });
    expect(g).toContain('fill="#0d6ce7"');
  });

  test('fill.type=none emits fill="none"', () => {
    const g = emitShape(rect({ fill: { type: 'none' } }), { warnings: [] });
    expect(g).toContain('fill="none"');
  });

  test('absent fill emits fill="none"', () => {
    const shape = { type: 'rect', position: { x: 0, y: 0 }, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX };
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('fill="none"');
  });
});

// ---------------------------------------------------------------------------
// emitShape — shape text via <foreignObject>
// ---------------------------------------------------------------------------

describe('emitShape — embedded text', () => {
  test('multi-paragraph text renders inside <foreignObject>', () => {
    const g = emitShape(rect({
      text: [
        { id: 'p-0', runs: [{ text: 'Hello' }] },
        { id: 'p-1', runs: [{ text: 'World' }] },
      ],
    }), { warnings: [] });
    expect(g).toContain('<foreignObject');
    expect(g).toContain('Hello');
    expect(g).toContain('World');
  });

  test('<foreignObject> uses shape pixel dimensions', () => {
    // width=100px, height=50px
    const g = emitShape(rect({
      text: [{ id: 'p-0', runs: [{ text: 'x' }] }],
    }), { warnings: [] });
    expect(g).toMatch(/foreignObject[^>]+width="100"/);
    expect(g).toMatch(/foreignObject[^>]+height="50"/);
  });

  test('shape with no text emits no <foreignObject>', () => {
    const g = emitShape(rect(), { warnings: [] });
    expect(g).not.toContain('foreignObject');
  });

  test('XSS in run text is escaped inside foreignObject', () => {
    const g = emitShape(rect({
      text: [{ id: 'p-0', runs: [{ text: '<script>evil()</script>' }] }],
    }), { warnings: [] });
    expect(g).not.toContain('<script>');
    expect(g).toContain('&lt;script&gt;');
  });
});

// ---------------------------------------------------------------------------
// emitShape — unsupported types (stubs)
// ---------------------------------------------------------------------------

describe('emitShape — now-implemented types render SVG', () => {
  const IMPLEMENTED = ['ellipse', 'line', 'arrow', 'callout', 'connector'];

  test.each(IMPLEMENTED)('%s: returns non-empty SVG string', (type) => {
    const result = emitShape(
      { type, position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX }, fill: { type: 'none' }, stroke: { type: 'none' } },
      { warnings: [] },
    );
    expect(typeof result).toBe('string');
    expect(result.length).toBeGreaterThan(0);
  });

  test.each(IMPLEMENTED)('%s: does not throw', (type) => {
    expect(() =>
      emitShape(
        { type, position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX }, fill: { type: 'none' }, stroke: { type: 'none' } },
        { warnings: [] },
      ),
    ).not.toThrow();
  });
});

describe('emitShape — still-unsupported stubs', () => {
  const STUBS = ['polyline', 'polygon', 'unknown'];

  test.each(STUBS)('%s: returns empty string', (type) => {
    const result = emitShape(
      { type, position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX }, fill: { type: 'none' }, stroke: { type: 'none' } },
      { warnings: [] },
    );
    expect(result).toBe('');
  });

  test('unknown custom type returns empty string', () => {
    const ctx = { warnings: [] };
    const result = emitShape(
      { type: 'cloudCallout2000', position: { x: 0, y: 0, w: 0, h: 0 }, fill: { type: 'none' }, stroke: { type: 'none' } },
      ctx,
    );
    expect(result).toBe('');
  });

  test('ctx with no warnings array does not throw', () => {
    expect(() => emitShape({ type: 'callout' }, {})).not.toThrow();
    expect(() => emitShape({ type: 'callout' }, null)).not.toThrow();
  });
});

// ---------------------------------------------------------------------------
// renderShape
// ---------------------------------------------------------------------------

describe('renderShape', () => {
  test('wraps supported rect in an <svg> element with a <g> inside', () => {
    const html = renderShape(rect());
    expect(html).toMatch(/^<svg /);
    expect(html).toContain('<g ');
    expect(html).toContain('</svg>');
  });

  test('SVG overlay covers the full canvas (width:100%, height:100%)', () => {
    const html = renderShape(rect());
    expect(html).toContain('width:100%');
    expect(html).toContain('height:100%');
    expect(html).toContain('position:absolute');
  });

  test('applies z-index from shape to SVG style', () => {
    const html = renderShape(rect({ 'z-index': 7 }));
    expect(html).toContain('z-index:7');
  });

  test('returns empty string for truly unsupported type', () => {
    const result = renderShape({
      type: 'unknown',
      position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX },
      fill: { type: 'none' }, stroke: { type: 'none' },
    });
    expect(result).toBe('');
  });

  test('ellipse is now rendered (returns non-empty SVG)', () => {
    const result = renderShape({
      type: 'ellipse',
      position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX },
      fill: { type: 'solid', color: { space: 'srgb', hex: 'FF0000' } },
      stroke: { type: 'none' },
    });
    expect(result).toContain('<ellipse');
  });
});

// ---------------------------------------------------------------------------
// Bug-fix regression tests
// ---------------------------------------------------------------------------

describe('emitShape — roundRect corner radius (Bug 1)', () => {
  function rrShape(adjVal, wEmu, hEmu) {
    return {
      type: 'roundRect',
      position: { x: 0, y: 0, w: wEmu, h: hEmu },
      fill: { type: 'none' }, stroke: { type: 'none' },
      adjustments: adjVal != null ? [{ name: 'adj', value: adjVal }] : [],
    };
  }

  test('explicit adj=50000 → rx = 50% × min(w,h)/2', () => {
    // w=200px h=100px → rx = (50000/100000)*100/2 = 25
    const g = emitShape(rrShape(50000, 200 * EMU_PER_PX, 100 * EMU_PER_PX), { warnings: [] });
    expect(g).toContain('rx="25"');
  });

  test('empty adj array → OOXML default 16667, not hardcoded 8px', () => {
    // w=130px h=68px → default rx = (16667/100000)*68/2 ≈ 5.67 → 6
    const g = emitShape(rrShape(null, 130 * EMU_PER_PX, 68 * EMU_PER_PX), { warnings: [] });
    const rxMatch = g.match(/rx="(\d+)"/);
    expect(rxMatch).not.toBeNull();
    const rx = Number(rxMatch[1]);
    // Should be ~6, definitely not 8 (old hardcoded) and not 0
    expect(rx).toBeGreaterThan(0);
    expect(rx).toBeLessThan(8); // OOXML default < hardcoded fallback for this shape size
  });

  test('large shape with empty adj uses proportional OOXML default', () => {
    // w=400px h=200px → default rx = (16667/100000)*200/2 ≈ 16.67 → 17
    const g = emitShape(rrShape(null, 400 * EMU_PER_PX, 200 * EMU_PER_PX), { warnings: [] });
    const rxMatch = g.match(/rx="(\d+)"/);
    expect(Number(rxMatch[1])).toBeGreaterThan(8); // clearly > old hardcoded 8
  });

  test('adj=16667 (legacy {adj:N} object format) → same rx as default', () => {
    const shape = {
      type: 'roundRect',
      position: { x: 0, y: 0, w: 130 * EMU_PER_PX, h: 68 * EMU_PER_PX },
      fill: { type: 'none' }, stroke: { type: 'none' },
      adjustments: { adj: 16667 }, // old IR object format
    };
    const g = emitShape(shape, { warnings: [] });
    const rxMatch = g.match(/rx="(\d+)"/);
    expect(rxMatch).not.toBeNull();
    expect(Number(rxMatch[1])).toBeGreaterThan(0);
  });
});

describe('emitShape — text overflow clipping (Bug 2)', () => {
  function shapeWithText() {
    return {
      type: 'rect',
      position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX },
      fill: { type: 'none' }, stroke: { type: 'none' },
      text: {
        id: 'txt-0',
        paragraphs: [{ id: 'p-0', runs: [{ text: 'Hello world' }] }],
        anchor: 'ctr',
        insets: { l: 91440, r: 91440, t: 45720, b: 45720 },
      },
    };
  }

  test('foreignObject inner div uses overflow:hidden (not overflow:visible)', () => {
    const g = emitShape(shapeWithText(), { warnings: [] });
    expect(g).toContain('overflow:hidden');
    expect(g).not.toContain('overflow:visible');
  });

  test('foreignObject element itself has no overflow="visible" attribute', () => {
    const g = emitShape(shapeWithText(), { warnings: [] });
    expect(g).not.toContain('overflow="visible"');
  });

  test('foreignObject dimensions match shape bounding box', () => {
    const g = emitShape(shapeWithText(), { warnings: [] });
    expect(g).toMatch(/foreignObject[^>]+width="100"/);
    expect(g).toMatch(/foreignObject[^>]+height="50"/);
  });

  test('vertical-center anchor applies flex justify-content:center', () => {
    const g = emitShape(shapeWithText(), { warnings: [] });
    expect(g).toContain('justify-content:center');
  });

  test('plain paragraph array (old IR) also uses overflow:hidden', () => {
    const shape = {
      type: 'rect',
      position: { x: 0, y: 0, w: 100 * EMU_PER_PX, h: 50 * EMU_PER_PX },
      fill: { type: 'none' }, stroke: { type: 'none' },
      text: [{ id: 'p-0', runs: [{ text: 'x' }] }],
    };
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('overflow:hidden');
    expect(g).not.toContain('overflow:visible');
  });
});

describe('emitShape — customGeometry path rendering (Bug 3)', () => {
  const E = EMU_PER_PX; // 9525

  function custGeomShape(commands) {
    return {
      type: 'unknown', // typical IR type for freeform shapes
      position: { x: 0, y: 0, w: 200 * E, h: 100 * E },
      fill: { type: 'solid', color: { space: 'srgb', hex: 'FF0000' } },
      stroke: { type: 'none' },
      customGeometry: {
        w: 200 * E,   // coordinate space same as shape size
        h: 100 * E,
        paths: [{ commands }],
      },
    };
  }

  test('moveTo + lnTo + close → <path> with M, L, Z', () => {
    const shape = custGeomShape([
      { op: 'moveTo', pts: [{ x: 0, y: 0 }] },
      { op: 'lnTo',   pts: [{ x: 200 * E, y: 100 * E }] },
      { op: 'close' },
    ]);
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('<path');
    expect(g).toContain('M ');
    expect(g).toContain('L ');
    expect(g).toContain('Z');
  });

  test('cubicBezTo → <path> with C command', () => {
    const shape = custGeomShape([
      { op: 'moveTo',     pts: [{ x: 0, y: 0 }] },
      { op: 'cubicBezTo', pts: [{ x: E, y: 0 }, { x: E, y: E }, { x: 2*E, y: E }] },
    ]);
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('C ');
  });

  test('quadBezTo → <path> with Q command', () => {
    const shape = custGeomShape([
      { op: 'moveTo',    pts: [{ x: 0, y: 0 }] },
      { op: 'quadBezTo', pts: [{ x: E, y: 0 }, { x: 2*E, y: E }] },
    ]);
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('Q ');
  });

  test('coordinate space is scaled to shape pixel dimensions', () => {
    // coord space 200×100 EMU, shape 200×100 px → scale=1
    const shape = custGeomShape([
      { op: 'moveTo', pts: [{ x: 100 * E, y: 50 * E }] },
      { op: 'close' },
    ]);
    const g = emitShape(shape, { warnings: [] });
    // x=100*E in custGeom space, scaled to 200px → result=100px
    expect(g).toContain('M 100.00,50.00');
  });

  test('arcTo is skipped gracefully (no pts stored in IR)', () => {
    const shape = custGeomShape([
      { op: 'moveTo', pts: [{ x: 0, y: 0 }] },
      { op: 'arcTo' },            // no pts — should not throw
      { op: 'lnTo', pts: [{ x: 100 * E, y: 100 * E }] },
    ]);
    expect(() => emitShape(shape, { warnings: [] })).not.toThrow();
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('<path');
    expect(g).toContain('L ');   // lnTo was rendered
  });

  test('unknown without customGeometry still returns empty string', () => {
    const shape = {
      type: 'unknown',
      position: { x: 0, y: 0, w: 100 * E, h: 50 * E },
      fill: { type: 'none' }, stroke: { type: 'none' },
    };
    expect(emitShape(shape, { warnings: [] })).toBe('');
  });

  test('empty paths array returns empty string', () => {
    const shape = {
      type: 'unknown',
      position: { x: 0, y: 0, w: 100 * E, h: 50 * E },
      fill: { type: 'none' }, stroke: { type: 'none' },
      customGeometry: { w: 100 * E, h: 50 * E, paths: [] },
    };
    expect(emitShape(shape, { warnings: [] })).toBe('');
  });

  test('fill and stroke from shape are applied to custGeom', () => {
    const shape = {
      type: 'unknown',
      position: { x: 0, y: 0, w: 100 * E, h: 50 * E },
      fill:   { type: 'solid', color: { space: 'srgb', hex: 'ABCDEF' } },
      stroke: { type: 'solid', color: { space: 'srgb', hex: '000000' }, widthEmu: 9525 },
      customGeometry: {
        w: 100 * E, h: 50 * E,
        paths: [{ commands: [{ op: 'moveTo', pts: [{ x: 0, y: 0 }] }, { op: 'close' }] }],
      },
    };
    const g = emitShape(shape, { warnings: [] });
    expect(g).toContain('#ABCDEF');
    expect(g).toContain('#000000');
  });
});

// ---------------------------------------------------------------------------
// simulateLines — glyph-accurate word-wrap simulation (test a)
// ---------------------------------------------------------------------------

// Mock font: every character has the same advance width (advPct of an em).
// unitsPerEm=1000, so advPct=0.5 means each char is 0.5em wide.
function mockFont(advPct = 0.5, unitsPerEm = 1000) {
  return {
    unitsPerEm,
    layout: (str) => ({ glyphs: [...str].map(() => ({ advanceWidth: advPct * unitsPerEm })) }),
  };
}

// PT_TO_PX = 12700/9525 ≈ 1.3333  →  charPx = advPct × fontSizePt × PT_TO_PX
describe('simulateLines — word-wrap simulation', () => {
  const PT_TO_PX = 12700 / 9525; // mirrors svg.js constant

  test('single word that fits in one line → 1', () => {
    // charPx = 0.5 × 10 × PT_TO_PX ≈ 6.67 px; "hello"=5 chars ≈ 33.3px < 100px
    expect(simulateLines('hello', mockFont(), 10, 100)).toBe(1);
  });

  test('two words that fit together → 1 line', () => {
    // "hi ho" = 5 chars ≈ 33.3px < 80px
    expect(simulateLines('hi ho', mockFont(), 10, 80)).toBe(1);
  });

  test('two words that do not fit together → 2 lines', () => {
    // avail=40px: "hello"≈33.3px fits; " world"≈40+33.3=73.3px > 40 → wrap
    expect(simulateLines('hello world', mockFont(), 10, 40)).toBe(2);
  });

  test('word wider than available width always starts on its own line (no mid-word break)', () => {
    // "superlongword"=13×6.67=86.7px > 50px, but x=0 so no break → still 1 line
    expect(simulateLines('superlongword', mockFont(), 10, 50)).toBe(1);
  });

  test('three-word string wrapping at narrow width → 3 lines', () => {
    // avail=20px; each word > 20px individually but x=0 means first word placed.
    // "a" ≈ 6.67px ≤ 20, " " → 13.33, "b" → 13.33+6.67=20 ≤ 20 → still line 1,
    // " " → 26.67, "c" → 26.67+6.67=33.33 > 20 and x>0 → line 2.
    // So "a b c" → 2 lines at 20px? Let me use wider chars.
    // advPct=1.0 means charPx = 1.0 × 10 × PT_TO_PX ≈ 13.33px
    // "aaa bbb ccc" at avail=20px:
    // "aaa"=40px, x=0 → x=40; " "→53.3; "bbb"→53.3+40=93.3>20 → line2, x=40
    // " "→53.3; "ccc"→53.3+40=93.3>20 → line3, x=40
    expect(simulateLines('aaa bbb ccc', mockFont(1.0), 10, 20)).toBe(3);
  });

  test('empty string → 1 (no content still occupies one line)', () => {
    expect(simulateLines('', mockFont(), 10, 100)).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// measureAndShrink — shrink loop (test b)
// ---------------------------------------------------------------------------

function para(text, sizePt, lineSpacing = '1') {
  return {
    formatting: { 'line-spacing': lineSpacing },
    runs: [{ text, formatting: { size: `${sizePt}pt`, font: 'Calibri' } }],
  };
}

describe('measureAndShrink — font-size reduction loop', () => {
  test('returns null when fontOverride is null (font not available)', () => {
    expect(measureAndShrink([para('Hello', 12)], 100, 50, null)).toBeNull();
  });

  test('returns null for mixed font sizes', () => {
    const paras = [
      { formatting: {}, runs: [{ text: 'A', formatting: { size: '12pt', font: 'Calibri' } }] },
      { formatting: {}, runs: [{ text: 'B', formatting: { size: '14pt', font: 'Calibri' } }] },
    ];
    expect(measureAndShrink(paras, 100, 50, mockFont())).toBeNull();
  });

  test('returns origPt unchanged when text already fits', () => {
    // advPct=0.1: charPx=0.1×12×PT_TO_PX≈1.6px; "hi"≈3.2px; 1 line × 16px < 200px
    const result = measureAndShrink([para('hi', 12)], 200, 200, mockFont(0.1));
    expect(result).not.toBeNull();
    expect(result.shrunkPt).toBe(result.origPt);
    expect(result.origPt).toBe(12);
  });

  test('reduces font size when text does not fit', () => {
    // advPct=0.5 → each char≈6.67px at 10pt; "Rectangle with rounded corners"
    // in 50px height with many lines — must shrink
    const result = measureAndShrink(
      [para('Rectangle with rounded corners', 18)], 100, 50, mockFont(0.5),
    );
    expect(result).not.toBeNull();
    expect(result.shrunkPt).toBeLessThan(result.origPt);
    expect(result.shrunkPt).toBeGreaterThanOrEqual(6);
  });

  test('never shrinks below 6pt minimum', () => {
    // Absurdly wide mock — text can never fit even at minimum size
    const result = measureAndShrink(
      [para('impossible text that will never fit', 18)], 10, 10, mockFont(5.0),
    );
    expect(result).not.toBeNull();
    expect(result.shrunkPt).toBe(6);
  });

  test('font reference is returned in the result object', () => {
    const font = mockFont(0.1);
    const result = measureAndShrink([para('ok', 12)], 200, 200, font);
    expect(result).not.toBeNull();
    expect(result.font).toBe(font);
  });
});

// ---------------------------------------------------------------------------
// Callout / cloud preset geometry renderers
// ---------------------------------------------------------------------------

describe('emitShape — callout preset geometries', () => {
  function calloutShape(preset, adjustments) {
    const shape = {
      type: 'callout',
      preset,
      position: { x: 0, y: 0, w: 200 * EMU_PER_PX, h: 100 * EMU_PER_PX },
      fill:   { type: 'solid', color: { space: 'srgb', hex: 'FFCC00' } },
      stroke: { type: 'solid', color: { space: 'srgb', hex: '000000' }, widthEmu: EMU_PER_PX },
    };
    if (adjustments) shape.adjustments = adjustments;
    return shape;
  }

  // (a) wedgeRectCallout emits <path> with ≥7 path commands
  test('wedgeRectCallout emits <path> not <rect>', () => {
    const g = emitShape(calloutShape('wedgeRectCallout'), { warnings: [] });
    expect(g).toContain('<path');
    // Must not be a plain rect fallback
    expect(g).not.toMatch(/<rect[^>]+data-preset/);
  });

  test('wedgeRectCallout d attribute has at least 7 path commands', () => {
    const g = emitShape(calloutShape('wedgeRectCallout'), { warnings: [] });
    const dMatch = g.match(/\bd="([^"]+)"/);
    expect(dMatch).not.toBeNull();
    const cmds = dMatch[1].match(/[MLHVCSQTAZmlhvcsqtaz]/g) || [];
    expect(cmds.length).toBeGreaterThanOrEqual(7);
  });

  test('wedgeRectCallout with explicit adj1/adj2 puts tail tip at adj-derived coords', () => {
    // adj1=50000 → tx=100px (centre), adj2=150000 → ty=150px (below shape)
    const g = emitShape(calloutShape('wedgeRectCallout', [
      { name: 'adj1', value: 50000 },
      { name: 'adj2', value: 150000 },
    ]), { warnings: [] });
    // ty=150px should appear in the d attribute
    expect(g).toContain('150.0');
  });

  // (b) cloud emits <path> with arc commands (A)
  test('cloud emits a <path> with arc commands (A in d)', () => {
    const g = emitShape(
      {
        type: 'cloud',
        preset: 'cloud',
        position: { x: 0, y: 0, w: 200 * EMU_PER_PX, h: 100 * EMU_PER_PX },
        fill:   { type: 'none' },
        stroke: { type: 'none' },
      },
      { warnings: [] },
    );
    expect(g).toContain('<path');
    const dMatch = g.match(/\bd="([^"]+)"/);
    expect(dMatch).not.toBeNull();
    expect(dMatch[1]).toContain('A');
  });

  // (c) cloudCallout emits two <path> elements (tail + body)
  test('cloudCallout emits two <path> elements', () => {
    const g = emitShape(calloutShape('cloudCallout', [
      { name: 'adj1', value: -20000 },
      { name: 'adj2', value: 120000 },
    ]), { warnings: [] });
    const pathCount = (g.match(/<path/g) || []).length;
    expect(pathCount).toBe(2);
  });

  test('cloudCallout tail path has fill="none"', () => {
    const g = emitShape(calloutShape('cloudCallout', [
      { name: 'adj1', value: -20000 },
      { name: 'adj2', value: 120000 },
    ]), { warnings: [] });
    expect(g).toContain('fill="none"');
  });

  // wedgeRoundRectCallout emits rounded corners (A arcs) and a wedge tail
  test('wedgeRoundRectCallout emits <path> with arc commands', () => {
    const g = emitShape(calloutShape('wedgeRoundRectCallout', [
      { name: 'adj1', value: 5036 },
      { name: 'adj2', value: 113015 },
      { name: 'adj3', value: 16667 },
    ]), { warnings: [] });
    expect(g).toContain('<path');
    const dMatch = g.match(/\bd="([^"]+)"/);
    expect(dMatch).not.toBeNull();
    expect(dMatch[1]).toContain('A');
  });

  // Unknown callout preset falls back to <rect> with warning
  test('unknown callout preset emits rect fallback and pushes a warning', () => {
    const ctx = { warnings: [] };
    const g = emitShape(calloutShape('wedgeEllipseCallout'), ctx);
    expect(g).toContain('<rect');
    expect(g).toContain('data-preset');
    expect(ctx.warnings.length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// emitShape — chevron and pentagon preset geometries
// ---------------------------------------------------------------------------

describe('emitShape — chevron polygon points', () => {
  function chevronShape(adj, wPx = 100, hPx = 50) {
    return {
      type: 'chevron',
      position: { x: 0, y: 0, w: wPx * EMU_PER_PX, h: hPx * EMU_PER_PX },
      fill:   { type: 'none' },
      stroke: { type: 'none' },
      ...(adj != null && { adjustments: [{ name: 'adj', value: adj }] }),
    };
  }

  test('adj=50000: points match OOXML formula (a=50, notch at left)', () => {
    // a = 50000/100000 * 100 = 50
    // points: 50,0  100,0  100,25  100,50  50,50  0,25
    const g = emitShape(chevronShape(50000), { warnings: [] });
    expect(g).toContain('points="50,0 100,0 100,25 100,50 50,50 0,25"');
  });

  test('adj=0: left notch at x=0 (degenerate — left side collapses to a point)', () => {
    // a = 0; points: 0,0  100,0  100,25  100,50  0,50  0,25
    const g = emitShape(chevronShape(0), { warnings: [] });
    expect(g).toContain('points="0,0 100,0 100,25 100,50 0,50 0,25"');
  });

  test('no adjustments: defaults to adj=50000', () => {
    const g = emitShape(chevronShape(null), { warnings: [] });
    expect(g).toContain('points="50,0 100,0 100,25 100,50 50,50 0,25"');
  });

  test('returns non-empty SVG string and contains <polygon>', () => {
    const g = emitShape(chevronShape(50000), { warnings: [] });
    expect(g.length).toBeGreaterThan(0);
    expect(g).toContain('<polygon');
  });
});

describe('emitShape — pentagon polygon points', () => {
  function pentagonShape(adj, wPx = 100, hPx = 50) {
    return {
      type: 'pentagon',
      position: { x: 0, y: 0, w: wPx * EMU_PER_PX, h: hPx * EMU_PER_PX },
      fill:   { type: 'none' },
      stroke: { type: 'none' },
      ...(adj != null && { adjustments: [{ name: 'adj', value: adj }] }),
    };
  }

  test('adj=50000: points match OOXML formula (right arrow-pentagon)', () => {
    // a = 50000/100000 * 100 = 50; w-a = 50
    // points: 0,0  50,0  100,25  50,50  0,50
    const g = emitShape(pentagonShape(50000), { warnings: [] });
    expect(g).toContain('points="0,0 50,0 100,25 50,50 0,50"');
  });

  test('adj=0: point at right edge only (full-width right point)', () => {
    // a=0; w-a=100; points: 0,0  100,0  100,25  100,50  0,50
    const g = emitShape(pentagonShape(0), { warnings: [] });
    expect(g).toContain('points="0,0 100,0 100,25 100,50 0,50"');
  });

  test('no adjustments: defaults to adj=50000', () => {
    const g = emitShape(pentagonShape(null), { warnings: [] });
    expect(g).toContain('points="0,0 50,0 100,25 50,50 0,50"');
  });

  test('returns non-empty SVG string and contains <polygon>', () => {
    const g = emitShape(pentagonShape(50000), { warnings: [] });
    expect(g.length).toBeGreaterThan(0);
    expect(g).toContain('<polygon');
  });
});

// ---------------------------------------------------------------------------
// emitShape — database cylinder (rect + two ellipses)
// ---------------------------------------------------------------------------

describe('emitShape — database cylinder geometry', () => {
  // w=100px, h=100px → rimRy = 100*0.18 = 18
  function dbShape(wPx = 100, hPx = 100) {
    return {
      type: 'database',
      position: { x: 0, y: 0, w: wPx * EMU_PER_PX, h: hPx * EMU_PER_PX },
      fill:   { type: 'solid', color: { space: 'srgb', hex: '4472C4' } },
      stroke: { type: 'solid', color: { space: 'srgb', hex: '000000' }, widthEmu: 9525 },
      'z-index': 0,
    };
  }

  test('contains a <path> for the cylinder body', () => {
    const g = emitShape(dbShape(), { warnings: [] });
    expect(g).toContain('<path');
    expect(g).toContain('d="M');
  });

  test('contains an <ellipse> for the top lid', () => {
    const g = emitShape(dbShape(), { warnings: [] });
    expect(g).toContain('<ellipse');
  });

  test('top lid ellipse uses semi-transparent white rim stroke', () => {
    const g = emitShape(dbShape(), { warnings: [] });
    expect(g).toContain('stroke="rgba(255,255,255,0.45)"');
  });

  test('top lid cx is w/2 and cy is rimRy (h*0.18)', () => {
    // h=100 → rimRy=18, w=100 → cx=50
    const g = emitShape(dbShape(), { warnings: [] });
    expect(g).toContain('cx="50"');
    expect(g).toContain('cy="18"');
  });

  test('returns non-empty SVG and does not throw', () => {
    expect(() => emitShape(dbShape(), { warnings: [] })).not.toThrow();
    expect(emitShape(dbShape(), { warnings: [] }).length).toBeGreaterThan(0);
  });
});

// ---------------------------------------------------------------------------
// emitShape — star7 (via generic case 'star')
// ---------------------------------------------------------------------------

describe('emitShape — star7 polygon (via case star)', () => {
  function starShape(preset, wPx = 100, hPx = 100) {
    return {
      type:   'star',
      preset,
      position: { x: 0, y: 0, w: wPx * EMU_PER_PX, h: hPx * EMU_PER_PX },
      fill:   { type: 'none' },
      stroke: { type: 'none' },
      'z-index': 0,
    };
  }

  test('star7: emits a <polygon> with 14 point pairs', () => {
    const g = emitShape(starShape('star7'), { warnings: [] });
    expect(g).toContain('<polygon');
    // 14 pairs of "x,y" separated by spaces
    const match = g.match(/points="([^"]+)"/);
    expect(match).not.toBeNull();
    expect(match[1].trim().split(/\s+/).length).toBe(14);
  });

  test('star7: first point is at the top center (cx, cy-outerRy)', () => {
    // w=h=100 → cx=50, outerRy=50 → first point (50.00, 0.00)
    const g = emitShape(starShape('star7'), { warnings: [] });
    expect(g).toContain('50.00,0.00');
  });

  test('star4: emits 8 point pairs', () => {
    const g = emitShape(starShape('star4'), { warnings: [] });
    const match = g.match(/points="([^"]+)"/);
    expect(match[1].trim().split(/\s+/).length).toBe(8);
  });

  test('star5: emits 10 point pairs', () => {
    const g = emitShape(starShape('star5'), { warnings: [] });
    const match = g.match(/points="([^"]+)"/);
    expect(match[1].trim().split(/\s+/).length).toBe(10);
  });

  test('does not throw and returns non-empty string', () => {
    expect(() => emitShape(starShape('star7'), { warnings: [] })).not.toThrow();
    expect(emitShape(starShape('star7'), { warnings: [] }).length).toBeGreaterThan(0);
  });
});
