'use strict';

const { emitShape, renderShape } = require('../src/generator/revealjs/svg');

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
