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

describe('emitShape — unsupported type stubs', () => {
  const STUBS = ['ellipse', 'line', 'arrow', 'polyline', 'polygon', 'callout', 'connector'];

  test.each(STUBS)('%s: returns empty string', (type) => {
    const result = emitShape(
      { type, position: { x: 0, y: 0 }, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX },
      { warnings: [] },
    );
    expect(result).toBe('');
  });

  test.each(STUBS)('%s: pushes exactly one warning containing the type name', (type) => {
    const ctx = { warnings: [] };
    emitShape(
      { type, position: { x: 0, y: 0 }, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX },
      ctx,
    );
    expect(ctx.warnings).toHaveLength(1);
    expect(ctx.warnings[0]).toContain(type);
  });

  test.each(STUBS)('%s: does not throw', (type) => {
    expect(() =>
      emitShape(
        { type, position: { x: 0, y: 0 }, width: 100 * EMU_PER_PX, height: 50 * EMU_PER_PX },
        { warnings: [] },
      ),
    ).not.toThrow();
  });

  test('unknown type pushes warning and returns empty string', () => {
    const ctx = { warnings: [] };
    const result = emitShape(
      { type: 'star', position: { x: 0, y: 0 }, width: 0, height: 0 },
      ctx,
    );
    expect(result).toBe('');
    expect(ctx.warnings[0]).toContain('star');
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

  test('returns empty string for unsupported type', () => {
    const result = renderShape({
      type: 'ellipse',
      position: { x: 0, y: 0 },
      width: 100 * EMU_PER_PX,
      height: 50 * EMU_PER_PX,
    });
    expect(result).toBe('');
  });
});
