const { generate } = require('../src/generator/revealjs');
const { renderRun, renderParagraph, formattingToCss, positioningToCss } = require('../src/generator/revealjs/text');
const { escapeHtml } = require('../src/generator/revealjs/escape');
const { warnOverflowElements, renderSlide } = require('../src/generator/revealjs/html');
const { runToIr, paragraphToIr } = require('../src/parser/pptx/text');
const fixture = require('./fixtures/minimal-ir.json');

describe('escapeHtml', () => {
  test('escapes the five HTML metacharacters', () => {
    expect(escapeHtml('<script>alert("x")</script>'))
      .toBe('&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;');
  });

  test('handles null and undefined', () => {
    expect(escapeHtml(null)).toBe('');
    expect(escapeHtml(undefined)).toBe('');
  });
});

// EMU constants (must match src/generator/revealjs/text.js)
const EMU_PER_PT = 12700;
const EMU_PER_PX = 9525;

describe('formattingToCss — font-size geometric scaling', () => {
  test('sz=2400 (24 pt) emits font-size: 32px  (pt × 12700 / 9525)', () => {
    const sz = 2400;
    const pt = sz / 100;
    const expectedPx = Math.round(pt * EMU_PER_PT / EMU_PER_PX); // 32
    const css = formattingToCss({ size: `${pt}pt` });
    expect(css).toContain(`font-size: ${expectedPx}px`);
  });

  test('sz=1800 (18 pt) emits font-size: 24px', () => {
    const css = formattingToCss({ size: '18pt' });
    expect(css).toContain('font-size: 24px');
  });

  test('sz=4000 (40 pt) emits font-size: 53px', () => {
    const css = formattingToCss({ size: '40pt' });
    expect(css).toContain('font-size: 53px');
  });

  test('non-pt size passes through unchanged', () => {
    const css = formattingToCss({ size: '2em' });
    expect(css).toContain('font-size: 2em');
  });

  test('emits margin-top from space-before', () => {
    const css = formattingToCss({ 'space-before': '6pt' });
    expect(css).toContain('margin-top: 6pt');
  });

  test('emits margin-bottom from space-after', () => {
    const css = formattingToCss({ 'space-after': '0pt' });
    expect(css).toContain('margin-bottom: 0pt');
  });

  test('emits line-height from line-spacing (unitless)', () => {
    const css = formattingToCss({ 'line-spacing': '1.5' });
    expect(css).toContain('line-height: 1.5');
  });
});

describe('renderRun', () => {
  test('renders plain text as a span', () => {
    expect(renderRun({ text: 'hello' })).toBe('<span>hello</span>');
  });

  test('applies bold formatting', () => {
    const out = renderRun({ text: 'bold', formatting: { weight: 'bold' } });
    expect(out).toContain('font-weight: bold');
  });

  test('escapes XSS attempts in text', () => {
    const out = renderRun({ text: '<img src=x onerror=alert(1)>' });
    expect(out).not.toContain('<img');
    expect(out).toContain('&lt;img');
  });

  test('wraps in anchor when link is present', () => {
    const out = renderRun({ text: 'click', link: { href: 'https://example.com' } });
    expect(out).toContain('<a href="https://example.com"');
  });
});

describe('renderParagraph', () => {
  test('produces a <p> containing all runs', () => {
    const html = renderParagraph({
      runs: [
        { text: 'Hello ' },
        { text: 'world', formatting: { weight: 'bold' } },
      ],
    });
    expect(html).toMatch(/^<p[^>]*>/);
    expect(html).toContain('Hello');
    expect(html).toContain('world');
  });
});

describe('generate (full document, FR-04)', () => {
  test('produces one section per IR slide', () => {
    const { html } = generate(fixture);
    const sections = html.match(/<section>/g) || [];
    expect(sections.length).toBe(fixture.slideset.slides.length);
  });

  test('output is a complete HTML document with reveal.js loaded (FR-05)', () => {
    const { html } = generate(fixture);
    expect(html).toMatch(/^<!DOCTYPE html>/);
    expect(html).toContain('reveal.js');
    expect(html).toContain('<div class="reveal">');
    expect(html).toContain('Reveal.initialize');
  });

  test('image media is emitted as <img> (FR-09)', () => {
    const { html } = generate(fixture);
    expect(html).toContain('<img src="media/diagram.png"');
  });

  test('rejects an invalid IR with a clear error', () => {
    expect(() => generate({ slideset: { /* missing slides */ } }))
      .toThrow(/IR validation failed/);
  });

  test('preserves slide order (FR-04)', () => {
    const { html } = generate(fixture);
    const welcomeIdx = html.indexOf('Welcome');
    const imageIdx = html.indexOf('image');
    expect(welcomeIdx).toBeLessThan(imageIdx);
    expect(welcomeIdx).toBeGreaterThan(-1);
  });

  test('emits theme colours as CSS custom properties (FR-12)', () => {
    const ir = JSON.parse(JSON.stringify(fixture));
    ir.slideset.master = {
      theme: { colors: { dk1: '#000000', accent1: '#4472C4' }, fonts: {} },
    };
    const { html } = generate(ir);
    expect(html).toContain('--theme-dk1: #000000');
    expect(html).toContain('--theme-accent1: #4472C4');
  });

  test('passes slide dimensions to Reveal.initialize (FR-07)', () => {
    const ir = JSON.parse(JSON.stringify(fixture));
    ir.slideset.master = { 'slide-dimensions': { width: 1280, height: 720 } };
    const { html } = generate(ir);
    // dimensions belong in Reveal.initialize, not in section CSS
    expect(html).toContain('width: 1280,');
    expect(html).toContain('height: 720,');
  });

  test('falls back to 960x540 when master has no dimensions (FR-07)', () => {
    const { html } = generate(fixture);
    expect(html).toContain('width: 960,');
    expect(html).toContain('height: 540,');
  });

  test('slide-canvas uses overflow:visible so near-edge elements are not clipped', () => {
    // overflow:hidden on .slide-canvas clips absolutely-positioned children
    // whose bottom edge extends past the declared slide height (e.g. footer and
    // URL placeholders resolved from the layout).  overflow:visible lets them
    // render fully while Reveal.js handles viewport-level clipping.
    // Note: .text-block correctly keeps overflow:hidden to clip long text at
    // its own box — only the outer canvas container needs the change.
    const { html } = generate(fixture);
    const canvasRule = html.match(/\.slide-canvas\s*\{[^}]+\}/)?.[0] ?? '';
    expect(canvasRule).toContain('overflow: visible');
    expect(canvasRule).not.toContain('overflow: hidden');
  });

  test('slide canvas dimensions in CSS match Reveal.initialize dimensions (scale once)', () => {
    const ir = JSON.parse(JSON.stringify(fixture));
    ir.slideset.master = { 'slide-dimensions': { width: 1280, height: 720 } };
    const { html } = generate(ir);
    // Same px values must appear in both the CSS block and the JS initializer
    // to guarantee the coordinate space inside .slide-canvas matches what
    // Reveal.js believes the slide size is.
    expect(html).toContain('width: 1280px');
    expect(html).toContain('height: 720px');
    expect(html).toContain('width: 1280,');
    expect(html).toContain('height: 720,');
  });

  test('Reveal.initialize sets center:false and margin:0 (prevent coordinate shift)', () => {
    // Reveal.js defaults (center:true, margin:0.04) shift and shrink the effective
    // slide area, clipping bottom-edge content that is geometrically within bounds.
    const { html } = generate(fixture);
    expect(html).toContain('center: false');
    expect(html).toContain('margin: 0,');
  });

  test('section rule sets overflow:visible to override Reveal.js default overflow:hidden', () => {
    // Reveal.js ships ".reveal .slides section { overflow: hidden }".  Our style block
    // (loaded after the CDN CSS) must override this so the .slide-canvas children are
    // not clipped at the section boundary.
    const { html } = generate(fixture);
    const sectionRule = html.match(/\.reveal\s+\.slides\s+section\s*\{[^}]+\}/)?.[0] ?? '';
    expect(sectionRule).toContain('overflow: visible');
    expect(sectionRule).not.toContain('overflow: hidden');
  });
});

describe('warnOverflowElements', () => {
  let warnSpy;
  beforeEach(() => { warnSpy = jest.spyOn(console, 'warn').mockImplementation(() => {}); });
  afterEach(() => { warnSpy.mockRestore(); });

  test('warns when element bottom exceeds slideHeight', () => {
    const slides = [{ contents: { text: [{ id: 'footer1', position: { y: 520 }, height: 48 }] } }];
    warnOverflowElements(slides, 540);
    expect(warnSpy).toHaveBeenCalled();
    const output = warnSpy.mock.calls.flat().join(' ');
    expect(output).toContain('slideHeight=540px');
    expect(output).toContain('footer1');
    expect(output).toContain('top=520px');
    expect(output).toContain('height=48px');
  });

  test('does not warn when element bottom is within slideHeight', () => {
    const slides = [{ contents: { text: [{ id: 'b1', position: { y: 400 }, height: 60 }] } }];
    warnOverflowElements(slides, 540);
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('does not warn when element sits exactly at the boundary', () => {
    const slides = [{ contents: { text: [{ id: 'b1', position: { y: 492 }, height: 48 }] } }];
    warnOverflowElements(slides, 540);  // 492 + 48 = 540, not >540
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('handles null slides gracefully', () => {
    expect(() => warnOverflowElements(null, 540)).not.toThrow();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('handles empty slides array gracefully', () => {
    expect(() => warnOverflowElements([], 540)).not.toThrow();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('handles blocks without position', () => {
    const slides = [{ contents: { text: [{ id: 'b1', height: 60 }] } }];
    expect(() => warnOverflowElements(slides, 540)).not.toThrow();
    expect(warnSpy).not.toHaveBeenCalled();
  });

  test('handles blocks without height (treats as 0)', () => {
    const slides = [{ contents: { text: [{ id: 'b1', position: { y: 600 } }] } }];
    warnOverflowElements(slides, 540);
    expect(warnSpy).toHaveBeenCalled();
    const output = warnSpy.mock.calls.flat().join(' ');
    expect(output).toContain('height=0px');
  });

  test('reports count of overflowing elements per slide', () => {
    const slides = [{
      contents: {
        text: [
          { id: 'a', position: { y: 520 }, height: 48 },
          { id: 'b', position: { y: 530 }, height: 30 },
          { id: 'c', position: { y: 100 }, height: 50 },  // in bounds
        ],
      },
    }];
    warnOverflowElements(slides, 540);
    const firstCall = warnSpy.mock.calls[0][0];
    expect(firstCall).toContain('2 element(s)');
  });
});

// ---------------------------------------------------------------------------
// renderSlide — hidden, transition, background, notes
// ---------------------------------------------------------------------------

describe('renderSlide hidden slide', () => {
  const emptySlide = { contents: { text: [], media: [], shapes: [] } };

  test('hidden:true adds data-visibility="hidden" to section', () => {
    const html = renderSlide({ ...emptySlide, hidden: true });
    expect(html).toContain('data-visibility="hidden"');
  });

  test('hidden:false produces no data-visibility attribute', () => {
    const html = renderSlide({ ...emptySlide, hidden: false });
    expect(html).not.toContain('data-visibility');
  });

  test('absent hidden field produces no data-visibility attribute', () => {
    const html = renderSlide(emptySlide);
    expect(html).not.toContain('data-visibility');
  });
});

describe('renderSlide transition', () => {
  const base = { contents: { text: [], media: [], shapes: [] } };

  test('transition in contents emits data-transition attribute', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, transition: 'fade' } });
    expect(html).toContain('data-transition="fade"');
  });

  test('transition value is HTML-escaped', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, transition: 'slide<x>' } });
    expect(html).toContain('data-transition="slide&lt;x&gt;"');
    expect(html).not.toContain('<x>');
  });

  test('absent transition produces no data-transition attribute', () => {
    const html = renderSlide(base);
    expect(html).not.toContain('data-transition');
  });
});

describe('renderSlide background', () => {
  const base = { contents: { text: [], media: [], shapes: [] } };

  test('background in contents emits inline style', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, background: '#FF0000' } });
    expect(html).toContain('style="background: #FF0000;"');
  });

  test('theme variable background is emitted as-is in style', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, background: 'var(--theme-accent1)' } });
    expect(html).toContain('style="background: var(--theme-accent1);"');
  });

  test('absent background produces no style attribute', () => {
    const html = renderSlide(base);
    expect(html).not.toContain('style=');
  });
});

describe('renderSlide notes', () => {
  const base = { contents: { text: [], media: [], shapes: [] } };

  test('notes in contents emits <aside class="notes">', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, notes: 'Speaker note' } });
    expect(html).toContain('<aside class="notes">Speaker note</aside>');
  });

  test('multi-line notes are joined with <br/>', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, notes: 'Line one\nLine two' } });
    expect(html).toContain('Line one<br/>Line two');
  });

  test('notes text is HTML-escaped', () => {
    const html = renderSlide({ ...base, contents: { ...base.contents, notes: '<b>bold</b>' } });
    expect(html).toContain('&lt;b&gt;bold&lt;/b&gt;');
    expect(html).not.toContain('<b>bold</b>');
  });

  test('absent notes produces no aside element', () => {
    const html = renderSlide(base);
    expect(html).not.toContain('<aside');
  });
});

// ---------------------------------------------------------------------------
// positioningToCss — text-anchor → CSS vertical alignment
// ---------------------------------------------------------------------------

describe('positioningToCss — text-anchor vertical alignment', () => {
  const base = { position: { x: 88, y: 674 }, width: 432, height: 21 };

  test('anchor="ctr" emits flex column centering', () => {
    const css = positioningToCss({ ...base, 'text-anchor': 'ctr' });
    expect(css).toContain('display: flex');
    expect(css).toContain('flex-direction: column');
    expect(css).toContain('justify-content: center');
  });

  test('anchor="b" emits flex column bottom alignment', () => {
    const css = positioningToCss({ ...base, 'text-anchor': 'b' });
    expect(css).toContain('display: flex');
    expect(css).toContain('flex-direction: column');
    expect(css).toContain('justify-content: flex-end');
  });

  test('anchor="t" emits no flex CSS (default block flow from top)', () => {
    const css = positioningToCss({ ...base, 'text-anchor': 't' });
    expect(css).not.toContain('flex');
    expect(css).not.toContain('justify-content');
  });

  test('absent text-anchor emits no flex CSS', () => {
    const css = positioningToCss(base);
    expect(css).not.toContain('flex');
    expect(css).not.toContain('justify-content');
  });

  test('anchor="ctr" does not add justify-content to blocks without position', () => {
    // footer-placement path is independent of text-anchor
    const css = positioningToCss({ 'footer-placement': true });
    expect(css).toContain('bottom: 5px');
    expect(css).not.toContain('justify-content');
  });

  // ── Key safety test ───────────────────────────────────────────────────────
  // A footer placeholder whose box bottom sits exactly at sldSz.cy (e.g.
  // y=699 height=21 in a 720px slide) must be top-anchored in the generated
  // CSS.  An anchor="ctr" or anchor="b" CSS would shift flex content toward or
  // past the slide boundary; top-anchoring keeps the readable cap/ascender
  // region within the slide.
  test('footer block at slide bottom boundary has no flex CSS (text within slideHeightPx)', () => {
    const slideHeight = 720;
    const blockHeight = 21;
    const blockY = slideHeight - blockHeight;  // bottom exactly at slide edge

    // Simulate what the parser sets for footer placeholders: text-anchor='t'
    // (the override in slide.js regardless of the PPTX template anchor value).
    const footerBlock = {
      position: { x: 88, y: blockY },
      width: 432,
      height: blockHeight,
      'text-anchor': 't',  // footer override
    };
    const css = positioningToCss(footerBlock);

    // No flex centering — with anchor="t" the text starts at top of box (blockY)
    // so top + line-height is at most blockY + font-size ≤ slideHeight.
    expect(css).not.toContain('justify-content');
    expect(css).not.toContain('flex');
    expect(css).toContain(`top: ${blockY}px`);
  });
});

// ---------------------------------------------------------------------------
// positioningToCss — overflow field
// ---------------------------------------------------------------------------

describe('positioningToCss — overflow field', () => {
  const base = { position: { x: 88, y: 674 }, width: 432, height: 21 };

  test('overflow="overflow-visible" emits overflow: visible', () => {
    const css = positioningToCss({ ...base, overflow: 'overflow-visible' });
    expect(css).toContain('overflow: visible');
  });

  test('absent overflow field does not emit overflow CSS (CSS rule handles it)', () => {
    const css = positioningToCss(base);
    expect(css).not.toContain('overflow');
  });

  test('overflow="overflow-visible" combined with text-anchor="ctr" emits both', () => {
    const css = positioningToCss({ ...base, 'text-anchor': 'ctr', overflow: 'overflow-visible' });
    expect(css).toContain('overflow: visible');
    expect(css).toContain('justify-content: center');
  });
});
// ---------------------------------------------------------------------------
// runToIr — fallbackFont applied when run has no explicit <a:latin> (c)
// ---------------------------------------------------------------------------

describe('runToIr — fallbackFont cascade', () => {
  // Minimal <a:r> node with only a text value and a lang attribute — no <a:latin>
  const runNoFont = { 'a:t': 'Hello', 'a:rPr': { '@_lang': 'en-US' } };
  // <a:r> that carries an explicit <a:latin typeface="Arial">
  const runWithFont = {
    'a:t': 'Hello',
    'a:rPr': { '@_lang': 'en-US', 'a:latin': { '@_typeface': 'Arial' } },
  };

  test('run with no <a:latin> inherits fallbackFont "Calibri"', () => {
    const run = runToIr(runNoFont, null, null, null, 'Calibri');
    expect(run.formatting.font).toBe('Calibri');
  });

  test('run with explicit <a:latin typeface="Arial"> keeps its own font, ignores fallbackFont', () => {
    const run = runToIr(runWithFont, null, null, null, 'Calibri');
    expect(run.formatting.font).toBe('Arial');
  });

  test('run with no font and no fallbackFont has no font in formatting', () => {
    const run = runToIr(runNoFont, null, null, null, null);
    // formatting may be absent entirely when no properties are set
    expect(run.formatting?.font).toBeUndefined();
  });

  test('fallbackFont does not override explicit font even when fallback is different', () => {
    const run = runToIr(runWithFont, null, null, null, 'Times New Roman');
    expect(run.formatting.font).toBe('Arial');
  });
});

// ---------------------------------------------------------------------------
// paragraphToIr — tab character and tab stop support
// ---------------------------------------------------------------------------

describe('paragraphToIr — tab run emission', () => {
  // (a) A <a:tab/> XML element between runs produces a { type:'tab' } run.
  test('emits { type:"tab" } run when <a:tab/> is present as a paragraph child', () => {
    const aP = {
      'a:r': [{ 'a:t': 'Before' }, { 'a:t': 'After' }],
      'a:tab': {}, // parsed representation of a self-closing <a:tab/> element
    };
    const para = paragraphToIr(aP, 0, null, null, null, null, null, true);
    expect(para.runs.some((r) => r.type === 'tab')).toBe(true);
  });

  // (b) <a:pPr><a:tabLst><a:tab l="..." algn="..."/> produces correct tabStops.
  test('populates tabStops from <a:tabLst> with correct EMU values', () => {
    const aP = {
      'a:pPr': {
        'a:tabLst': {
          // Single tab stop at 2 743 200 EMU ≈ 3 inches from left margin
          'a:tab': { '@_l': '2743200', '@_algn': 'l' },
        },
      },
      'a:r': [{ 'a:t': 'text' }],
    };
    const para = paragraphToIr(aP, 0, null, null, null, null, null, true);
    expect(para.tabStops).toEqual([{ pos: 2743200, align: 'l' }]);
  });

  // Bonus: literal \t in <a:t> is split into text runs with a tab marker between them.
  test('splits literal \\t in run text into two runs with { type:"tab" } between them', () => {
    const aP = {
      'a:r': [{ 'a:t': 'Aufgabe 1\t10 min' }],
    };
    const para = paragraphToIr(aP, 0, null, null, null, null, null, true);
    expect(para.runs).toHaveLength(3);
    expect(para.runs[0].text).toBe('Aufgabe 1');
    expect(para.runs[1]).toEqual({ type: 'tab' });
    expect(para.runs[2].text).toBe('10 min');
  });
});

describe('renderParagraph — empty paragraph blank-line rendering', () => {
  // (b) A zero-runs paragraph emits a <p> whose height is derived from formatting.size.
  test('zero-runs paragraph emits <p> with height derived from font size (28pt→37px)', () => {
    // ptToPx('28pt') = Math.round(28 * 12700 / 9525) = Math.round(37.33) = 37 → '37px'
    const html = renderParagraph({ id: 'p-0', runs: [], formatting: { size: '28pt' } });
    expect(html).toContain('height:37px');
    expect(html).toContain('margin:0');
    expect(html).toContain('line-height:1');
  });

  test('zero-runs paragraph with no size falls back to 12pt (→16px)', () => {
    // ptToPx('12pt') = Math.round(12 * 12700 / 9525) = Math.round(15.99) = 16 → '16px'
    const html = renderParagraph({ id: 'p-0', runs: [] });
    expect(html).toContain('height:16px');
  });

  test('zero-runs paragraph does not emit a <span> or inner content', () => {
    const html = renderParagraph({ id: 'p-0', runs: [], formatting: { size: '28pt' } });
    expect(html).not.toContain('<span');
    expect(html).not.toContain('</span>');
  });
});

describe('renderParagraph — tab run spacing', () => {
  test('tab run with no tabStops emits min-width:2em spacer', () => {
    const para = {
      id: 'p-0',
      runs: [{ text: 'Before' }, { type: 'tab' }, { text: 'After' }],
    };
    const html = renderParagraph(para);
    expect(html).toContain('display:inline-block;min-width:2em');
    expect(html).toContain('Before');
    expect(html).toContain('After');
  });

  test('tab run with tabStops emits a fixed-width spacer (not min-width)', () => {
    const para = {
      id: 'p-0',
      tabStops: [{ pos: 914400, align: 'l' }], // 1 inch = 96 px
      runs: [{ text: 'A' }, { type: 'tab' }, { text: 'B' }],
    };
    const html = renderParagraph(para);
    // Should use width:Npx, not min-width
    expect(html).toContain('display:inline-block;width:');
    expect(html).not.toContain('min-width');
  });

  test('paragraph with no tab runs is unaffected (same output as before)', () => {
    const para = {
      id: 'p-0',
      runs: [{ text: 'Hello ' }, { text: 'world', formatting: { weight: 'bold' } }],
    };
    const html = renderParagraph(para);
    expect(html).toMatch(/^<p[^>]*>/);
    expect(html).toContain('Hello');
    expect(html).toContain('world');
    expect(html).not.toContain('inline-block');
  });
});
