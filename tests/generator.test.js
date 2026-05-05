const { generate } = require('../src/generator/revealjs');
const { renderRun, renderParagraph } = require('../src/generator/revealjs/text');
const { escapeHtml } = require('../src/generator/revealjs/escape');
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
});