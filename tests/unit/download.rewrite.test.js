const { rewriteHtmlForBundle } = require("../../src/api/download");

describe("rewriteHtmlForBundle", () => {
  const ID = "abc123-def-456";

  test("rewrites /api/v1/media/{id}/... to assets/...", () => {
    const input = `<img src="/api/v1/media/${ID}/slide1.png" alt="x" />`;
    const out = rewriteHtmlForBundle(input, ID);
    expect(out).toBe(`<img src="assets/slide1.png" alt="x" />`);
  });

  test("rewrites multiple media URLs in one document", () => {
    const input = `
      <img src="/api/v1/media/${ID}/a.png" />
      <img src="/api/v1/media/${ID}/b.jpg" />
      <img src="/api/v1/media/${ID}/c.gif" />
    `;
    const out = rewriteHtmlForBundle(input, ID);
    expect(out).toContain('src="assets/a.png"');
    expect(out).toContain('src="assets/b.jpg"');
    expect(out).toContain('src="assets/c.gif"');
    expect(out).not.toContain("/api/v1/media/");
  });

  test("rewrites reveal.js CDN URLs to local reveal/ paths", () => {
    const input = `
      <link href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reset.css">
      <link href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.css">
      <link href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/theme/white.css">
      <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.js"></script>
    `;
    const out = rewriteHtmlForBundle(input, ID);
    expect(out).toContain('href="reveal/reset.css"');
    expect(out).toContain('href="reveal/reveal.css"');
    expect(out).toContain('href="reveal/theme/white.css"');
    expect(out).toContain('src="reveal/reveal.js"');
    expect(out).not.toContain("cdn.jsdelivr.net");
  });

  test("leaves URLs of other result IDs untouched", () => {
    const otherId = "xyz999";
    const input = `<img src="/api/v1/media/${otherId}/leak.png" />`;
    const out = rewriteHtmlForBundle(input, ID);
    expect(out).toBe(input);
  });

  test("is a no-op for HTML with no media or CDN references", () => {
    const input = `<section><h2>Hello</h2><p>No images here.</p></section>`;
    expect(rewriteHtmlForBundle(input, ID)).toBe(input);
  });

  test("handles HTML with both media and CDN references together", () => {
    const input = `
      <link href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.css">
      <img src="/api/v1/media/${ID}/diagram.png" />
    `;
    const out = rewriteHtmlForBundle(input, ID);
    expect(out).toContain('href="reveal/reveal.css"');
    expect(out).toContain('src="assets/diagram.png"');
  });
});
