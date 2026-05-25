const { escapeHtml } = require('./escape');
const { renderTextBlock } = require('./text');
const { renderMedia } = require('./media');

/**
 * Wrap a slide's contents in a reveal.js <section> with semantic structure.
 * Per FR-04: one section per IR slide, in the same order.
 */
function renderSlide(slide) {
  const parts = [];
  const contents = slide.contents || {};

  for (const textBlock of (contents.text || [])) {
    parts.push('  ' + renderTextBlock(textBlock).replace(/\n/g, '\n  '));
  }
  for (const mediaEl of (contents.media || [])) {
    parts.push('  ' + renderMedia(mediaEl));
  }

  return `<section>\n${parts.join('\n')}\n</section>`;
}

/**
 * Build the complete reveal.js HTML document.
 * Uses CDN-hosted reveal.js v4.6 per spec TC-03.
 *
 * @param {object} ir - the validated IR document
 * @returns {string} - complete HTML
 */
function renderDocument(ir) {
  const slideset = ir.slideset || {};
  const docTitle = escapeHtml(slideset.title || slideset.filename || 'Presentation');
  const slidesHtml = (slideset.slides || []).map(renderSlide).join('\n\n');

  // FR-07: slide canvas size (defaults to standard 960×540 when not in IR)
  const slideWidth  = (slideset.master && slideset.master.slideWidth)  || 960;
  const slideHeight = (slideset.master && slideset.master.slideHeight) || 540;

  // FR-12: emit theme colours as CSS custom properties so FR-06 schemeClr
  // references like var(--theme-accent1) resolve in the browser
  const themeColors = (slideset.master && slideset.master.theme && slideset.master.theme.colors) || {};
  const cssVarBlock = Object.keys(themeColors).length
    ? `\n  :root {\n${Object.entries(themeColors).map(([k, v]) => `    --theme-${k}: ${v};`).join('\n')}\n  }`
    : '';

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${docTitle}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reset.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/theme/white.css">
  <style>${cssVarBlock}
    .reveal .slides section { text-align: left; }
    .reveal .slides section .text-block { box-sizing: border-box; }
    .reveal .slides section img { max-width: 100%; height: auto; }
  </style>
</head>
<body>
  <div class="reveal">
    <div class="slides">
${slidesHtml}
    </div>
  </div>
  <script src="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.js"></script>
  <script>
    Reveal.initialize({
      hash: true,
      controls: true,
      progress: true,
      width: ${slideWidth},
      height: ${slideHeight},
    });
  </script>
</body>
</html>`;
}

module.exports = { renderDocument, renderSlide };