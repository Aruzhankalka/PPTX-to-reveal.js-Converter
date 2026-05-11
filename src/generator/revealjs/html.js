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

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${docTitle}</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reset.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/reveal.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/reveal.js@4.6.1/dist/theme/white.css">
  <style>
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
    });
  </script>
</body>
</html>`;
}

module.exports = { renderDocument, renderSlide };