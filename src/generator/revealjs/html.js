const { escapeHtml } = require('./escape');
const { renderTextBlock } = require('./text');
const { renderMedia } = require('./media');
const { renderShape } = require('./svg');

/**
 * Wrap a slide's contents in a reveal.js <section> with semantic structure.
 * Per FR-04: one section per IR slide, in the same order.
 *
 * A .slide-canvas div acts as the positioned ancestor for absolutely-placed
 * children. The section itself must keep position:absolute (Reveal.js depends
 * on it for transitions), so we cannot use position:relative on the section.
 */
function renderSlide(slide) {
  const parts = [];
  const contents = slide.contents || {};

  for (const textBlock of (contents.text || [])) {
    parts.push('    ' + renderTextBlock(textBlock).replace(/\n/g, '\n    '));
  }
  for (const mediaEl of (contents.media || [])) {
    parts.push('    ' + renderMedia(mediaEl));
  }

  // Shapes
  for (const shape of contents.shapes || []) {
    parts.push('    ' + renderShape(shape));
  }

  const layoutAttr = slide.layoutName ? ` data-layout="${escapeHtml(slide.layoutName)}"` : '';
  return `<section${layoutAttr}>\n  <div class="slide-canvas">\n${parts.join('\n')}\n  </div>\n</section>`;
}

/**
 * Log any text-block elements whose bottom edge (top + height) extends past
 * the declared slide height.  Distinguishes genuine placeholder-geometry
 * overflow (root cause: the layout/master footer box extends beyond sldSz)
 * from a pure Reveal.js config issue (center:true or margin>0 shrinking the
 * effective slide area).
 *
 * @param {object[]} slides        IR slide array
 * @param {number}   slideHeightPx declared slide height in pixels
 */
function warnOverflowElements(slides, slideHeightPx) {
  (slides || []).forEach((slide, si) => {
    const overflowing = ((slide.contents && slide.contents.text) || []).filter(
      (b) => b.position && (b.position.y + (b.height || 0)) > slideHeightPx
    );
    if (overflowing.length === 0) return;

    console.warn(
      `[generate] slideHeight=${slideHeightPx}px. ` +
      `Slide ${si + 1} has ${overflowing.length} element(s) extending past the slide bottom:`
    );
    for (const b of overflowing) {
      const bottom = b.position.y + (b.height || 0);
      console.warn(
        `  ${b.id || '?'}: top=${b.position.y}px height=${b.height || 0}px ` +
        `bottom=${bottom}px (${bottom - slideHeightPx}px past edge)`
      );
    }
  });
}

/**
 * Build the complete reveal.js HTML document.
 * Uses CDN-hosted reveal.js v4.6 per spec TC-03.
 *
 * @param {object} ir - the validated IR document
 * @returns {string} - complete HTML
 */

function renderThemeVariables(master) {
  const colors = master?.theme?.colors || {};
  const fonts = master?.theme?.fonts || {};

  const cssVariables = [];

  const colorMap = {
    accent1: colors.accent1,
    accent2: colors.accent2,
    accent3: colors.accent3,
    accent4: colors.accent4,
    accent5: colors.accent5,
    accent6: colors.accent6,
    'text-dark': colors.dk1,
    'text-light': colors.lt1,
    'bg-dark': colors.dk2,
    'bg-light': colors.lt2,
    link: colors.hlink,
    'link-visited': colors.folHlink
  };

  for (const [name, value] of Object.entries(colorMap)) {
    if (value) {
      cssVariables.push(`--${name}: ${escapeHtml(value)};`);
    }
  }

  if (fonts.major) {
    cssVariables.push(`--font-major: ${escapeHtml(fonts.major)};`);
  }

  if (fonts.minor) {
    cssVariables.push(`--font-minor: ${escapeHtml(fonts.minor)};`);
  }

  if (cssVariables.length === 0) {
    return '';
  }

  return `:root {
      ${cssVariables.join('\n      ')}
    }`;
}


function renderDocument(ir) {
  const slideset = ir.slideset || {};
  const themeCss = renderThemeVariables(slideset.master);

  const docTitle = escapeHtml(slideset.title || slideset.filename || 'Presentation');
  const slidesHtml = (slideset.slides || []).map(renderSlide).join('\n\n');

  // FR-07: slide canvas size (defaults to standard 960×540 when not in IR)
  const slideWidth  = (slideset.master && slideset.master.slideWidth)  || 960;
  const slideHeight = (slideset.master && slideset.master.slideHeight) || 540;

  // Diagnostic: warn if any element's bottom extends past the slide boundary.
  warnOverflowElements(slideset.slides, slideHeight);

  // FR-12: emit theme colours as CSS custom properties so FR-06 schemeClr
  // references like var(--theme-accent1) resolve in the browser.
  // Also emit IR-level alias names used by shape fills/strokes (SCHEME_TO_REF
  // maps dk1→text1, lt1→bg1, hlink→link, folHlink→linkVisited) so that
  // var(--theme-text1) and var(--theme-bg1) resolve instead of falling back to
  // the SVG initial fill value (black).
  const themeColors = (slideset.master && slideset.master.theme && slideset.master.theme.colors) || {};
  const THEME_ALIASES = {
    text1: 'dk1', text2: 'dk2',
    bg1: 'lt1',   bg2: 'lt2',
    link: 'hlink', linkVisited: 'folHlink',
  };
  const cssVarLines = Object.entries(themeColors).map(([k, v]) => `    --theme-${k}: ${v};`);
  for (const [alias, base] of Object.entries(THEME_ALIASES)) {
    if (themeColors[base]) cssVarLines.push(`    --theme-${alias}: ${themeColors[base]};`);
  }
  const cssVarBlock = cssVarLines.length
    ? `\n  :root {\n${cssVarLines.join('\n')}\n  }`
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
    /* overflow:visible overrides Reveal.js's own "section { overflow:hidden }" so
       elements whose bottom edge sits at the slide boundary are not clipped. */
    .reveal .slides section { text-align: left; overflow: visible; width: ${slideWidth}px; height: ${slideHeight}px; }
    .slide-canvas { position: relative; width: ${slideWidth}px; height: ${slideHeight}px; overflow: visible; }
    .slide-canvas .text-block { box-sizing: border-box; overflow: hidden; }
    .slide-canvas p { margin: 0; }
    .slide-canvas img { max-width: none; max-height: none; margin: 0; }
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
      // center:false and margin:0 are required for pixel-positioned layouts.
      // Reveal.js defaults (center:true, margin:0.04) shift and shrink the
      // effective slide area, clipping bottom-edge content.
      center: false,
      margin: 0,
      width: ${slideWidth},
      height: ${slideHeight},
    });
  </script>
</body>
</html>`;
}

module.exports = { renderDocument, renderSlide, warnOverflowElements };