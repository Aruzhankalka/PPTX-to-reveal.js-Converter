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
function renderSlide(slide, slideIndex) {
  const parts = [];
  const contents = slide.contents || {};
//   console.log(`SLIDE ${slideIndex + 1} KEYS:`, Object.keys(contents));
// console.log(`SLIDE ${slideIndex + 1} SHAPES COUNT:`, (contents.shapes || []).length);
  // if (slideIndex === 2) {
  //   console.log("SLIDE 2 MEDIA:", (contents.media || []).map(m => ({
  //     id: m.id,
  //     file: m['file-link'],
  //     x: m.position?.x,
  //     y: m.position?.y,
  //     w: m.width,
  //     h: m.height,
  //     z: m['z-index']
  //   })));
  // }

  const filteredMedia = [];
  const inheritedSmallMedia = [];
  
  for (const m of contents.media || []) {
    const isSmallInherited =
      m.id?.startsWith('inherited-img') &&
      typeof m.width === 'number' &&
      typeof m.height === 'number' &&
      m.width <= 250 &&
      m.height <= 140;
  
    if (isSmallInherited) {
      inheritedSmallMedia.push(m);
    } else {
      filteredMedia.push(m);
    }
  }
  
  if (inheritedSmallMedia.length > 0) {
    inheritedSmallMedia.sort((a, b) => (b['z-index'] || 0) - (a['z-index'] || 0));
    filteredMedia.push(inheritedSmallMedia[0]);
  }

 // STEP 1: Render background media first (z-index: 1)
  // These are large images from master slide that should be behind everything
  const backgroundMedia = filteredMedia.filter(m => {
    const width = typeof m.width === 'number' ? m.width : 0;
    const height = typeof m.height === 'number' ? m.height : 0;
    return (width > 500 || height > 300);
  });

  for (const mediaEl of backgroundMedia) {
    parts.push('    ' + renderMedia(mediaEl).replace(/\n/g, '\n    '));
  }

  // STEP 2: Render shapes (z-index: 10-40 from IR)
  for (const shape of contents.shapes || []) {
    // parts.push('    ' + renderShape(shape));
    parts.push('    ' + renderShape(shape, { renderText: false }));
  }

  // STEP 3: Render text blocks
  for (const textBlock of (contents.text || [])) {
    parts.push('    ' + renderTextBlock(textBlock).replace(/\n/g, '\n    '));
  }

  // STEP 4: Render regular images and logos (z-index: 5 and 50)
  const regularMedia = filteredMedia.filter(m => {
    const width = typeof m.width === 'number' ? m.width : 0;
    const height = typeof m.height === 'number' ? m.height : 0;
    return !(width > 500 || height > 300); // Not background image
  });
  
  for (const mediaEl of regularMedia) {
    parts.push('    ' + renderMedia(mediaEl).replace(/\n/g, '\n    '));
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

  const aliases = {
    bg1: colors.lt1 || colors.bg1 || '#FFFFFF',
    bg2: colors.lt2 || colors.bg2 || '#E7E6E6',
    text1: colors.dk1 || colors.text1 || '#000000',
    text2: colors.dk2 || colors.text2 || '#44546A',
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

  for (const [name, value] of Object.entries(aliases)) {
    if (value) {
      cssVariables.push(`--theme-${name}: ${escapeHtml(value)};`);
    }
  }

  return `:root {
      ${cssVariables.join('\n      ')}
    }`;
}


function renderDocument(ir) {
  const slideset = ir.slideset || {};
  const themeCss = renderThemeVariables(slideset.master);

  const docTitle = escapeHtml(slideset.title || slideset.filename || 'Presentation');
  // const slidesHtml = (slideset.slides || []).map(renderSlide).join('\n\n');
  const slidesHtml = (slideset.slides || []).map((slide, index) => renderSlide(slide, index)).join('\n\n');

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
      ${themeCss}
    /* overflow:visible overrides Reveal.js's own "section { overflow:hidden }" so
       elements whose bottom edge sits at the slide boundary are not clipped. */
    .reveal .slides section { text-align: left; overflow: visible; width: ${slideWidth}px; height: ${slideHeight}px; }
    .slide-canvas { position: relative; width: ${slideWidth}px; height: ${slideHeight}px; overflow: visible; }
    .slide-canvas .text-block {box-sizing: border-box; overflow: hidden; z-index: 20 !important;}
    .slide-canvas p { margin: 0; padding: 0; }
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