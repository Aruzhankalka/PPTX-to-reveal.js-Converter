const { escapeHtml, escapeCss } = require('./escape');

// EMU constants mirroring the parser — kept here so the generator is self-contained.
const EMU_PER_PT = 12700;
const EMU_PER_PX = 9525;

/**
 * Convert a CSS pt string to CSS px using the same EMU-based scale as geometry.
 *   px = pt × EMU_PER_PT / EMU_PER_PX
 * At 96 dpi this equals pt × (96/72), i.e. the standard CSS pt→px ratio.
 * Using the EMU path makes the relationship to geometry explicit and testable.
 *
 * Non-pt strings (unitless, em, px already, …) are returned unchanged.
 */
function ptToPx(size) {
  if (typeof size !== 'string') return escapeCss(String(size));

  const match = size.match(/^([\d.]+)pt$/);
  if (!match) return escapeCss(size);

  const pt = Number(match[1]);
  const px = Math.round(pt * 1.333);

  return `${px}px`;
}

/**
 * Convert a run's formatting object to an inline CSS string.
 * Handles bold, italics, underline/strikethrough, color, font-family, font-size,
 * line-height, and paragraph space-before/space-after.
 *
 * Font sizes stored as "Xpt" are converted to px via the EMU scale so they are
 * geometrically consistent with element positions/sizes (option-a approach).
 *
 * @param {object} formatting - the run's or paragraph's formatting object
 * @returns {string} - CSS declarations separated by ";", or ""
 */
function formattingToCss(formatting) {
  if (!formatting) return '';
  const decls = [];

  if (formatting.weight === 'bold') {
    decls.push('font-weight: bold');
  }
  if (formatting.italics === true) {
    decls.push('font-style: italic');
  }
  if (formatting['text-decoration'] === 'underline') {
    decls.push('text-decoration: underline');
  } else if (formatting['text-decoration'] === 'strikethrough') {
    decls.push('text-decoration: line-through');
  }
  if (formatting.color) {
    decls.push(`color: ${escapeCss(formatting.color)}`);
  }
  if (formatting.font) {
    decls.push(`font-family: ${escapeCss(formatting.font)}`);
  }
  if (formatting.size) {
    // Convert PowerPoint font size from pt to CSS px.
    decls.push(`font-size: ${ptToPx(formatting.size)}`);
  }
  if (formatting.align) {
    decls.push(`text-align: ${escapeCss(formatting.align)}`);
  }
  if (formatting['line-spacing']) {
    // Unitless values (from spcPct) pass through as-is; pt values are kept as pt
    // since CSS line-height in pt is valid and exact spacing is preserved.
    decls.push(`line-height: ${escapeCss(formatting['line-spacing'])}`);
  }
  if (formatting['space-before']) {
    decls.push(`margin-top: ${escapeCss(formatting['space-before'])}`);
  }
  if (formatting['space-after']) {
    decls.push(`margin-bottom: ${escapeCss(formatting['space-after'])}`);
  }

  return decls.join('; ');
}

/**
 * Render a single run as an HTML <span>.
 * Wraps in <a> when the run has a link.
 */
function renderRun(run) {
  const text = escapeHtml(run.text || '');
  const css = formattingToCss(run.formatting);
  const styleAttr = css ? ` style="${css}"` : '';

  let html = `<span${styleAttr}>${text}</span>`;

  if (run['super-sub-script'] === 'super') {
    html = `<sup>${html}</sup>`;
  } else if (run['super-sub-script'] === 'sub') {
    html = `<sub>${html}</sub>`;
  }

  if (run.link && run.link.href) {
    const href = escapeHtml(run.link.href);
    const target = run.link.target ? ` target="${escapeHtml(run.link.target)}"` : '';
    html = `<a href="${href}"${target}>${html}</a>`;
  }

  // console.log("RUN FORMATTING:", run.formatting);

  return html;
}

/**
 * Render a paragraph as an HTML <p> with paragraph-level formatting applied
 * and child <span>s for each run.
 */
function renderParagraph(paragraph) {
  const css = formattingToCss(paragraph.formatting);
  const styleAttr = css ? ` style="${css}"` : '';
  const runs = (paragraph.runs || []).map(renderRun).join('');
  return `<p${styleAttr}>${runs}</p>`;
}

/**
 * Render a text block (a positioned container of paragraphs).
 * Sprint 1: normal document flow — no absolute positioning.
 * Sprint 2 will apply exact PPTX geometry once slide dimensions are extracted.
 */
function renderTextBlock(textBlock) {
  const css = positioningToCss(textBlock);
  const styleAttr = css ? ` style="${css}"` : '';
  const paragraphs = (textBlock.paragraphs || []).map(renderParagraph).join('\n');
  return `<div class="text-block"${styleAttr}>\n${paragraphs}\n</div>`;
}

/**
 * Convert position/width/height/rotation into absolute-positioning CSS.
 * Shared between text blocks and media in Sprint 1.
 */
function positioningToCss(element) {
  // Footer placeholder: no explicit coordinates in the slide XML (position lives
  // in the slide layout). Pin it to the bottom of the slide canvas.
  if (element['footer-placement']) {
    return `position: absolute; bottom: 5px; left: 10px; right: 10px; font-size: ${ptToPx('12pt')}`;
  }

  const decls = [];
  if (element.position) {
    decls.push('position: absolute');
    if (typeof element.position.x === 'number') {
      decls.push(`left: ${element.position.x}px`);
    }
    if (typeof element.position.y === 'number') {
      decls.push(`top: ${element.position.y}px`);
    }
  }
  if (typeof element.width === 'number') {
    decls.push(`width: ${element.width}px`);
  }
  if (typeof element.height === 'number') {
    decls.push(`height: ${element.height}px`);
  }
  if (typeof element.rotation === 'number' && element.rotation !== 0) {
    decls.push(`transform: rotate(${element.rotation}deg)`);
  }
  if (typeof element['z-index'] === 'number') {
    decls.push(`z-index: ${element['z-index']}`);
  }
  // <a:bodyPr anchor> — vertical text alignment within the text body box.
  // anchor="t" is the CSS default (block flow from top), so no rule needed.
  // anchor="ctr" / anchor="b" need flex to push content to the correct edge.
  if (element['text-anchor'] === 'ctr') {
    decls.push('display: flex', 'flex-direction: column', 'justify-content: center');
  } else if (element['text-anchor'] === 'b') {
    decls.push('display: flex', 'flex-direction: column', 'justify-content: flex-end');
  }
  // IR overflow field: 'overflow-visible' overrides the blanket overflow:hidden
  // on .text-block so small footer/sldNum/dt boxes do not clip their content.
  if (element.overflow === 'overflow-visible') {
    decls.push('overflow: visible');
  }
  return decls.join('; ');
}

module.exports = { renderTextBlock, renderParagraph, renderRun, positioningToCss, formattingToCss };