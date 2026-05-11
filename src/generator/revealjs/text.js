const { escapeHtml, escapeCss } = require('./escape');

/**
 * Convert a run's formatting object to an inline CSS string.
 * Handles bold, italics, underline/strikethrough, color, font-family, font-size.
 * Returns empty string when no formatting is present, so we can omit style="".
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
    decls.push(`font-size: ${escapeCss(formatting.size)}`);
  }
  if (formatting.align) {
    decls.push(`text-align: ${escapeCss(formatting.align)}`);
  }
  if (formatting['line-spacing']) {
    decls.push(`line-height: ${escapeCss(formatting['line-spacing'])}`);
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
  return decls.join('; ');
}

module.exports = { renderTextBlock, renderParagraph, renderRun, positioningToCss };