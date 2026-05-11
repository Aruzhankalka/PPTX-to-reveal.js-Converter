const { escapeHtml } = require('./escape');
const { positioningToCss } = require('./text');

/**
 * Render a media element. Sprint 1: images only.
 * Videos are accepted by the schema but emitted as a static placeholder
 * per the Requirements Analysis OoS-02.
 */
function renderMedia(media) {
  const css = positioningToCss(media);
  const styleAttr = css ? ` style="${css}"` : '';

  if (media['media-type'] === 'image') {
    const src = escapeHtml(media['file-link'] || '');
    const alt = escapeHtml(media.id || 'image');
    const imgCss = css ? `${css}; height: auto` : '';
    const imgStyle = imgCss ? ` style="${imgCss}"` : '';
    return `<img src="${src}" alt="${alt}"${imgStyle} />`;
  }

  if (media['media-type'] === 'video') {
    return `<div class="video-placeholder"${styleAttr}>[Video placeholder: ${escapeHtml(media['file-link'] || '')}]</div>`;
  }

  return '';
}

module.exports = { renderMedia };