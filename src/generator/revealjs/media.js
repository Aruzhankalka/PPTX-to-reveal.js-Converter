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
    return `<img src="${src}" alt="${alt}"${styleAttr} />`;
  }

  if (media['media-type'] === 'video') {
    // Placeholder per OoS-02 (no embedded video conversion in Sprint 1)
    return `<div class="video-placeholder"${styleAttr}>[Video placeholder: ${escapeHtml(media['file-link'] || '')}]</div>`;
  }

  return '';
}

module.exports = { renderMedia };