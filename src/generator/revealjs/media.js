/**
 * Media renderer — turns an IR media item (image or video reference) into
 * an absolutely-positioned HTML element, using the same positioning CSS
 * helper as text blocks (text.js's positioningToCss).
 */

const { escapeHtml } = require('./escape');
const { positioningToCss } = require('./text');

/**
 * Render a media element. Sprint 1: images only.
 * Videos are accepted by the schema but emitted as a static placeholder
 * per the Requirements Analysis OoS-02.
 *
 * @param {object} media - IR media item ({media-type, file-link, id, ...position fields})
 * @returns {string} an HTML <img> (or placeholder) element
 */
function renderMedia(media) {
  const css = positioningToCss(media);
  const styleAttr = css ? ` style="${css}"` : '';

  if (media['media-type'] === 'image') {
    const src = escapeHtml(media['file-link'] || '');
    const alt = escapeHtml(media.id || 'image');

    // crop: [top, right, bottom, left] as decimal fractions (0–1).
    // The xfrm dimensions (width/height) describe the VISIBLE area after
    // cropping.  We recover the full image size, then offset it so only
    // the intended region shows through the overflow:hidden wrapper.
    const crop = media.crop;
    if (crop && (crop[0] || crop[1] || crop[2] || crop[3])) {
      const [ct, cr, cb, cl] = crop;
      const visW = media.width  || 0;
      const visH = media.height || 0;

      // Full (uncropped) image dimensions in px.
      const fullW = cl + cr < 1 ? visW / (1 - cl - cr) : visW;
      const fullH = ct + cb < 1 ? visH / (1 - ct - cb) : visH;

      // Negative offset shifts the image so the crop origin aligns with the
      // top-left corner of the wrapper.
      const offX = -(cl * fullW);
      const offY = -(ct * fullH);

      // Wrapper carries all positioning; img is absolute inside it.
      const wrapStyle = css + '; overflow: hidden';
      const imgStyle  = [
        'position: absolute',
        `width: ${fullW}px`,
        `height: ${fullH}px`,
        `left: ${offX}px`,
        `top: ${offY}px`,
      ].join('; ');

      return `<div style="${wrapStyle}"><img src="${src}" alt="${alt}" style="${imgStyle}" /></div>`;
    }

    const imgStyle = css || '';
    const imgStyleAttr = imgStyle ? ` style="${imgStyle}"` : '';
    return `<img src="${src}" alt="${alt}"${imgStyleAttr} />`;
  }

  if (media['media-type'] === 'video') {
    return `<div class="video-placeholder"${styleAttr}>[Video placeholder: ${escapeHtml(media['file-link'] || '')}]</div>`;
  }

  return '';
}

module.exports = { renderMedia };
