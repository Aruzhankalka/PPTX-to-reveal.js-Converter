const { escapeHtml } = require('./escape');
const { positioningToCss } = require('./text');

/**
 * Render a media element. Sprint 1: images only.
 * Videos are accepted by the schema but emitted as a static placeholder
 * per the Requirements Analysis OoS-02.
 */
function renderMedia(media) {
  let css = positioningToCss(media);

  const isLargeInheritedImage =
  media.id?.startsWith('inherited-img') &&
  media['media-type'] === 'image' &&
  media.width > 500 &&
  media.height > 300;


  const isImage = media['media-type'] === 'image';
  const width = typeof media.width === 'number' ? media.width : 0;
  const height = typeof media.height === 'number' ? media.height : 0;
  
  // Small logo: width <= 250px AND height <= 120px
  const isSmallLogo = isImage && width <= 250 && height <= 120;
  
  // Large image (likely background): width > 500px OR height > 300px
  const isBackgroundImage = isImage && (width > 500 || height > 300);
  
  // Regular image
  const isRegularImage = isImage && !isSmallLogo && !isBackgroundImage;

    // Apply z-index based on media type
  if (isSmallLogo) {
    // Logo on top of everything
    css += '; z-index: 50';
  } else if (isLargeInheritedImage) {
    css += '; z-index: 0';  
  } else if (isBackgroundImage) {
    // Background image behind shapes and text
    css += '; z-index: 1';
  } else if (isRegularImage) {
    // Regular images above shapes
    css += '; z-index: 5';
  }

  const styleAttr = css ? ` style="${css}"` : '';

  if (media['media-type'] === 'image') {
    const src = escapeHtml(media['file-link'] || '');
    const alt = escapeHtml(media.id || 'image');

    
    return `<img src="${src}" alt="${alt}"${styleAttr} />`;
  }

  if (media['media-type'] === 'video') {
    return `<div class="video-placeholder"${styleAttr}>[Video placeholder: ${escapeHtml(media['file-link'] || '')}]</div>`;
  }

  return '';
}

module.exports = { renderMedia };