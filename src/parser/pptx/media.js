const { emuToPx, pptxRotationToDegrees } = require('./units');
const { IDENTITY_TRANSFORM, mapBoxThroughTransform } = require('./shapes');

/**
 * Detect a picture shape <p:pic> and convert it to an IR media item.
 *
 * <p:pic>
 *   <p:nvPicPr><p:cNvPr id="..." name="..."/></p:nvPicPr>
 *   <p:blipFill>
 *     <a:blip r:embed="rId3"/>     <-- relationship to the actual image file
 *   </p:blipFill>
 *   <p:spPr>
 *     <a:xfrm>...</a:xfrm>
 *   </p:spPr>
 * </p:pic>
 *
 * @param {object} pPic - the parsed <p:pic> node
 * @param {object} slideRels - relationships of this slide (rId -> { target })
 * @param {string} slideDir - directory of the slide file (for resolveTarget)
 * @param {function} resolveTarget - from relationships.js
 * @param {number} idx - index for stable ID
 * @param {object} [transform] - accumulated ancestor group transform (see
 *   shapes.js); IDENTITY_TRANSFORM when the picture is not inside a <p:grpSp>
 */
function pictureToMedia(pPic, slideRels, slideDir, resolveTarget, idx, transform = IDENTITY_TRANSFORM) {
  const blipFill = pPic['p:blipFill'];
  if (!blipFill || !blipFill['a:blip']) return null;
  const rId = blipFill['a:blip']['@_r:embed'];
  if (!rId) return null;

  const rel = slideRels[rId];
  if (!rel) return null;

  // The media file is relative to the slide directory.
  const filePath = resolveTarget(slideDir, rel.target);

  const media = {
    id: 'img-' + idx,
    'file-link': filePath,
    'media-type': 'image',
    position: { x: 0, y: 0 },
    width: 0,
    height: 0,
  };

  // Position from <p:spPr><a:xfrm>...</a:xfrm></p:spPr>, corrected through
  // any ancestor group transform (chOff/chExt scale + rotation) before
  // converting EMU to px, so grouped pictures land in the right spot.
  const xfrm = pPic['p:spPr'] && pPic['p:spPr']['a:xfrm'];
  if (xfrm) {
    const offX = xfrm['a:off'] ? (Number(xfrm['a:off']['@_x']) || 0) : 0;
    const offY = xfrm['a:off'] ? (Number(xfrm['a:off']['@_y']) || 0) : 0;
    const extW = xfrm['a:ext'] ? (Number(xfrm['a:ext']['@_cx']) || 0) : 0;
    const extH = xfrm['a:ext'] ? (Number(xfrm['a:ext']['@_cy']) || 0) : 0;
    const rotUnits = xfrm['@_rot'] != null ? (Number(xfrm['@_rot']) || 0) : 0;

    const { position, rotation } = mapBoxThroughTransform(offX, offY, extW, extH, rotUnits, transform);
    media.position = { x: emuToPx(position.x) || 0, y: emuToPx(position.y) || 0 };
    media.width  = emuToPx(position.w) || 0;
    media.height = emuToPx(position.h) || 0;
    const deg = pptxRotationToDegrees(rotation);
    if (deg !== 0) media.rotation = deg;
  }

  // Crop from <p:blipFill><a:srcRect l t r b> — values are percentage × 1000
  // (100000 = 100%). Store as decimal fractions [top, right, bottom, left].
  const srcRect = blipFill['a:srcRect'];
  if (srcRect) {
    const t = (Number(srcRect['@_t']) || 0) / 100000;
    const r = (Number(srcRect['@_r']) || 0) / 100000;
    const b = (Number(srcRect['@_b']) || 0) / 100000;
    const l = (Number(srcRect['@_l']) || 0) / 100000;
    if (t || r || b || l) media.crop = [t, r, b, l];
  }

  return media;
}

module.exports = { pictureToMedia };