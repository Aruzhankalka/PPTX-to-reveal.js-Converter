const { asArray } = require('./xml');
const { emuToPx, pptxRotationToDegrees } = require('./units');

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
 */
function pictureToMedia(pPic, slideRels, slideDir, resolveTarget, idx) {
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

  // Position from <p:spPr><a:xfrm>...</a:xfrm></p:spPr>
  const xfrm = pPic['p:spPr'] && pPic['p:spPr']['a:xfrm'];
  if (xfrm) {
    if (xfrm['a:off']) {
      media.position = {
        x: emuToPx(xfrm['a:off']['@_x']) || 0,
        y: emuToPx(xfrm['a:off']['@_y']) || 0,
      };
    }
    if (xfrm['a:ext']) {
      media.width = emuToPx(xfrm['a:ext']['@_cx']) || 0;
      media.height = emuToPx(xfrm['a:ext']['@_cy']) || 0;
    }
    if (xfrm['@_rot']) {
      const deg = pptxRotationToDegrees(xfrm['@_rot']);
      if (deg !== 0) media.rotation = deg;
    }
  }

  return media;
}

/**
 * Walk a tree node looking for <p:pic> elements at any depth.
 * PPTX shape trees can be nested via groups (<p:grpSp>); for Sprint 1 we
 * recurse into groups but flatten the result. Sprint 2 will preserve groups.
 */
function findAllPictures(node, out = []) {
  if (!node || typeof node !== 'object') return out;
  if (node['p:pic']) {
    for (const pic of asArray(node['p:pic'])) out.push(pic);
  }
  if (node['p:grpSp']) {
    for (const grp of asArray(node['p:grpSp'])) findAllPictures(grp, out);
  }
  return out;
}

module.exports = { pictureToMedia, findAllPictures };