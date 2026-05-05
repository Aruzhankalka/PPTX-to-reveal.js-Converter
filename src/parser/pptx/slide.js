const { readText } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { shapeToTextBlock } = require('./text');
const { pictureToMedia, findAllPictures } = require('./media');

/**
 * Parse a single slide XML into an IR slide.
 *
 * @param {JSZip} zip - the open .pptx archive
 * @param {string} slidePath - e.g. 'ppt/slides/slide1.xml'
 * @returns {Promise<{ ir: object, mediaRefs: Array<{src, dest}> }>}
 *   ir       - the IR slide object
 *   mediaRefs - list of {src: in-zip path, dest: bundle-relative path}
 *               so the caller can extract the actual image bytes later.
 */
async function parseSlide(zip, slidePath) {
  const slideXml = await readText(zip, slidePath);
  if (!slideXml) {
    return { ir: { contents: { text: [], media: [] } }, mediaRefs: [] };
  }

  // Slide directory is everything before the filename, e.g. 'ppt/slides'
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
  const filename = slidePath.substring(slidePath.lastIndexOf('/') + 1);
  const relsPath = `${slideDir}/_rels/${filename}.rels`;

  const relsXml = await readText(zip, relsPath);
  const slideRels = parseRelationships(relsXml);

  const parsed = parseXml(slideXml);
  // Slide root is <p:sld><p:cSld><p:spTree>...</p:spTree></p:cSld></p:sld>
  const spTree = parsed
    && parsed['p:sld']
    && parsed['p:sld']['p:cSld']
    && parsed['p:sld']['p:cSld']['p:spTree'];

  if (!spTree) {
    return { ir: { contents: { text: [], media: [] } }, mediaRefs: [] };
  }

  // -- Extract text blocks from <p:sp> shapes --
  const textBlocks = [];
  let textIdx = 0;
  for (const sp of asArray(spTree['p:sp'])) {
    const block = shapeToTextBlock(sp, textIdx++);
    if (block) textBlocks.push(block);
  }

  // -- Extract media from <p:pic> shapes (including inside groups) --
  const mediaItems = [];
  const mediaRefs = [];
  let picIdx = 0;
  for (const pic of findAllPictures(spTree)) {
    const media = pictureToMedia(pic, slideRels, slideDir, resolveTarget, picIdx++);
    if (media) {
      mediaItems.push(media);
      // Track the in-zip path so the orchestrator can extract the bytes.
      // 'file-link' in IR is bundle-relative; we keep the original path
      // separately for extraction.
      mediaRefs.push({
        zipPath: media['file-link'],         // e.g. 'ppt/media/image1.png'
        bundlePath: 'media/' + media['file-link'].split('/').pop(),
      });
      // Update the IR media item to use the bundle path
      media['file-link'] = 'media/' + media['file-link'].split('/').pop();
    }
  }

  // -- Find a slide title for the IR --
  // Convention: the first text block whose first paragraph is short is the title.
  // PPTX has a proper "title placeholder" mechanism but it requires layout
  // resolution which is Sprint 2 work. For Sprint 1 we keep this heuristic.
  let title;
  for (const block of textBlocks) {
    const firstPara = block.paragraphs[0];
    if (firstPara && firstPara.runs.length > 0) {
      const text = firstPara.runs.map((r) => r.text).join('');
      if (text.length > 0 && text.length < 100) {
        title = text;
        break;
      }
    }
  }

  const ir = {
    contents: {
      text: textBlocks,
      media: mediaItems,
    },
  };
  if (title) ir.title = title;

  return { ir, mediaRefs };
}

module.exports = { parseSlide };