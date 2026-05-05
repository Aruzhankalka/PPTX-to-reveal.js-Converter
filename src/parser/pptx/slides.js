const { readText } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');

/**
 * Read presentation.xml + its .rels to determine the ordered list of slide files.
 *
 * Returns an array of { rId, path } in slide order, where path is e.g.
 * 'ppt/slides/slide1.xml'. The order comes from <p:sldIdLst> in presentation.xml,
 * NOT from filename sort — slide7.xml may appear before slide2.xml if the user
 * reordered slides in PowerPoint.
 */
async function listSlides(zip) {
  const presentationXml = await readText(zip, 'ppt/presentation.xml');
  if (!presentationXml) {
    const e = new Error('Missing ppt/presentation.xml — not a valid PPTX');
    e.code = 'INVALID_PPTX';
    throw e;
  }

  const presRelsXml = await readText(zip, 'ppt/_rels/presentation.xml.rels');
  const presRels = parseRelationships(presRelsXml);

  const parsed = parseXml(presentationXml);
  const sldIdLst = parsed
    && parsed['p:presentation']
    && parsed['p:presentation']['p:sldIdLst'];
  if (!sldIdLst) return [];

  const result = [];
  for (const sldId of asArray(sldIdLst['p:sldId'])) {
    const rId = sldId['@_r:id'];
    if (!rId) continue;
    const rel = presRels[rId];
    if (!rel) continue;
    const path = resolveTarget('ppt', rel.target);
    result.push({ rId, path });
  }
  return result;
}

module.exports = { listSlides };