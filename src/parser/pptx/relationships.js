/**
 * OOXML relationship (.rels) resolver — every part that references another
 * part by id (slide->layout, slide->media, presentation->slides, etc.) goes
 * through parseRelationships + resolveTarget here. Used throughout the
 * parser (slides.js, layouts.js, master.js, media.js, fonts.js).
 */

const { parseXml, asArray } = require('./xml');

/**
 * Parse a .rels file into a map: { rId1: { type, target }, ... }
 *
 * Example .rels content:
 *   <Relationships ...>
 *     <Relationship Id="rId1"
 *                   Type=".../slide"
 *                   Target="slides/slide1.xml"/>
 *   </Relationships>
 *
 * @param {string|null} relsXml - raw .rels XML, or falsy if the part is absent
 * @returns {Record<string, {type: string, target: string}>} map keyed by rId,
 *   empty object when relsXml is absent or has no root element
 */
function parseRelationships(relsXml) {
  if (!relsXml) return {};
  const parsed = parseXml(relsXml);
  const root = parsed && parsed.Relationships;
  if (!root) return {};

  const out = {};
  for (const r of asArray(root.Relationship)) {
    const id = r['@_Id'];
    if (!id) continue;
    out[id] = {
      type: r['@_Type'] || '',
      target: r['@_Target'] || '',
    };
  }
  return out;
}

/**
 * Resolve a relationship target relative to a base directory.
 *
 * In PPTX, .rels targets are relative to the directory of the part that owns
 * the .rels file. For example, ppt/slides/_rels/slide1.xml.rels has targets
 * relative to ppt/slides/, so "../media/image1.png" resolves to ppt/media/image1.png.
 *
 * @param {string} baseDir - directory of the part that owns the .rels file
 *   (e.g. 'ppt/slides'), no trailing slash
 * @param {string} target - the relationship's Target attribute, possibly
 *   relative ("../media/image1.png") or absolute ("/media/image1.png")
 * @returns {string} zip-relative path with no leading slash (e.g. 'ppt/media/image1.png')
 */
function resolveTarget(baseDir, target) {
  if (!target) return '';
  // Absolute targets start with '/'; rare in PPTX, but possible.
  if (target.startsWith('/')) return target.slice(1);

  const baseParts = baseDir.split('/').filter(Boolean);
  const targetParts = target.split('/');
  for (const part of targetParts) {
    if (part === '..') baseParts.pop();
    else if (part !== '.') baseParts.push(part);
  }
  return baseParts.join('/');
}

module.exports = { parseRelationships, resolveTarget };