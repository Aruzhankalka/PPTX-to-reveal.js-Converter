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