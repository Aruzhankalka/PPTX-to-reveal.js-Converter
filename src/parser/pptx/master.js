const { readText } = require('./zip');
const { parseXml } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { parseTheme } = require('./theme');

// OOXML relationship type suffixes (full URIs are long; we match the tail)
const TYPE_SLIDE_MASTER = '/slidemaster';
const TYPE_SLIDE_LAYOUT = '/slidelayout';
const TYPE_THEME        = '/theme';

function typeEndsWith(rel, suffix) {
  return rel.type.toLowerCase().endsWith(suffix);
}

/**
 * Parse a single slide layout XML into an IR layout entry.
 * Exported for unit testing — no ZIP I/O.
 *
 * @param {string} xmlString
 * @param {string} layoutPath - ZIP path used as the stable id
 * @returns {{ id: string, name: string|null, type: string|null }}
 */
function parseLayoutXml(xmlString, layoutPath) {
  const parsed = parseXml(xmlString);
  const root = parsed && parsed['p:sldLayout'];

  const name = (root && root['p:cSld'] && root['p:cSld']['@_name']) || null;
  const type = (root && root['@_type']) || null;

  return { id: layoutPath, name, type };
}

/**
 * Parse the first slide master in the PPTX: discovers its theme path (feeding
 * FR-12) and enumerates all slide layouts (FR-11).
 *
 * Multiple masters exist only in rare merged presentations; Sprint 2 processes
 * the first one, which is the active master for the vast majority of files.
 *
 * @param {JSZip} zip
 * @returns {Promise<{
 *   theme: object|null,
 *   masterName: string|null,
 *   layouts: Array<{ id: string, name: string|null, type: string|null }>
 * } | null>}  null when no slide master relationship is found.
 */
async function parseMaster(zip) {
  // 1. Find the master path via presentation.xml.rels
  const presRelsXml = await readText(zip, 'ppt/_rels/presentation.xml.rels');
  const presRels = parseRelationships(presRelsXml);

  const masterRel = Object.values(presRels).find((r) => typeEndsWith(r, TYPE_SLIDE_MASTER));
  if (!masterRel) return null;

  const masterPath = resolveTarget('ppt', masterRel.target);
  const masterDir  = masterPath.substring(0, masterPath.lastIndexOf('/'));
  const masterFile = masterPath.substring(masterPath.lastIndexOf('/') + 1);
  const masterRelsPath = `${masterDir}/_rels/${masterFile}.rels`;

  // 2. Parse the master XML and its relationships
  const [masterXml, masterRelsXml] = await Promise.all([
    readText(zip, masterPath),
    readText(zip, masterRelsPath),
  ]);
  const masterRels = parseRelationships(masterRelsXml);

  // 3. Resolve the theme (FR-12) via the master's rels, not the fallback guess
  const themeRel  = Object.values(masterRels).find((r) => typeEndsWith(r, TYPE_THEME));
  const themePath = themeRel ? resolveTarget(masterDir, themeRel.target) : undefined;
  const theme = await parseTheme(zip, themePath);

  // 4. Collect layout paths in rels-file order (creation order in PowerPoint)
  const layoutRels = Object.values(masterRels).filter((r) => typeEndsWith(r, TYPE_SLIDE_LAYOUT));

  // 5. Parse each layout file — skip silently if the entry is missing in the ZIP
  const layouts = [];
  for (const rel of layoutRels) {
    const layoutPath = resolveTarget(masterDir, rel.target);
    const layoutXml  = await readText(zip, layoutPath);
    if (!layoutXml) continue;
    layouts.push(parseLayoutXml(layoutXml, layoutPath));
  }

  // 6. Extract master name and txStyles from the master XML
  const masterParsed = parseXml(masterXml);
  const masterRoot   = masterParsed && masterParsed['p:sldMaster'];
  const masterName   = (masterRoot && masterRoot['p:cSld'] && masterRoot['p:cSld']['@_name']) || null;
  const txStyles     = parseTxStyles(masterRoot);

  return { theme, masterName, layouts, txStyles };
}

/**
 * Parse <p:txStyles> from a slide master root node.
 * Returns per-level entries with size, lineSpacing, spaceBefore, spaceAfter
 * for title, body, and other text.
 *
 * @param {object|null} masterRoot - parsed p:sldMaster node
 * @returns {{ title: object, body: object, other: object } | null}
 */
function parseTxStyles(masterRoot) {
  const txStyles = masterRoot && masterRoot['p:txStyles'];
  if (!txStyles) return null;

  function parseLevelEntries(styleNode) {
    if (!styleNode) return {};
    const entries = {};
    for (let i = 1; i <= 9; i++) {
      const lvlPPr = styleNode[`a:lvl${i}pPr`];
      if (!lvlPPr) continue;
      const entry = {};

      // Font size from <a:defRPr sz="..."> (sz in 100ths of a point)
      const defRPr = lvlPPr['a:defRPr'];
      if (defRPr && defRPr['@_sz']) {
        const pt = Number(defRPr['@_sz']) / 100;
        if (!Number.isNaN(pt)) entry.size = `${pt}pt`;
      }

      // Line spacing from <a:lnSpc> — spcPct val in 1000ths of a percent,
      // spcPts val in 100ths of a point.
      const lnSpc = lvlPPr['a:lnSpc'];
      if (lnSpc) {
        const pct = lnSpc['a:spcPct'];
        const pts = lnSpc['a:spcPts'];
        if (pct && pct['@_val']) {
          const v = Number(pct['@_val']) / 100000;
          if (!Number.isNaN(v)) entry.lineSpacing = String(v);
        } else if (pts && pts['@_val']) {
          const v = Number(pts['@_val']) / 100;
          if (!Number.isNaN(v)) entry.lineSpacing = `${v}pt`;
        }
      }

      // Space before / after from <a:spcBef> / <a:spcAft>
      for (const [attr, key] of [['a:spcBef', 'spaceBefore'], ['a:spcAft', 'spaceAfter']]) {
        const node = lvlPPr[attr];
        if (!node) continue;
        const ptsNode = node['a:spcPts'];
        const pctNode = node['a:spcPct'];
        if (ptsNode && ptsNode['@_val'] != null) {
          const v = Number(ptsNode['@_val']) / 100;
          if (!Number.isNaN(v)) entry[key] = `${v}pt`;
        } else if (pctNode && pctNode['@_val'] != null) {
          const v = Number(pctNode['@_val']) / 100000;
          if (!Number.isNaN(v)) entry[key] = `${v}em`;
        }
      }

      if (Object.keys(entry).length > 0) entries[i] = entry;
    }
    return entries;
  }

  return {
    title: parseLevelEntries(txStyles['p:titleStyle']),
    body:  parseLevelEntries(txStyles['p:bodyStyle']),
    other: parseLevelEntries(txStyles['p:otherStyle']),
  };
}

module.exports = { parseMaster, parseLayoutXml };
