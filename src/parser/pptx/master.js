/**
 * Slide-master parser (FR-11/FR-12) — resolves the master's own relationship
 * links to its theme and slide layouts, then builds the spec's
 * layouts[].placeholders[] array from each layout's placeholder shapes.
 * The single entry point into the theme (parseTheme) and the layout list
 * that layouts.js's per-slide lookups are built from.
 */

const { readText } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { parseTheme } = require('./theme');
const { extractXfrm } = require('./layouts');
const { resolveSolidFillCss } = require('./color');
const { lvl1pPrToFormatting } = require('./slides');
const { emuToPx } = require('./units');

// OOXML relationship type suffixes (full URIs are long; we match the tail)
const TYPE_SLIDE_MASTER = '/slidemaster';
const TYPE_SLIDE_LAYOUT = '/slidelayout';
const TYPE_THEME        = '/theme';

function typeEndsWith(rel, suffix) {
  return rel.type.toLowerCase().endsWith(suffix);
}

// OOXML placeholder <p:ph type="..."> -> spec's simplified role enum.
// Absent type, or any value not listed here (e.g. a generic "obj" content
// placeholder that can hold any media), defaults to 'body' — there is no
// closer fit among the spec's six roles for a generic placeholder.
const PH_TYPE_TO_ROLE = {
  title: 'title', ctrTitle: 'title',
  subTitle: 'subtitle',
  ftr: 'footer',
  dt: 'date',
  sldNum: 'slide-number',
};

// OOXML placeholder type -> spec's simplified type enum (text|image|video|table|other).
// OOXML doesn't distinguish audio/video at the placeholder-type level (both
// use "media"), so media placeholders are conservatively mapped to 'other'
// rather than guessing 'video'.
const PH_TYPE_TO_TYPE = {
  pic: 'image', clipArt: 'image',
  tbl: 'table',
  chart: 'other', dgm: 'other', media: 'other',
};

/**
 * Build the spec's layouts[].placeholders[] array from a layout's own
 * <p:cSld><p:spTree> — one entry per placeholder shape, with position/size,
 * padding, background, type/role, and default text formatting.
 *
 * Reuses layouts.js's extractXfrm (the same placeholder-geometry extraction
 * already used to resolve slide-level inheritance) and slides.js's
 * lvl1pPrToFormatting (the same <a:lstStyle><a:lvl1pPr> shape as
 * presentation.xml's <p:defaultTextStyle>, already used for master.formatting).
 *
 * @param {object|null} spTree - parsed <p:spTree> from a layout XML
 * @returns {object[]} placeholders array (never null, may be empty)
 */
function buildLayoutPlaceholders(spTree) {
  if (!spTree) return [];

  const placeholders = [];

  for (const sp of asArray(spTree['p:sp'])) {
    const ph = sp['p:nvSpPr']
      && sp['p:nvSpPr']['p:nvPr']
      && sp['p:nvSpPr']['p:nvPr']['p:ph'];
    if (!ph) continue; // non-placeholder shapes are decoration, not layout slots

    const phType = ph['@_type'] || null;
    const phIdx  = ph['@_idx'] !== undefined ? Number(ph['@_idx']) : 0;

    const placeholder = {
      'placeholder-id': `ph-${phType || 'body'}-${phIdx}`,
      type: PH_TYPE_TO_TYPE[phType] || 'text',
      role: PH_TYPE_TO_ROLE[phType] || 'body',
    };

    // Position/size — absent (not an error) when the placeholder has no
    // explicit <a:xfrm> of its own and inherits geometry from the master.
    const geo = extractXfrm(sp);
    if (geo) {
      placeholder.position = geo.position;
      if (geo.width  != null) placeholder.width  = geo.width;
      if (geo.height != null) placeholder.height = geo.height;
    }

    const txBody = sp['p:txBody'];
    const bodyPr = txBody && txBody['a:bodyPr'];
    if (bodyPr) {
      const l = bodyPr['@_lIns'] != null ? emuToPx(bodyPr['@_lIns']) : null;
      const r = bodyPr['@_rIns'] != null ? emuToPx(bodyPr['@_rIns']) : null;
      const t = bodyPr['@_tIns'] != null ? emuToPx(bodyPr['@_tIns']) : null;
      const b = bodyPr['@_bIns'] != null ? emuToPx(bodyPr['@_bIns']) : null;
      if (l != null || r != null || t != null || b != null) {
        placeholder.padding = `${t ?? 0}px ${r ?? 0}px ${b ?? 0}px ${l ?? 0}px`;
      }
    }

    const background = resolveSolidFillCss(sp['p:spPr'] && sp['p:spPr']['a:solidFill']);
    if (background) placeholder.background = background;

    const lstStyle = txBody && txBody['a:lstStyle'];
    const formatting = lvl1pPrToFormatting(lstStyle && lstStyle['a:lvl1pPr']);
    if (formatting) placeholder.formatting = formatting;

    placeholders.push(placeholder);
  }

  return placeholders;
}

/**
 * Parse a single slide layout XML into an IR layout entry.
 * Exported for unit testing — no ZIP I/O.
 *
 * @param {string} xmlString
 * @param {string} layoutPath - ZIP path used as the stable layout-id
 * @returns {{ 'layout-id': string, name: string|null, type: string|null, placeholders: object[] }}
 */
function parseLayoutXml(xmlString, layoutPath) {
  const parsed = parseXml(xmlString);
  const root = parsed && parsed['p:sldLayout'];

  const name   = (root && root['p:cSld'] && root['p:cSld']['@_name']) || null;
  const type   = (root && root['@_type']) || null;
  const spTree = root && root['p:cSld'] && root['p:cSld']['p:spTree'];

  return { 'layout-id': layoutPath, name, type, placeholders: buildLayoutPlaceholders(spTree) };
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
 *   layouts: Array<{ 'layout-id': string, name: string|null, type: string|null, placeholders: object[] }>
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

      // Font size, BIU, and latin font family from <a:defRPr>
      const defRPr = lvlPPr['a:defRPr'];
      if (defRPr) {
        if (defRPr['@_sz']) {
          const pt = Number(defRPr['@_sz']) / 100;
          if (!Number.isNaN(pt)) entry.size = `${pt}pt`;
        }
        const b = defRPr['@_b'];
        if (b === '1' || b === 'true') entry.bold = true;
        const _i = defRPr['@_i'];
        if (_i === '1' || _i === 'true') entry.italic = true;
        const u = defRPr['@_u'];
        if (u && u !== 'none') entry.underline = true;
        // Latin font family — propagated as fallback to runs that have no explicit <a:latin>.
        const latinFont = defRPr['a:latin']?.['@_typeface'];
        if (latinFont) entry.font = latinFont;
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

module.exports = { parseMaster, parseLayoutXml, buildLayoutPlaceholders };
