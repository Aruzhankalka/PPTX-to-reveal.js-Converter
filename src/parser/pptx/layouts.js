'use strict';

const { readText } = require('./zip');
const { parseXml, asArray } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { emuToPx, pptxRotationToDegrees } = require('./units');
const { parseShapes } = require('./shapes');

// OOXML relationship type suffix for the slide master link inside a layout .rels
const TYPE_SLIDE_MASTER = '/slidemaster';

// OOXML scheme-color aliases: tx1/tx2 map to dk1/dk2, bg1/bg2 map to lt1/lt2.
const SCHEME_ALIAS = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };

// ---------------------------------------------------------------------------
// Internal helpers
// ---------------------------------------------------------------------------

/**
 * Extract position/size/rotation geometry from the <a:xfrm> of a <p:sp> node.
 * Returns null when the shape has no <a:xfrm> with both off+ext (the shape
 * inherits its position from the next level up the inheritance chain).
 *
 * Coordinates are converted from EMU to px (same factor as units.js:
 * 1 px = 9525 EMU) so the result can be applied directly to the IR block.
 *
 * @param {object} sp  parsed <p:sp> node
 * @returns {{ position:{x,y}, width?:number, height?:number, rotation?:number } | null}
 */
function extractXfrm(sp) {
  const xfrm = sp['p:spPr'] && sp['p:spPr']['a:xfrm'];
  if (!xfrm) return null;

  const off = xfrm['a:off'];
  const ext = xfrm['a:ext'];
  if (!off || !ext) return null;

  const x = emuToPx(off['@_x']);
  const y = emuToPx(off['@_y']);
  if (x == null || y == null) return null; // position must be present

  const geo = { position: { x: x || 0, y: y || 0 } };

  const w = emuToPx(ext['@_cx']);
  const h = emuToPx(ext['@_cy']);
  if (w != null) geo.width  = w;
  if (h != null) geo.height = h;

  if (xfrm['@_rot']) {
    const deg = pptxRotationToDegrees(xfrm['@_rot']);
    if (deg !== 0) geo.rotation = deg;
  }

  return geo;
}

/**
 * Walk an spTree node and index every placeholder that has an explicit <a:xfrm>.
 *
 * OOXML §19.2.1.27: the default value of <p:ph idx> is 0 when the attribute
 * is absent (commonly seen on title placeholders).
 *
 * @param {object} spTree  parsed <p:spTree> node
 * @returns {{ byIdx: Map<number,geo>, byType: Map<string,geo> }}
 */
function collectPlaceholders(spTree) {
  const byIdx  = new Map();
  const byType = new Map();

  for (const sp of asArray(spTree['p:sp'])) {
    const ph = sp['p:nvSpPr']
      && sp['p:nvSpPr']['p:nvPr']
      && sp['p:nvSpPr']['p:nvPr']['p:ph'];
    if (!ph) continue;

    const geo = extractXfrm(sp);
    if (!geo) continue; // placeholder inherits position — skip; resolved at a higher level

    // Also capture the text body's vertical anchor so the generator can apply
    // the correct CSS (flex centering for anchor="ctr", flex-end for anchor="b").
    const txBody  = sp['p:txBody'];
    const bodyPr  = txBody && txBody['a:bodyPr'];
    const anchor  = bodyPr && bodyPr['@_anchor'];
    if (anchor) geo.textAnchor = anchor;

    // Capture normAutofit from the layout/master bodyPr so slide shapes that
    // carry no explicit <a:normAutofit> in their own txBody can inherit the
    // shrink parameters from the template.
    const normAutofit = bodyPr && bodyPr['a:normAutofit'];
    if (normAutofit) {
      const fontScaleRaw = normAutofit['@_fontScale'];
      const lnSpcRedRaw  = normAutofit['@_lnSpcReduction'];
      const fontScale = fontScaleRaw != null ? Number(fontScaleRaw) / 100000 : 1;
      const lnSpcRed  = lnSpcRedRaw  != null ? Number(lnSpcRedRaw)  / 100000 : 0;
      if (fontScale !== 1 || lnSpcRed !== 0) {
        geo.normAutofit = { fontScale, lnSpcRed };
      }
    }

    // Capture per-level formatting from this placeholder's lstStyle.
    // Size and color are applied in slide.js post-processing; the full lstStyle
    // node is stored in geo.lstStyle so paragraphToIr can use it as an
    // intermediate cascade level between the slide shape's own lstStyle and the
    // master txStyles (OOXML inheritance order: slide → layout → master).
    const lstStyle = txBody && txBody['a:lstStyle'];
    const lvl1pPr  = lstStyle && lstStyle['a:lvl1pPr'];
    const defRPr   = lvl1pPr  && lvl1pPr['a:defRPr'];
    if (defRPr && defRPr['@_sz']) {
      const pt = Number(defRPr['@_sz']) / 100;
      if (!Number.isNaN(pt)) geo.defaultFontSize = `${pt}pt`;
    }

    // Capture the default text color from this placeholder's lstStyle so that
    // special placeholders (ftr, sldNum, dt) carry an explicit color and are
    // not invisible when the theme defines text color via tx1/dk1 alias.
    if (defRPr) {
      const fill = defRPr['a:solidFill'];
      if (fill) {
        if (fill['a:srgbClr'] && fill['a:srgbClr']['@_val']) {
          geo.defaultColor = '#' + fill['a:srgbClr']['@_val'];
        } else if (fill['a:schemeClr'] && fill['a:schemeClr']['@_val']) {
          const raw = fill['a:schemeClr']['@_val'];
          geo.defaultColor = `var(--theme-${SCHEME_ALIAS[raw] || raw})`;
        }
      }
    }

    // Store the full lstStyle node so slide.js can pass it to shapeToTextBlock
    // as the layout-level BIU fallback (bold/italic/underline per indent level).
    if (lstStyle && typeof lstStyle === 'object' && Object.keys(lstStyle).length > 0) {
      geo.lstStyle = lstStyle;
    }

    const idx  = ph['@_idx'] !== undefined ? Number(ph['@_idx']) : 0;
    const type = ph['@_type'] || null;

    // First-wins: layout entries should shadow master entries, so callers must
    // add layout results before master results.
    if (!byIdx.has(idx))          byIdx.set(idx, geo);
    if (type && !byType.has(type)) byType.set(type, geo);
  }

  return { byIdx, byType };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Load placeholder geometry through the two-level inheritance chain:
 *   slide layout XML → slide master XML (fallback for slots absent from layout).
 *
 * Returns a merged lookup with two indexes (byIdx / byType).  Layout entries
 * take precedence: master entries are only inserted for keys not already set.
 *
 * The result's position/size values are already in CSS pixels (converted from
 * EMU) so they can be written directly onto IR block objects.
 *
 * @param {JSZip}        zip         open PPTX archive
 * @param {string|null}  layoutPath  ZIP path to the slide layout, e.g.
 *                                   'ppt/slideLayouts/slideLayout3.xml'
 * @returns {Promise<{ byIdx: Map<number,geo>, byType: Map<string,geo> }>}
 */
async function loadLayoutGeometry(zip, layoutPath) {
  const byIdx  = new Map();
  const byType = new Map();

  if (!layoutPath) return { byIdx, byType };

  // ---- Slide layout --------------------------------------------------------

  const layoutXml = await readText(zip, layoutPath);
  if (!layoutXml) return { byIdx, byType };

  const layoutParsed  = parseXml(layoutXml);
  const layoutSpTree  = layoutParsed
    && layoutParsed['p:sldLayout']
    && layoutParsed['p:sldLayout']['p:cSld']
    && layoutParsed['p:sldLayout']['p:cSld']['p:spTree'];

  if (layoutSpTree) {
    const { byIdx: lIdx, byType: lType } = collectPlaceholders(layoutSpTree);
    for (const [k, v] of lIdx)  byIdx.set(k, v);
    for (const [k, v] of lType) byType.set(k, v);
  }

  // ---- Slide master (fills gaps left by the layout) ------------------------

  const layoutDir  = layoutPath.substring(0, layoutPath.lastIndexOf('/'));
  const layoutFile = layoutPath.substring(layoutPath.lastIndexOf('/') + 1);
  const layoutRelsPath = `${layoutDir}/_rels/${layoutFile}.rels`;

  const layoutRelsXml = await readText(zip, layoutRelsPath);
  if (!layoutRelsXml) return { byIdx, byType };

  const layoutRels = parseRelationships(layoutRelsXml);
  const masterRel  = Object.values(layoutRels).find((r) =>
    r.type.toLowerCase().endsWith(TYPE_SLIDE_MASTER)
  );
  if (!masterRel) return { byIdx, byType };

  const masterPath = resolveTarget(layoutDir, masterRel.target);
  const masterXml  = await readText(zip, masterPath);
  if (!masterXml) return { byIdx, byType };

  const masterParsed = parseXml(masterXml);
  const masterSpTree = masterParsed
    && masterParsed['p:sldMaster']
    && masterParsed['p:sldMaster']['p:cSld']
    && masterParsed['p:sldMaster']['p:cSld']['p:spTree'];

  if (masterSpTree) {
    const { byIdx: mIdx, byType: mType } = collectPlaceholders(masterSpTree);
    // Master fills in only what the layout didn't provide.
    for (const [k, v] of mIdx)  if (!byIdx.has(k))  byIdx.set(k, v);
    for (const [k, v] of mType) if (!byType.has(k)) byType.set(k, v);
  }

  return { byIdx, byType };
}

/**
 * Look up placeholder geometry by idx (primary) then type (fallback).
 *
 * @param {{ byIdx: Map<number,geo>, byType: Map<string,geo> }} geoMap
 * @param {number}      phIdx   placeholder idx (use 0 when absent in XML)
 * @param {string|null} phType  placeholder type attribute value, or null
 * @returns {geo | null}
 */
function lookupGeo(geoMap, phIdx, phType) {
  const byIdxResult = geoMap.byIdx.get(phIdx);
  if (byIdxResult !== undefined) return byIdxResult;
  if (phType) {
    const byTypeResult = geoMap.byType.get(phType);
    if (byTypeResult !== undefined) return byTypeResult;
  }
  return null;
}

/**
 * Convert a <p:pic> node plus its rels to a minimal media-like object.
 * Reuses extractXfrm because <p:pic> has the same <p:spPr><a:xfrm> structure.
 *
 * @param {object} pPic  parsed <p:pic> node
 * @param {object} rels  relationships keyed by rId
 * @param {string} dir   directory of the XML file (for resolveTarget)
 * @returns {{ 'file-link', 'media-type', position, width, height } | null}
 */
function picToLayoutMedia(pPic, rels, dir) {
  const blipFill = pPic['p:blipFill'];
  if (!blipFill || !blipFill['a:blip']) return null;
  const rId = blipFill['a:blip']['@_r:embed'];
  if (!rId) return null;
  const rel = rels[rId];
  if (!rel) return null;

  const zipPath = resolveTarget(dir, rel.target);
  const geo     = extractXfrm(pPic);

  return {
    'file-link':  zipPath,
    'media-type': 'image',
    position: geo ? geo.position          : { x: 0, y: 0 },
    width:    geo ? (geo.width  || 0)     : 0,
    height:   geo ? (geo.height || 0)     : 0,
    ...(geo && geo.rotation != null ? { rotation: geo.rotation } : {}),
  };
}

/**
 * Recursively collect all <p:pic> nodes from an spTree, including those inside
 * <p:grpSp> group containers, resolving each to a media object.
 *
 * @param {object}   tree  parsed spTree or grpSp node
 * @param {object}   rels  relationships keyed by rId
 * @param {string}   dir   directory of the XML file (for resolveTarget)
 * @param {object[]} out   accumulator — items are pushed here
 */
function collectPicsFromTree(tree, rels, dir, out) {
  for (const pic of asArray(tree['p:pic'])) {
    const m = picToLayoutMedia(pic, rels, dir);
    if (m) out.push(m);
  }
  for (const grp of asArray(tree['p:grpSp'])) {
    collectPicsFromTree(grp, rels, dir, out);
  }
}

/**
 * Collect <p:pic> images from the slide layout and its master so they can be
 * injected into the slide's media list (logos, decorative backgrounds, etc.).
 *
 * Layout pics take precedence in z-order over master pics; both use the SAME
 * EMU→px conversion as slide geometry so they slot into the existing renderer.
 *
 * @param {JSZip}        zip         open PPTX archive
 * @param {string|null}  layoutPath  ZIP path to the slide layout
 * @returns {Promise<{ layoutMedia: object[], masterMedia: object[] }>}
 */
async function collectLayoutMedia(zip, layoutPath) {
  const result = { layoutMedia: [], masterMedia: [] };
  if (!layoutPath) return result;

  // ---- Layout pics ----------------------------------------------------------

  const layoutXml = await readText(zip, layoutPath);
  if (!layoutXml) return result;

  const layoutParsed = parseXml(layoutXml);
  const layoutSpTree = layoutParsed
    && layoutParsed['p:sldLayout']
    && layoutParsed['p:sldLayout']['p:cSld']
    && layoutParsed['p:sldLayout']['p:cSld']['p:spTree'];

  const layoutDir  = layoutPath.substring(0, layoutPath.lastIndexOf('/'));
  const layoutFile = layoutPath.substring(layoutPath.lastIndexOf('/') + 1);
  const layoutRelsPath = `${layoutDir}/_rels/${layoutFile}.rels`;
  const layoutRelsXml  = await readText(zip, layoutRelsPath);
  const layoutRels     = layoutRelsXml ? parseRelationships(layoutRelsXml) : {};

  if (layoutSpTree) {
    collectPicsFromTree(layoutSpTree, layoutRels, layoutDir, result.layoutMedia);
  }

  // ---- Master pics ----------------------------------------------------------

  if (!layoutRelsXml) return result;
  const masterRel = Object.values(parseRelationships(layoutRelsXml)).find((r) =>
    r.type.toLowerCase().endsWith(TYPE_SLIDE_MASTER)
  );
  if (!masterRel) return result;

  const masterPath   = resolveTarget(layoutDir, masterRel.target);
  const masterXml    = await readText(zip, masterPath);
  if (!masterXml) return result;

  const masterParsed = parseXml(masterXml);
  const masterSpTree = masterParsed
    && masterParsed['p:sldMaster']
    && masterParsed['p:sldMaster']['p:cSld']
    && masterParsed['p:sldMaster']['p:cSld']['p:spTree'];

  const masterDir  = masterPath.substring(0, masterPath.lastIndexOf('/'));
  const masterFile = masterPath.substring(masterPath.lastIndexOf('/') + 1);
  const masterRelsPath = `${masterDir}/_rels/${masterFile}.rels`;
  const masterRelsXml  = await readText(zip, masterRelsPath);
  const masterRels     = masterRelsXml ? parseRelationships(masterRelsXml) : {};

  if (masterSpTree) {
    collectPicsFromTree(masterSpTree, masterRels, masterDir, result.masterMedia);
  }

  // Deduplicate by file path: if the layout defines the same image as the master
  // (e.g. a logo repositioned per layout), the layout's copy wins and the master's
  // copy is dropped.  Position proximity is NOT used — the layout may intentionally
  // place the same image at a completely different position than the master.
  const layoutFiles = new Set(result.layoutMedia.map(m => m['file-link']));
  result.masterMedia = result.masterMedia.filter(mp => !layoutFiles.has(mp['file-link']));

  return result;
}

/**
 * Collect non-placeholder <p:sp> and <p:cxnSp> shapes from the slide layout
 * and its master so they can be injected into the slide's shape list as
 * background / decorative elements (colored blocks, lines, branding shapes).
 *
 * @param {JSZip}        zip         open PPTX archive
 * @param {string|null}  layoutPath  ZIP path to the slide layout
 * @returns {Promise<{ layoutShapes: object[], masterShapes: object[] }>}
 */
async function collectLayoutShapes(zip, layoutPath) {
  const result = { layoutShapes: [], masterShapes: [] };
  if (!layoutPath) return result;

  const warnings = [];

  // ---- Layout shapes -------------------------------------------------------

  const layoutXml = await readText(zip, layoutPath);
  if (!layoutXml) return result;

  const layoutParsed = parseXml(layoutXml);
  const layoutSpTree = layoutParsed
    && layoutParsed['p:sldLayout']
    && layoutParsed['p:sldLayout']['p:cSld']
    && layoutParsed['p:sldLayout']['p:cSld']['p:spTree'];

  if (layoutSpTree) {
    result.layoutShapes = parseShapes(layoutSpTree, null, warnings);
  }

  // ---- Master shapes -------------------------------------------------------

  const layoutDir  = layoutPath.substring(0, layoutPath.lastIndexOf('/'));
  const layoutFile = layoutPath.substring(layoutPath.lastIndexOf('/') + 1);
  const layoutRelsPath = `${layoutDir}/_rels/${layoutFile}.rels`;
  const layoutRelsXml  = await readText(zip, layoutRelsPath);
  if (!layoutRelsXml) return result;

  const masterRel = Object.values(parseRelationships(layoutRelsXml)).find((r) =>
    r.type.toLowerCase().endsWith(TYPE_SLIDE_MASTER)
  );
  if (!masterRel) return result;

  const masterPath = resolveTarget(layoutDir, masterRel.target);
  const masterXml  = await readText(zip, masterPath);
  if (!masterXml) return result;

  const masterParsed = parseXml(masterXml);
  const masterSpTree = masterParsed
    && masterParsed['p:sldMaster']
    && masterParsed['p:sldMaster']['p:cSld']
    && masterParsed['p:sldMaster']['p:cSld']['p:spTree'];

  if (masterSpTree) {
    result.masterShapes = parseShapes(masterSpTree, null, warnings);
  }

  return result;
}

module.exports = { loadLayoutGeometry, lookupGeo, collectPlaceholders, extractXfrm, collectLayoutMedia, collectLayoutShapes };
