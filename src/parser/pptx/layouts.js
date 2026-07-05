'use strict';

const { readText } = require('./zip');
const { parseXml, asArray, getSpTreeOrder } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { emuToPx, pptxRotationToDegrees } = require('./units');
const { parseShapes, parsePlaceholderBackgrounds } = require('./shapes');
const { resolveSolidFillCss } = require('./color');

// OOXML relationship type suffix for the slide master link inside a layout .rels
const TYPE_SLIDE_MASTER = '/slidemaster';

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
      const fillColor = resolveSolidFillCss(defRPr['a:solidFill']);
      if (fillColor) geo.defaultColor = fillColor;
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

  if (!layoutPath) return { byIdx, byType, layoutName: null, showMasterSp: true };

  // ---- Slide layout --------------------------------------------------------

  const layoutXml = await readText(zip, layoutPath);
  if (!layoutXml) return { byIdx, byType, layoutName: null, showMasterSp: true };

  const layoutParsed   = parseXml(layoutXml);
  const sldLayoutEl    = layoutParsed && layoutParsed['p:sldLayout'];
  // Prefer explicit name; fall back to OOXML type (e.g. "blank", "title", "body").
  const layoutName     = (sldLayoutEl && (sldLayoutEl['@_name'] || sldLayoutEl['@_type'])) || null;
  // showMasterSp defaults to true; only false when the attribute is literally "0"
  const rawSMS         = sldLayoutEl && sldLayoutEl['@_showMasterSp'];
  const showMasterSp   = rawSMS !== '0' && rawSMS !== false;

  const layoutSpTree  = sldLayoutEl
    && sldLayoutEl['p:cSld']
    && sldLayoutEl['p:cSld']['p:spTree'];

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

  return { byIdx, byType, layoutName, showMasterSp };
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
  const blip = blipFill['a:blip'];

  // Primary: standard r:embed attribute on <a:blip>
  let rId = blip['@_r:embed'];

  // Fallback: Office 2016+ SVG extension — <a:blip><a:extLst><a:ext><asvg:svgBlip r:embed="..."/>
  if (!rId) {
    const extLst = blip['a:extLst'];
    if (extLst) {
      for (const ext of asArray(extLst['a:ext'])) {
        const svgBlip = ext['asvg:svgBlip'];
        if (svgBlip && svgBlip['@_r:embed']) {
          rId = svgBlip['@_r:embed'];
          break;
        }
      }
    }
  }

  if (!rId) return null;
  const rel = rels[rId];
  if (!rel) return null;

  const zipPath = resolveTarget(dir, rel.target);
  const geo     = extractXfrm(pPic);

  const media = {
    'file-link':  zipPath,
    'media-type': 'image',
    position: geo ? geo.position          : { x: 0, y: 0 },
    width:    geo ? (geo.width  || 0)     : 0,
    height:   geo ? (geo.height || 0)     : 0,
    ...(geo && geo.rotation != null ? { rotation: geo.rotation } : {}),
  };

  // Crop from <p:blipFill><a:srcRect l t r b> — values are percentage × 1000.
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

// ---------------------------------------------------------------------------
// Document-order layout/master content collection
// ---------------------------------------------------------------------------

/**
 * Return true when a <p:sp> placeholder has a visible fill that would cause
 * parsePlaceholderBackgrounds to emit a rect shape for it.
 *
 * Mirrors the filter logic in parsePlaceholderBackgrounds:
 *   hasExplicitFillNode ? resolveFill() : resolveStyleFill()
 *   → include when fill.type !== 'none'
 *
 * We only need to detect the common cases (solid/grad explicit fill, or style
 * fillRef with idx > 0 + color child) without calling the full resolver chain.
 */
function phHasVisibleFill(sp) {
  const spPr   = sp['p:spPr'] || {};
  const pStyle = sp['p:style'];

  // Explicit fill nodes
  if (spPr['a:noFill'])    return false;
  if (spPr['a:solidFill'] || spPr['a:gradFill']) return true;
  // pattFill / blipFill / grpFill — resolveFill returns {type:'none'} for these
  if (spPr['a:pattFill'] || spPr['a:blipFill'] || spPr['a:grpFill']) return false;

  // No explicit fill → fall through to style
  const fillRef = pStyle && pStyle['a:fillRef'];
  if (!fillRef) return false;
  if (Number(fillRef['@_idx']) === 0) return false;
  // fillRef has a color child → resolveStyleFill returns solid
  return !!(fillRef['a:srgbClr'] || fillRef['a:schemeClr'] || fillRef['a:sysClr']);
}

/**
 * Process an spTree in document order, interleaving parsed shape objects and
 * media objects as they appear in the XML.
 *
 * Returns an array of { _isMedia:boolean, item:object } in document order.
 * Items from inside <p:grpSp> groups are appended at the end (fast-xml-parser
 * loses cross-type ordering within groups the same way it does at the top level;
 * for the templates tested so far, groups only appear at the top-level anyway).
 *
 * @param {string}   rawXml    raw XML for the layout or master file
 * @param {string}   rootTag   'p:sldLayout' or 'p:sldMaster'
 * @param {object}   spTree    parsed spTree node (from the main parser)
 * @param {object}   rels      parsed rels for this file
 * @param {string}   dir       directory of this file
 * @param {string[]} warnings  mutable warnings array
 * @returns {{ _isMedia:boolean, item:object }[]}
 */
function collectSpTreeOrdered(rawXml, rootTag, spTree, rels, dir, warnings) {
  const { shapes: nonPhShapes } = parseShapes(spTree, null, warnings);
  const phBgShapes  = parsePlaceholderBackgrounds(spTree, warnings);

  // Map each p:sp list index → its parsed shape (if any)
  const spList = asArray(spTree['p:sp']);
  let nonPhPtr = 0;
  let phBgPtr  = 0;
  const spToItem = new Map();

  for (let i = 0; i < spList.length; i++) {
    const sp = spList[i];
    const ph = sp['p:nvSpPr']
      && sp['p:nvSpPr']['p:nvPr']
      && sp['p:nvSpPr']['p:nvPr']['p:ph'];

    if (!ph) {
      if (nonPhPtr < nonPhShapes.length) {
        spToItem.set(i, { _isMedia: false, item: nonPhShapes[nonPhPtr++] });
      }
    } else if (phHasVisibleFill(sp) && phBgPtr < phBgShapes.length) {
      spToItem.set(i, { _isMedia: false, item: phBgShapes[phBgPtr++] });
    }
  }

  // Collect pics in order (top-level only; group pics appended at end below)
  const topLevelPics = [];
  for (const pic of asArray(spTree['p:pic'])) {
    const m = picToLayoutMedia(pic, rels, dir);
    if (m) topLevelPics.push(m);
  }

  // Interleave using document order from the raw XML
  const docOrder = getSpTreeOrder(rawXml, rootTag);
  const result = [];
  let picPtr = 0;

  for (const { tag, idx } of docOrder) {
    if (tag === 'p:sp') {
      const entry = spToItem.get(idx);
      if (entry) result.push(entry);
    } else if (tag === 'p:pic') {
      if (picPtr < topLevelPics.length) {
        result.push({ _isMedia: true, item: topLevelPics[picPtr++] });
      }
    }
    // p:grpSp: no direct shape; group pics collected below
  }

  // Append anything not reached via docOrder (items inside groups, or fallback)
  while (nonPhPtr < nonPhShapes.length) result.push({ _isMedia: false, item: nonPhShapes[nonPhPtr++] });
  while (phBgPtr  < phBgShapes.length)  result.push({ _isMedia: false, item: phBgShapes[phBgPtr++] });

  // Group-embedded pics
  const groupPics = [];
  for (const grp of asArray(spTree['p:grpSp'])) {
    collectPicsFromTree(grp, rels, dir, groupPics);
  }
  for (const m of groupPics) result.push({ _isMedia: true, item: m });

  return result;
}

/**
 * Collect layout and master content in spTree document order so that shapes and
 * media items are correctly layered relative to each other.
 *
 * @param {JSZip}        zip         open PPTX archive
 * @param {string|null}  layoutPath  ZIP path to the slide layout
 * @returns {Promise<{ layoutContent: object[], masterContent: object[] }>}
 *   Each array contains { _isMedia:boolean, _source:'layout'|'master', item } in document order.
 */
async function collectLayoutContent(zip, layoutPath) {
  const result = { layoutContent: [], masterContent: [] };
  if (!layoutPath) return result;

  const layoutXml = await readText(zip, layoutPath);
  if (!layoutXml) return result;

  const layoutParsed = parseXml(layoutXml);
  const layoutSpTree = layoutParsed
    && layoutParsed['p:sldLayout']
    && layoutParsed['p:sldLayout']['p:cSld']
    && layoutParsed['p:sldLayout']['p:cSld']['p:spTree'];

  const layoutDir      = layoutPath.substring(0, layoutPath.lastIndexOf('/'));
  const layoutFile     = layoutPath.substring(layoutPath.lastIndexOf('/') + 1);
  const layoutRelsPath = `${layoutDir}/_rels/${layoutFile}.rels`;
  const layoutRelsXml  = await readText(zip, layoutRelsPath);
  const layoutRels     = layoutRelsXml ? parseRelationships(layoutRelsXml) : {};

  const warnings = [];
  if (layoutSpTree) {
    result.layoutContent = collectSpTreeOrdered(
      layoutXml, 'p:sldLayout', layoutSpTree, layoutRels, layoutDir, warnings,
    ).map((c) => ({ ...c, _source: 'layout' }));
  }

  // ---- Master ----------------------------------------------------------------

  if (!layoutRelsXml) return result;
  const masterRel = Object.values(parseRelationships(layoutRelsXml)).find((r) =>
    r.type.toLowerCase().endsWith(TYPE_SLIDE_MASTER)
  );
  if (!masterRel) return result;

  const masterPath    = resolveTarget(layoutDir, masterRel.target);
  const masterXml     = await readText(zip, masterPath);
  if (!masterXml) return result;

  const masterParsed = parseXml(masterXml);
  const masterSpTree = masterParsed
    && masterParsed['p:sldMaster']
    && masterParsed['p:sldMaster']['p:cSld']
    && masterParsed['p:sldMaster']['p:cSld']['p:spTree'];

  const masterDir      = masterPath.substring(0, masterPath.lastIndexOf('/'));
  const masterFile     = masterPath.substring(masterPath.lastIndexOf('/') + 1);
  const masterRelsPath = `${masterDir}/_rels/${masterFile}.rels`;
  const masterRelsXml  = await readText(zip, masterRelsPath);
  const masterRels     = masterRelsXml ? parseRelationships(masterRelsXml) : {};

  if (masterSpTree) {
    result.masterContent = collectSpTreeOrdered(
      masterXml, 'p:sldMaster', masterSpTree, masterRels, masterDir, warnings,
    ).map((c) => ({ ...c, _source: 'master' }));
  }

  // ---- Deduplicate master media covered by layout media ----------------------
  // Same logic as the old collectLayoutMedia dedup passes:
  // Pass 1: same file → drop master copy
  // Pass 2: same position (≥50% overlap) → drop master copy
  const layoutMediaItems = result.layoutContent
    .filter((c) => c._isMedia)
    .map((c) => c.item);
  const layoutFiles = new Set(layoutMediaItems.map((m) => m['file-link']));

  result.masterContent = result.masterContent.filter((c) => {
    if (!c._isMedia) return true;
    const mp = c.item;
    if (layoutFiles.has(mp['file-link'])) return false; // pass 1

    // pass 2 — bounding box overlap ≥ 50%
    const mx1 = mp.position.x, my1 = mp.position.y;
    const mx2 = mx1 + (mp.width || 0), my2 = my1 + (mp.height || 0);
    const mArea = (mx2 - mx1) * (my2 - my1);
    if (mArea <= 0) return true;

    return !layoutMediaItems.some((lp) => {
      const lx1 = lp.position.x, ly1 = lp.position.y;
      const lx2 = lx1 + (lp.width || 0), ly2 = ly1 + (lp.height || 0);
      const lArea = (lx2 - lx1) * (ly2 - ly1);
      if (lArea <= 0) return false;
      const ow = Math.max(0, Math.min(mx2, lx2) - Math.max(mx1, lx1));
      const oh = Math.max(0, Math.min(my2, ly2) - Math.max(my1, ly1));
      return (ow * oh) / Math.min(mArea, lArea) >= 0.5;
    });
  });

  return result;
}

module.exports = { loadLayoutGeometry, lookupGeo, collectPlaceholders, extractXfrm, collectLayoutContent };
