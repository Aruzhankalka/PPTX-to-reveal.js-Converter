const { readText } = require('./zip');
const { parseXml, asArray, getSpTreeChildOrder } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { shapeToTextBlock } = require('./text');
const { pictureToMedia, findAllPictures } = require('./media');
const { loadLayoutGeometry, lookupGeo, collectLayoutMedia, collectLayoutShapes } = require('./layouts');
const { parseShapes } = require('./shapes');
const { parseAnimations } = require('./anim');

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
async function parseSlide(zip, slidePath, txStyles) {
  const slideXml = await readText(zip, slidePath);
  if (!slideXml) {
    return { ir: { contents: { text: [], media: [], shapes: [], animations: [] } }, mediaRefs: [], warnings: [] };
  }

  // Slide directory is everything before the filename, e.g. 'ppt/slides'
  const slideDir = slidePath.substring(0, slidePath.lastIndexOf('/'));
  const filename = slidePath.substring(slidePath.lastIndexOf('/') + 1);
  const relsPath = `${slideDir}/_rels/${filename}.rels`;

  const relsXml = await readText(zip, relsPath);
  const slideRels = parseRelationships(relsXml);

  // FR-11: resolve which layout this slide uses
  const layoutRel = Object.values(slideRels).find((r) =>
    r.type.toLowerCase().endsWith('/slidelayout')
  );
  const layoutId = layoutRel ? resolveTarget(slideDir, layoutRel.target) : null;

  // Load placeholder geometry from the layout and its master so we can fill in
  // position/size for slide shapes that have no explicit <a:xfrm> of their own.
  const layoutGeometry = await loadLayoutGeometry(zip, layoutId);

  // Collect layout/master <p:pic> images (logos, decorative elements) so they
  // appear on the slide even though they live outside the slide's own spTree.
  const { layoutMedia, masterMedia } = await collectLayoutMedia(zip, layoutId);

  // Collect non-placeholder shapes from the layout/master (colored blocks,
  // lines, branding elements).  These are injected below all slide shapes.
  const { layoutShapes, masterShapes } = await collectLayoutShapes(zip, layoutId);

  const parsed = parseXml(slideXml);
  // Slide root is <p:sld><p:cSld><p:spTree>...</p:spTree></p:cSld></p:sld>
  const spTree = parsed
    && parsed['p:sld']
    && parsed['p:sld']['p:cSld']
    && parsed['p:sld']['p:cSld']['p:spTree'];

  if (!spTree) {
    return { ir: { contents: { text: [], media: [], shapes: [], animations: [] } }, mediaRefs: [], warnings: [] };
  }

  // -- Extract text blocks from <p:sp> shapes --
  const textBlocks = [];
  let textIdx = 0;
  for (const sp of asArray(spTree['p:sp'])) {
    // Resolve placeholder metadata BEFORE shapeToTextBlock so that the layout/
    // master placeholder's <a:lstStyle> can be passed into the BIU inheritance
    // cascade (OOXML level 4: slide lstStyle → layout lstStyle → txStyles).
    const ph = sp['p:nvSpPr']
      && sp['p:nvSpPr']['p:nvPr']
      && sp['p:nvSpPr']['p:nvPr']['p:ph'];
    const phIdx  = ph ? (ph['@_idx'] !== undefined ? Number(ph['@_idx']) : 0) : null;
    const phType = ph ? (ph['@_type'] || null) : null;
    const geo    = ph ? lookupGeo(layoutGeometry, phIdx, phType) : null;

    const block = shapeToTextBlock(sp, textIdx++, txStyles, geo ? geo.lstStyle : null);
    if (!block) continue;

    // FR-11: if the slide's own <p:spPr><a:xfrm> was absent, the position was
    // not set by shapeToTextBlock.  Resolve it through the inheritance chain:
    // layout placeholder first, then master placeholder.
    if (!block.position && geo) {
      block.position = geo.position;
      if (geo.width  != null) block.width  = geo.width;
      if (geo.height != null) block.height = geo.height;
      if (geo.rotation)       block.rotation = geo.rotation;
      // Resolve text anchor: layout/master value is the fallback when the
      // slide's own <a:bodyPr> carried no anchor attribute.
      if (!block['text-anchor'] && geo.textAnchor) {
        block['text-anchor'] = geo.textAnchor;
      }
      // Footer: always top-anchor regardless of the template's anchor
      // setting.  The footer box is intentionally small (≈21 px in typical
      // templates); anchor="ctr" would center-clip the text so only letter
      // middles show.  Top-anchoring displays caps/ascenders, which is
      // readable even at the tight height.
      if (phType === 'ftr') {
        block['text-anchor'] = 't';
      }
      // The footer-placement CSS fallback is superseded by the real coords.
      delete block['footer-placement'];

      // For special placeholder types (ftr, sldNum, dt) the master txStyles
      // body section does NOT define their font size.  The layout/master
      // placeholder's own lstStyle does.  Apply that size to runs that have
      // no explicit size set from the slide XML.
      if (geo.defaultFontSize) {
        for (const para of block.paragraphs) {
          for (const run of para.runs) {
            if (!run.formatting || !run.formatting.size) {
              if (!run.formatting) run.formatting = {};
              run.formatting.size = geo.defaultFontSize;
            }
          }
        }
      }

      // Apply the placeholder's default color to runs that carry no explicit
      // color.  This resolves the tx1 → dk1 alias path before we reach the
      // hard fallback below.
      if (geo.defaultColor) {
        for (const para of block.paragraphs) {
          for (const run of para.runs) {
            if (!run.formatting || !run.formatting.color) {
              if (!run.formatting) run.formatting = {};
              run.formatting.color = geo.defaultColor;
            }
          }
        }
      }

      // normAutofit from the layout/master: apply only when the slide's own
      // shape did not already carry <a:normAutofit> (text.js sets
      // _normAutofitApplied when it applies the slide-level values).
      if (geo.normAutofit && !block._normAutofitApplied) {
        const { fontScale, lnSpcRed } = geo.normAutofit;
        for (const para of block.paragraphs) {
          if (!Number.isNaN(fontScale) && fontScale !== 1 && fontScale > 0) {
            for (const run of para.runs) {
              const sz = run.formatting && run.formatting.size;
              if (sz && typeof sz === 'string' && sz.endsWith('pt')) {
                const scaled = parseFloat(sz) * fontScale;
                run.formatting.size = `${Math.round(scaled * 100) / 100}pt`;
              }
            }
          }
          if (!Number.isNaN(lnSpcRed) && lnSpcRed > 0) {
            const f = para.formatting;
            const ls = f && f['line-spacing'];
            if (ls && typeof ls === 'string' && !ls.endsWith('pt')) {
              const unitless = parseFloat(ls);
              if (!Number.isNaN(unitless)) {
                const reduced = parseFloat(Math.max(0.5, unitless * (1 - lnSpcRed)).toFixed(4));
                para.formatting['line-spacing'] = String(reduced);
              }
            } else {
              if (!para.formatting) para.formatting = {};
              para.formatting['line-spacing'] = String(Math.max(0.5, 1.0 - lnSpcRed));
            }
          }
        }
      }
    }

    // Post-process special footer/number/date placeholders regardless of whether
    // they had their own xfrm or inherited it.
    if (phType === 'ftr' || phType === 'sldNum' || phType === 'dt') {
      // These boxes are small and must not clip their single-line content.
      block.overflow = 'overflow-visible';

      // Guarantee a visible text color: if no color was set by the run itself
      // or inherited from the placeholder lstStyle, emit the theme dark color
      // so text is never invisible (white-on-white / unresolved CSS variable).
      for (const para of block.paragraphs) {
        for (const run of para.runs) {
          if (!run.formatting || !run.formatting.color) {
            if (!run.formatting) run.formatting = {};
            run.formatting.color = 'var(--theme-dk1)';
          }
        }
      }
    }

    // Always clean up the internal flag regardless of which branch ran.
    delete block._normAutofitApplied;

    textBlocks.push(block);
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

  // -- Inject layout/master media (logos, decorative images) --
  // Master content sits below layout which sits below slide in z-order, so
  // give these items IDs that sort after slide-specific media. Their z-indices
  // are assigned via the fallbackZ mechanism below (after all slide content).
  let inheritedPicIdx = 0;
  for (const m of [...masterMedia, ...layoutMedia]) {
    const zipPath    = m['file-link'];
    const bundlePath = 'media/' + zipPath.split('/').pop();
    m.id = 'inherited-img-' + inheritedPicIdx++;
    m['file-link'] = bundlePath;
    mediaItems.push(m);
    mediaRefs.push({ zipPath, bundlePath });
  }

  // -- Extract shapes (non-placeholder p:sp + p:cxnSp) --
  const shapeWarnings = [];
  const shapeItems = parseShapes(spTree, txStyles, shapeWarnings);

  // Build a map from raw p:sp ordinal → shape so the z-index walk below can
  // assign correct values.  text.js produces a textBlock for placeholder p:sp
  // nodes; parseShapes produces a shape for non-placeholder ones — the two
  // parsers partition the p:sp list cleanly.
  const spRawToShape = new Map();
  {
    const spList = asArray(spTree['p:sp']);
    let nonPhCount = 0;
    for (let i = 0; i < spList.length; i++) {
      const sp = spList[i];
      const ph = sp['p:nvSpPr']
        && sp['p:nvSpPr']['p:nvPr']
        && sp['p:nvSpPr']['p:nvPr']['p:ph'];
      if (!ph && nonPhCount < shapeItems.length) {
        spRawToShape.set(i, shapeItems[nonPhCount++]);
      }
    }
  }
  // cxnSp shapes occupy shapeItems indices starting after the p:sp shapes.
  const cxnShapeOffset = spRawToShape.size;

  // -- Extract animations --
  const animWarnings = [];
  const sldRoot = parsed['p:sld'];
  const { animations: animItems } = parseAnimations(sldRoot, animWarnings);

  // -- Assign z-index from spTree document order (FR-13) --
  // Use a preserveOrder parse to recover the interleaved sequence of p:sp,
  // p:pic, p:cxnSp, and p:grpSp children, then map each back to its parsed
  // object by same-type index.
  const spTreeOrder = getSpTreeChildOrder(slideXml);
  const assignedShapes = new Set();

  for (let z = 0; z < spTreeOrder.length; z++) {
    const { tag, idx } = spTreeOrder[z];
    if (tag === 'p:sp') {
      if (textBlocks[idx]) textBlocks[idx]['z-index'] = z;
      const shape = spRawToShape.get(idx);
      if (shape) { shape.z = z; assignedShapes.add(shape); }
    } else if (tag === 'p:pic' && mediaItems[idx]) {
      mediaItems[idx]['z-index'] = z;
    } else if (tag === 'p:cxnSp') {
      const cxnShape = shapeItems[cxnShapeOffset + idx];
      if (cxnShape) { cxnShape.z = z; assignedShapes.add(cxnShape); }
    }
    // p:grpSp: no direct element to assign; its contained pics use fallbackZ below.
  }

  // Pics extracted from inside groups and any shapes not reached via spTreeOrder
  // get z-indices above all direct spTree children so they stack above everything.
  let fallbackZ = spTreeOrder.length;
  for (const item of mediaItems) {
    if (item['z-index'] === undefined) item['z-index'] = fallbackZ++;
  }
  for (const block of textBlocks) {
    if (block['z-index'] === undefined) block['z-index'] = fallbackZ++;
  }
  for (const shape of shapeItems) {
    if (!assignedShapes.has(shape)) shape.z = fallbackZ++;
  }

  // -- Inject layout/master non-placeholder shapes as background layer --
  // Master shapes are lowest (most negative z), layout shapes sit above them,
  // both below all slide-level content (z >= 0).
  const M = masterShapes.length;
  const L = layoutShapes.length;
  for (let i = 0; i < M; i++) {
    masterShapes[i].id = `master-${masterShapes[i].id}`;
    masterShapes[i].z  = -(M + L) + i;
    shapeItems.push(masterShapes[i]);
  }
  for (let i = 0; i < L; i++) {
    layoutShapes[i].id = `layout-${layoutShapes[i].id}`;
    layoutShapes[i].z  = -L + i;
    shapeItems.push(layoutShapes[i]);
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
      shapes: shapeItems,
      animations: animItems,
    },
  };
  if (title) ir.title = title;

  const warnings = [...shapeWarnings, ...animWarnings];
  return { ir, mediaRefs, layoutId, warnings };
}

module.exports = { parseSlide };