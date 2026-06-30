const { readText } = require('./zip');
const { parseXml, asArray, getSpTreeChildOrder } = require('./xml');
const { parseRelationships, resolveTarget } = require('./relationships');
const { shapeToTextBlock } = require('./text');
const { pictureToMedia } = require('./media');
const { loadLayoutGeometry, lookupGeo, collectLayoutContent } = require('./layouts');
const { parseShapes, resolveColorNode } = require('./shapes');
const { parseAnimations } = require('./anim');
const { parseTable } = require('./table');

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
  // Also returns layoutName and showMasterSp from the <p:sldLayout> element.
  const layoutGeometry = await loadLayoutGeometry(zip, layoutId);
  const { layoutName, showMasterSp: layoutShowMasterSp } = layoutGeometry;

  // Collect layout/master shapes and images in spTree document order so that
  // shapes and media are correctly layered relative to each other.
  const { layoutContent, masterContent } = await collectLayoutContent(zip, layoutId);

  // Gate master-level content via OOXML showMasterSp attribute.
  // Layout showMasterSp="0" → this layout type hides master shapes/media.
  if (!layoutShowMasterSp) {
    masterContent.length = 0;
  }

  const parsed = parseXml(slideXml);
  // Slide root is <p:sld><p:cSld><p:spTree>...</p:spTree></p:cSld></p:sld>
  const sldRoot = parsed && parsed['p:sld'];
  // Slide showMasterSp="0" → this individual slide suppresses ALL inherited content.
  const sldShowMasterSp = sldRoot && sldRoot['@_showMasterSp'];
  if (sldShowMasterSp === '0' || sldShowMasterSp === false) {
    masterContent.length = 0;
    layoutContent.length = 0;
  }

  const spTree = sldRoot
    && sldRoot['p:cSld']
    && sldRoot['p:cSld']['p:spTree'];

  if (!spTree) {
    return { ir: { contents: { text: [], media: [], shapes: [], animations: [] } }, mediaRefs: [], warnings: [] };
  }

  // -- Extract text blocks from <p:sp> placeholder shapes --
  //
  // textBlocksBySp is a SPARSE array indexed by the position of the p:sp
  // element in the spTree (0-based).  Sparse indexing is required because the
  // z-index assignment loop (below) uses the same p:sp position index via
  // getSpTreeChildOrder, and skipping elements would misalign a compact array.
  //
  // Non-placeholder shapes (no <p:ph>) are skipped here: their text is embedded
  // inside the shape IR object by parseShapes → extractEmbeddedText and rendered
  // as part of the SVG <foreignObject>.  Including them here too would produce a
  // duplicate text block floating over the shape.
  const textBlocksBySp = []; // sparse; textBlocksBySp[spIdx] = block | undefined
  let textIdx = 0;
  const spListForText = asArray(spTree['p:sp']);
  for (let spI = 0; spI < spListForText.length; spI++) {
    const sp = spListForText[spI];
    const ph = sp['p:nvSpPr']
      && sp['p:nvSpPr']['p:nvPr']
      && sp['p:nvSpPr']['p:nvPr']['p:ph'];

    // Graphical shapes (prstGeom or custGeom, no ph) have their text embedded
    // inside the shape via parseShapes → extractEmbeddedText.  Parsing their text
    // here too produces a duplicate floating text block on top of the SVG shape.
    // Content-container shapes (no ph AND no geometry) are kept as standalone
    // text blocks because parseSp emits them as type:'unknown' (invisible SVG).
    const spPrForCheck = sp['p:spPr'] || {};
    const isGraphicalShape = !ph && (spPrForCheck['a:prstGeom'] || spPrForCheck['a:custGeom']);
    if (isGraphicalShape) continue;

    // Resolve placeholder metadata so the layout/master <a:lstStyle> can be
    // passed into the BIU inheritance cascade (OOXML level 4).
    // ph may still be undefined for content-container shapes (no ph, no prstGeom).
    const phIdx  = (ph && ph['@_idx'] !== undefined) ? Number(ph['@_idx']) : 0;
    const phType = (ph && ph['@_type']) || null;
    const geo    = ph ? lookupGeo(layoutGeometry, phIdx, phType) : null;

    const block = shapeToTextBlock(sp, textIdx++, txStyles, geo ? geo.lstStyle : null, slideRels);
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

    // Small layout-defined compact placeholders (e.g. footer-area annotation boxes
    // typed as 'body' but sized 9pt by the layout lstStyle) also need overflow
    // visible — the spacing cascade fix prevents master body spacing, but this is
    // a safety net against any residual overflow on small boxes.
    if (!block.overflow && geo && geo.defaultFontSize && block.height != null && block.height < 35) {
      block.overflow = 'overflow-visible';
    }

    // Always clean up the internal flag regardless of which branch ran.
    delete block._normAutofitApplied;

    textBlocksBySp[spI] = block; // sparse: spI = position of this p:sp in the spTree
  }

  // Compact array for IR output and fallback z-index iteration.
  const textBlocks = textBlocksBySp.filter(Boolean);

  // -- Extract shapes (non-placeholder p:sp + p:cxnSp), groups, and pictures --
  // One pass over spTree, including nested p:grpSp, so the group transform
  // computed for shapes (FR-10 groups[]) is shared by the media extraction
  // below instead of being recomputed by a second, uncoordinated walk.
  const shapeWarnings = [];
  const { shapes: shapeItems, groups: groupItems, topLevelGroupsByIdx, pictures } =
    parseShapes(spTree, txStyles, shapeWarnings);

  // -- Extract media from <p:pic> shapes (including inside groups) --
  // Each picture's transform (identity, or composed from its ancestor groups)
  // corrects its position/rotation into absolute slide EMU; its id is then
  // recorded into elementsOut — its owning group's elements[], or a
  // throwaway array for pictures that aren't inside any group.
  const mediaItems = [];
  const mediaRefs = [];
  let picIdx = 0;
  for (const { pPic, transform, elementsOut } of pictures) {
    const media = pictureToMedia(pPic, slideRels, slideDir, resolveTarget, picIdx++, transform);
    if (media) {
      mediaItems.push(media);
      elementsOut.push(media.id);
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

  // -- Prepare layout/master media (logos, decorative images) --
  // Resolve IDs and bundle paths before z-index assignment so mediaRefs is
  // populated regardless of the stacking order.
  let inheritedPicIdx = 0;
  for (const { _isMedia, item } of [...masterContent, ...layoutContent]) {
    if (!_isMedia) continue;
    const zipPath    = item['file-link'];
    const bundlePath = 'media/' + zipPath.split('/').pop();
    item.id = 'inherited-img-' + inheritedPicIdx++;
    item['file-link'] = bundlePath;
    mediaRefs.push({ zipPath, bundlePath });
  }

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

  // -- Extract tables from <p:graphicFrame> elements --
  const tableItems = [];
  let tableIdx = 0;
  for (const frame of asArray(spTree['p:graphicFrame'])) {
    const table = parseTable(frame, tableIdx, txStyles, slideRels);
    if (table) {
      tableItems.push(table);
      tableIdx++;
    }
  }

  // -- Extract animations --
  const animWarnings = [];
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
      const block = textBlocksBySp[idx]; // sparse lookup — correct for any mix of ph / non-ph
      if (block) block['z-index'] = z;
      const shape = spRawToShape.get(idx);
      if (shape) { shape['z-index'] = z; assignedShapes.add(shape); }
    } else if (tag === 'p:pic' && mediaItems[idx]) {
      mediaItems[idx]['z-index'] = z;
    } else if (tag === 'p:cxnSp') {
      const cxnShape = shapeItems[cxnShapeOffset + idx];
      if (cxnShape) { cxnShape['z-index'] = z; assignedShapes.add(cxnShape); }
    } else if (tag === 'p:graphicFrame' && tableItems[idx]) {
      tableItems[idx]['z-index'] = z;
    } else if (tag === 'p:grpSp' && idx < topLevelGroupsByIdx.length) {
      topLevelGroupsByIdx[idx]['z-index'] = z;
    }
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
    if (!assignedShapes.has(shape)) shape['z-index'] = fallbackZ++;
  }
  for (const table of tableItems) {
    if (table['z-index'] === undefined) table['z-index'] = fallbackZ++;
  }
  // Nested groups (not in spTreeOrder) get fallback z-indices above slide content.
  for (const group of groupItems) {
    if (group['z-index'] === 0 && !topLevelGroupsByIdx.includes(group)) {
      group['z-index'] = fallbackZ++;
    }
  }

  // -- Inject inherited shapes/media as background layer below all slide content --
  // Render order: master layer (in document order) → layout layer (in document order)
  // → slide content (z≥0).  Each layer preserves the spTree order so that shapes
  // and media interleave correctly (e.g. a decorative JPEG that appears before
  // accent rectangles in the layout XML will get a lower z than those rectangles).
  // Text on inherited shapes is stripped so placeholder prompt text never shows.
  const allInherited = [...masterContent, ...layoutContent];
  let iZ = -allInherited.length;

  for (const { _isMedia, _source, item } of allInherited) {
    if (_isMedia) {
      item['z-index'] = iZ++;
      mediaItems.push(item);
    } else {
      item.id = `${_source}-${item.id}`;
      delete item.text;
      item['z-index'] = iZ++;
      shapeItems.push(item);
    }
  }

  // -- Resolve animation targetIds from raw spid-N to stable IR element ids --
  // Each parser stamps _pptxId (the OOXML p:cNvPr @id integer) onto its output
  // objects. We collect these into a lookup so that animation.targetId can be
  // remapped from "spid-3" to the actual IR id (e.g. "shp-0") that validators
  // and downstream consumers can use as a cross-reference.
  if (animItems.length > 0) {
    const spidToIrId = new Map();
    for (const s of shapeItems)  { if (s._pptxId)  spidToIrId.set(s._pptxId,  s.id); }
    for (const b of textBlocks)  { if (b._pptxId)  spidToIrId.set(b._pptxId,  b.id); }
    for (const m of mediaItems)  { if (m._pptxId)  spidToIrId.set(m._pptxId,  m.id); }
    for (const t of tableItems)  { if (t._pptxId)  spidToIrId.set(t._pptxId,  t.id); }
    for (const g of groupItems)  { if (g._pptxId)  spidToIrId.set(g._pptxId,  g.id); }
    for (const anim of animItems) {
      if (anim.targetId && anim.targetId.startsWith('spid-')) {
        const spid    = Number(anim.targetId.slice(5));
        const resolved = spidToIrId.get(spid);
        if (resolved) anim.targetId = resolved;
        // Unresolved spid-N: target is on a layout/master or an unsupported type;
        // keep the raw spid so validator's spid-N exemption still applies.
      }
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

  // -- Background --
  // <p:cSld><p:bg><p:bgPr> holds the slide-level background fill.
  // We resolve solid fills only; gradient/image backgrounds are left for a future pass.
  let bgCss = null;
  const bgPr = sldRoot
    && sldRoot['p:cSld']
    && sldRoot['p:cSld']['p:bg']
    && sldRoot['p:cSld']['p:bg']['p:bgPr'];
  if (bgPr && bgPr['a:solidFill']) {
    const c = resolveColorNode(bgPr['a:solidFill']);
    if (c && c.space === 'srgb')   bgCss = `#${c.hex}`;
    else if (c && c.space === 'theme') bgCss = `var(--theme-${c.ref})`;
  }

  // -- Transition --
  // Map PPTX transition child element names to reveal.js transition names.
  const TRANSITION_MAP = {
    'p:fade':     'fade',
    'p:push':     'slide',
    'p:wipe':     'slide',
    'p:zoom':     'zoom',
    'p:wheel':    'convex',
    'p:cut':      'none',
    'p:dissolve': 'fade',
    'p:strips':   'slide',
    'p:split':    'zoom',
    'p:blinds':   'slide',
    'p:circle':   'convex',
    'p:newsflash':'zoom',
    'p:plus':     'zoom',
    'p:wedge':    'zoom',
  };
  let transitionName = null;
  const transitionEl = sldRoot && sldRoot['p:transition'];
  if (transitionEl) {
    for (const [tag, name] of Object.entries(TRANSITION_MAP)) {
      if (transitionEl[tag] !== undefined) { transitionName = name; break; }
    }
  }

  const ir = {
    contents: {
      text: textBlocks,
      media: mediaItems,
      shapes: shapeItems,
      groups: groupItems,
      tables: tableItems,
      animations: animItems,
      ...(bgCss         && { background: bgCss }),
      ...(transitionName && { transition: transitionName }),
    },
  };
  if (title) ir.title = title;
  if (layoutName) ir.layoutName = layoutName;

  // PPTX <p:sld show="0"> marks a slide as hidden (not shown during presentation).
  // Absent or show="1" means visible — only set the field when explicitly hidden.
  const show = sldRoot && sldRoot['@_show'];
  if (show === '0' || show === false || show === 0) ir.hidden = true;

  // -- Notes --
  // Notes live in a separate notesSlide XML file referenced via the slide's rels.
  // We extract plain text from the body placeholder (ph idx=1); paragraphs are
  // joined with \n so the generator can render them as <br/> in the notes pane.
  const notesRel = Object.values(slideRels).find((r) =>
    r.type.toLowerCase().endsWith('/notesslide')
  );
  if (notesRel) {
    const notesPath = resolveTarget(slideDir, notesRel.target);
    const notesXml  = await readText(zip, notesPath);
    if (notesXml) {
      const notesParsed  = parseXml(notesXml);
      const notesSpTree  = notesParsed
        && notesParsed['p:notes']
        && notesParsed['p:notes']['p:cSld']
        && notesParsed['p:notes']['p:cSld']['p:spTree'];
      if (notesSpTree) {
        for (const sp of asArray(notesSpTree['p:sp'])) {
          const ph = sp['p:nvSpPr']
            && sp['p:nvSpPr']['p:nvPr']
            && sp['p:nvSpPr']['p:nvPr']['p:ph'];
          // Skip the slide-image thumbnail placeholder (type="sp")
          if (ph && ph['@_type'] === 'sp') continue;
          const txBody = sp['p:txBody'];
          if (!txBody) continue;
          const lines = asArray(txBody['a:p']).map((para) =>
            asArray(para['a:r']).map((r) => r['a:t'] || '').join('')
          );
          const notes = lines.filter(Boolean).join('\n');
          if (notes) { ir.contents.notes = notes; break; }
        }
      }
    }
  }

  const warnings = [...shapeWarnings, ...animWarnings];
  return { ir, mediaRefs, layoutId, warnings };
}

module.exports = { parseSlide };