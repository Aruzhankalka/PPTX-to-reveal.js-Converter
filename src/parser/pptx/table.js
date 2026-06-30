'use strict';

const { asArray } = require('./xml');
const { emuToPx } = require('./units');
const { paragraphToIr } = require('./text');
const { resolveSolidFillCss } = require('./color');

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/**
 * Resolve a single <a:lnX> border node to a CSS border shorthand string,
 * e.g. "1px solid #000000", or 'none' when the node carries <a:noFill/>.
 * Returns null when the node is absent (inherit from table style).
 */
function parseBorderLine(lnNode) {
  if (!lnNode) return null;
  if (lnNode['a:noFill']) return 'none';

  const solidFill = lnNode['a:solidFill'];
  const color = solidFill ? resolveSolidFillCss(solidFill) : null;
  if (!color) return null;

  const widthEmu = lnNode['@_w'] != null ? (Number(lnNode['@_w']) || 12700) : 12700;
  const widthPx  = parseFloat((widthEmu / 9525).toFixed(2));

  // Dash style (solid is default)
  const prstDash = lnNode['a:prstDash'];
  const dashVal  = prstDash && prstDash['@_val'];
  const style    = dashVal === 'dot' || dashVal === 'sysDot' ? 'dotted'
                 : (dashVal && dashVal !== 'solid' && dashVal !== 'solidEdit') ? 'dashed'
                 : 'solid';

  return `${widthPx}px ${style} ${color}`;
}

/**
 * Parse all four border sides from <a:tcPr>.
 * Returns a border descriptor or null when no explicit borders are set.
 *
 * Output shape:
 *   - Single string when all four sides are identical: "1px solid #000000"
 *   - Object { top, right, bottom, left } when sides differ; null for absent sides.
 */
function parseCellBorder(tcPr) {
  if (!tcPr) return null;

  const top    = parseBorderLine(tcPr['a:lnT']);
  const right  = parseBorderLine(tcPr['a:lnR']);
  const bottom = parseBorderLine(tcPr['a:lnB']);
  const left   = parseBorderLine(tcPr['a:lnL']);

  if (top == null && right == null && bottom == null && left == null) return null;

  // All four sides present and identical → CSS shorthand string
  if (top != null && top === right && right === bottom && bottom === left) return top;

  // Mixed sides → object; include only the non-null entries
  const border = {};
  if (top    != null) border.top    = top;
  if (right  != null) border.right  = right;
  if (bottom != null) border.bottom = bottom;
  if (left   != null) border.left   = left;
  return Object.keys(border).length > 0 ? border : null;
}

/**
 * Parse a single <a:tc> cell into an IR cell object.
 *
 * is-header is set to false by default; parseTable upgrades to true for
 * cells in the header row / header column after the full table is built.
 */
function parseCellToIr(aTc, txStyles, slideRels) {
  const txBody = aTc['a:txBody'];
  const tcPr   = aTc['a:tcPr'] || {};

  let paragraphs = [];
  if (txBody) {
    const lstStyle = txBody['a:lstStyle'];
    paragraphs = asArray(txBody['a:p']).map((aP, pIdx) =>
      paragraphToIr(aP, pIdx, lstStyle, null, txStyles, null, slideRels, true)
    );
  }

  const cell = {
    paragraphs,
    'is-header': false, // upgraded by parseTable for header-row / header-column cells
  };

  // background (spec field name; was 'fill' in pre-spec IR)
  const solidFill = tcPr['a:solidFill'];
  if (solidFill) {
    const bg = resolveSolidFillCss(solidFill);
    if (bg) cell.background = bg;
  }

  // border — per-side CSS strings, only present when explicitly set in XML
  const border = parseCellBorder(tcPr);
  if (border) cell.border = border;

  // Vertical anchor (generator internal — not a spec field but kept for rendering)
  const anchor = tcPr['@_anchor'] || null;
  if (anchor) cell.anchor = anchor;

  // Cell padding in px (OOXML margin attrs); kept for generator use
  if (tcPr['@_marL'] != null) cell.marL = emuToPx(tcPr['@_marL']) ?? 0;
  if (tcPr['@_marR'] != null) cell.marR = emuToPx(tcPr['@_marR']) ?? 0;
  if (tcPr['@_marT'] != null) cell.marT = emuToPx(tcPr['@_marT']) ?? 0;
  if (tcPr['@_marB'] != null) cell.marB = emuToPx(tcPr['@_marB']) ?? 0;

  // Spanning
  const gridSpan = tcPr['@_gridSpan'];
  const rowSpan  = tcPr['@_rowSpan'];
  if (gridSpan && Number(gridSpan) > 1) cell.colspan = Number(gridSpan);
  if (rowSpan  && Number(rowSpan)  > 1) cell.rowspan = Number(rowSpan);

  // Merge-continuation markers (generator skips these cells)
  if (aTc['@_hMerge'] === '1' || aTc['@_hMerge'] === true) cell.hMerge = true;
  if (aTc['@_vMerge'] === '1' || aTc['@_vMerge'] === true) cell.vMerge = true;

  return cell;
}

// ---------------------------------------------------------------------------
// Main parser
// ---------------------------------------------------------------------------

/**
 * Parse a <p:graphicFrame> containing an <a:tbl> into a spec-compliant table IR.
 *
 * Spec fields implemented:
 *   id, placeholder-id, position {x,y}, pos-type, width, height, z-index,
 *   table-style {style-id, header-row, header-column, banded-rows,
 *                banded-columns, result-row, result-column},
 *   col-def [{width}],
 *   rows[].height, rows[].cells[].{colspan, rowspan, is-header, background, border, paragraphs}
 *
 * @param {object} pGraphicFrame - parsed p:graphicFrame node
 * @param {number} idx           - table index on this slide (for stable id)
 * @param {object} txStyles      - master txStyles (passed to paragraphToIr)
 * @param {object} slideRels     - slide relationship map (for hyperlink resolution)
 * @returns {object|null}
 */
function parseTable(pGraphicFrame, idx, txStyles, slideRels) {
  const graphic     = pGraphicFrame['a:graphic'];
  const graphicData = graphic && graphic['a:graphicData'];
  const tbl         = graphicData && graphicData['a:tbl'];
  if (!tbl) return null;

  // ── Placeholder metadata ────────────────────────────────────────────────────
  const nvGrFrPr = pGraphicFrame['p:nvGraphicFramePr'];
  const nvPr     = nvGrFrPr && nvGrFrPr['p:nvPr'];
  const ph        = nvPr && nvPr['p:ph'];
  const tblPptxId = nvGrFrPr && nvGrFrPr['p:cNvPr'] && Number(nvGrFrPr['p:cNvPr']['@_id']);
  const placeholderId = ph
    ? (ph['@_type'] ? `${ph['@_type']}-${ph['@_idx'] ?? 0}` : `idx-${ph['@_idx'] ?? 0}`)
    : null;
  const posType = ph ? 'relative to placeholder' : 'relative to slide';

  // ── Position and size ───────────────────────────────────────────────────────
  const xfrm   = pGraphicFrame['p:xfrm'];
  const off    = xfrm && xfrm['a:off'];
  const ext    = xfrm && xfrm['a:ext'];
  const x      = off ? (emuToPx(off['@_x']) ?? 0) : 0;
  const y      = off ? (emuToPx(off['@_y']) ?? 0) : 0;
  const width  = ext ? (emuToPx(ext['@_cx']) ?? null) : null;
  const height = ext ? (emuToPx(ext['@_cy']) ?? null) : null;

  // ── Column definitions ──────────────────────────────────────────────────────
  const tblGrid = tbl['a:tblGrid'];
  const colDef  = asArray(tblGrid && tblGrid['a:gridCol']).map(col => ({
    width: emuToPx(col['@_w']) ?? 0,
  }));

  // ── Table-level style flags from <a:tblPr> ──────────────────────────────────
  const tblPr    = tbl['a:tblPr'] || {};
  const bool     = (attr) => tblPr[attr] === '1' || tblPr[attr] === true;

  // Style-id from <a:tableStyleId> child element
  const styleIdEl = tblPr['a:tableStyleId'];
  const styleId   = typeof styleIdEl === 'string' ? styleIdEl.trim()
                  : (styleIdEl && styleIdEl['#text']) || null;

  const tableStyle = {
    'header-row':     bool('@_firstRow'),
    'header-column':  bool('@_firstCol'),
    'banded-rows':    bool('@_bandRow'),
    'banded-columns': bool('@_bandCol'),
    'result-row':     bool('@_lastRow'),
    'result-column':  bool('@_lastCol'),
  };
  if (styleId) tableStyle['style-id'] = styleId;

  // ── Rows and cells ──────────────────────────────────────────────────────────
  const rows = asArray(tbl['a:tr']).map((aRow, rowIdx) => {
    const rowHeight = emuToPx(aRow['@_h']) ?? null;
    const cells = asArray(aRow['a:tc']).map((aTc, colIdx) => {
      const cell = parseCellToIr(aTc, txStyles, slideRels);
      // Mark is-header based on table-style flags
      if (tableStyle['header-row'] && rowIdx === 0) cell['is-header'] = true;
      if (tableStyle['header-column'] && colIdx === 0) cell['is-header'] = true;
      return cell;
    });
    const row = { cells };
    if (rowHeight != null) row.height = rowHeight;
    return row;
  });

  // ── Assemble ────────────────────────────────────────────────────────────────
  const table = {
    id:           'tbl-' + idx,
    position:     { x, y },
    'pos-type':   posType,
    'table-style': tableStyle,
    rows,
    'z-index':    0, // overridden by slide.js via getSpTreeChildOrder
  };

  if (placeholderId) table['placeholder-id'] = placeholderId;
  if (tblPptxId) table._pptxId = tblPptxId;
  if (colDef.length > 0) table['col-def'] = colDef;
  if (width  != null) table.width  = width;
  if (height != null) table.height = height;

  return table;
}

module.exports = { parseTable };
