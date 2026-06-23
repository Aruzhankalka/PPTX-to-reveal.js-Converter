const { asArray } = require('./xml');
const { emuToPx } = require('./units');
const { paragraphToIr } = require('./text');

/**
 * Parse a cell fill from an <a:tcPr> node.
 * Returns a CSS color string or null.
 */
function parseCellFill(tcPr) {
  if (!tcPr) return null;
  const solidFill = tcPr['a:solidFill'];
  if (!solidFill) return null;
  if (solidFill['a:srgbClr'] && solidFill['a:srgbClr']['@_val']) {
    return '#' + solidFill['a:srgbClr']['@_val'];
  }
  if (solidFill['a:schemeClr'] && solidFill['a:schemeClr']['@_val']) {
    return `var(--theme-${solidFill['a:schemeClr']['@_val']})`;
  }
  return null;
}

/**
 * Parse a single <a:tc> cell into an IR cell object.
 * paragraphToIr is called with skipMasterSizeFallbacks=true because table cells
 * are not placeholder shapes and must not inherit the master body font size.
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

  const cell = { paragraphs };

  const fill = parseCellFill(tcPr);
  if (fill) cell.fill = fill;

  const anchor = tcPr['@_anchor'] || null;
  if (anchor) cell.anchor = anchor;

  // Cell padding in px (OOXML calls them margins)
  if (tcPr['@_marL'] != null) cell.marL = emuToPx(tcPr['@_marL']) ?? 0;
  if (tcPr['@_marR'] != null) cell.marR = emuToPx(tcPr['@_marR']) ?? 0;
  if (tcPr['@_marT'] != null) cell.marT = emuToPx(tcPr['@_marT']) ?? 0;
  if (tcPr['@_marB'] != null) cell.marB = emuToPx(tcPr['@_marB']) ?? 0;

  // Spanning
  const gridSpan = tcPr['@_gridSpan'];
  const rowSpan  = tcPr['@_rowSpan'];
  if (gridSpan && Number(gridSpan) > 1) cell.colspan = Number(gridSpan);
  if (rowSpan  && Number(rowSpan)  > 1) cell.rowspan = Number(rowSpan);

  // Continuation cells (hMerge/vMerge) — these are rendered by the spanning cell
  if (aTc['@_hMerge'] === '1' || aTc['@_hMerge'] === true) cell.hMerge = true;
  if (aTc['@_vMerge'] === '1' || aTc['@_vMerge'] === true) cell.vMerge = true;

  return cell;
}

/**
 * Parse a <p:graphicFrame> containing an <a:tbl> into a table IR object.
 *
 * @param {object} pGraphicFrame  - parsed p:graphicFrame node
 * @param {number} idx            - table index on this slide (for ID generation)
 * @param {object} txStyles       - master txStyles (passed to paragraphToIr)
 * @param {object} slideRels      - slide relationship map (for hyperlink resolution)
 * @returns {object|null}         - table IR or null if the frame has no table
 */
function parseTable(pGraphicFrame, idx, txStyles, slideRels) {
  const graphic     = pGraphicFrame['a:graphic'];
  const graphicData = graphic && graphic['a:graphicData'];
  const tbl         = graphicData && graphicData['a:tbl'];
  if (!tbl) return null;

  // Position and size from <p:xfrm>
  const xfrm = pGraphicFrame['p:xfrm'];
  const off   = xfrm && xfrm['a:off'];
  const ext   = xfrm && xfrm['a:ext'];
  const x      = off ? (emuToPx(off['@_x']) ?? 0) : 0;
  const y      = off ? (emuToPx(off['@_y']) ?? 0) : 0;
  const width  = ext ? emuToPx(ext['@_cx']) : null;
  const height = ext ? emuToPx(ext['@_cy']) : null;

  // Column widths from <a:tblGrid>
  const tblGrid   = tbl['a:tblGrid'];
  const colWidths = asArray(tblGrid && tblGrid['a:gridCol']).map(col =>
    emuToPx(col['@_w']) ?? 0
  );

  // Table-level flags from <a:tblPr>
  const tblPr   = tbl['a:tblPr'] || {};
  const firstRow = tblPr['@_firstRow'] === '1' || tblPr['@_firstRow'] === true;
  const bandRow  = tblPr['@_bandRow']  === '1' || tblPr['@_bandRow']  === true;

  // Rows
  const rows = asArray(tbl['a:tr']).map(aRow => {
    const rowHeight = emuToPx(aRow['@_h']) ?? null;
    const cells = asArray(aRow['a:tc']).map(aTc =>
      parseCellToIr(aTc, txStyles, slideRels)
    );
    const row = { cells };
    if (rowHeight != null) row.height = rowHeight;
    return row;
  });

  const table = {
    id:       'tbl-' + idx,
    position: { x, y },
    rows,
  };
  if (colWidths.length > 0) table.colWidths = colWidths;
  if (width  != null) table.width  = width;
  if (height != null) table.height = height;
  if (firstRow) table.firstRow = true;
  if (bandRow)  table.bandRow  = true;

  return table;
}

module.exports = { parseTable };
