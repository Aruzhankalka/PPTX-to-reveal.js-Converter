'use strict';

const { escapeCss } = require('./escape');
const { renderParagraphList, positioningToCss } = require('./text');

function verticalAlign(anchor) {
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b')   return 'bottom';
  return 'top';
}

/**
 * Convert a border field to a CSS border shorthand string.
 * Accepts a plain string ("1px solid #000") or a per-side object
 * {top, right, bottom, left} from the new IR, OR undefined (fall back
 * to default).
 */
function resolveBorderCss(borderField, defaultCss) {
  if (!borderField) return defaultCss;
  if (typeof borderField === 'string') return borderField;
  // Per-side object: CSS individual properties concatenated with ';'
  const parts = [];
  if (borderField.top)    parts.push(`border-top:${borderField.top}`);
  if (borderField.right)  parts.push(`border-right:${borderField.right}`);
  if (borderField.bottom) parts.push(`border-bottom:${borderField.bottom}`);
  if (borderField.left)   parts.push(`border-left:${borderField.left}`);
  return parts.length > 0 ? parts.join(';') : defaultCss;
}

function renderCell(cell, rowIndex, colIndex, table) {
  if (cell.hMerge === true || cell.vMerge === true) return '';

  const ts  = table['table-style'] || {};
  const tag = cell['is-header'] ? 'th' : 'td';
  const attrs = [];

  if (cell.colspan && cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
  if (cell.rowspan && cell.rowspan > 1) attrs.push(`rowspan="${cell.rowspan}"`);

  const isBandedRow    = ts['banded-rows']    === true && rowIndex > 0 && rowIndex % 2 === 1;
  const isResultRow    = ts['result-row']      === true && rowIndex === (table.rows || []).length - 1;
  const isBandedCol    = ts['banded-columns']  === true && colIndex % 2 === 1;
  const isResultCol    = ts['result-column']   === true;

  // Background — use explicit per-cell value, then fall back to style-based defaults
  const background =
    cell.background ||
    (cell['is-header'] ? '#4472C4' : isBandedRow || isBandedCol ? '#D9E1F2' : '#E9EDF7');

  // Border — use explicit per-cell value (new IR field), or hardcoded default
  const borderCss = resolveBorderCss(cell.border, 'border:1px solid #FFFFFF');
  // Wrap per-side properties under a 'border-X:...' form already handled above,
  // but the default is a shorthand. CSS is concatenated into the style string.
  const borderStyle = borderCss.startsWith('border:') ? borderCss : borderCss;

  const style = [
    `vertical-align:${verticalAlign(cell.anchor)}`,
    `padding:${cell.marT ?? 2}px ${cell.marR ?? 6}px ${cell.marB ?? 2}px ${cell.marL ?? 6}px`,
    'overflow:hidden',
    'box-sizing:border-box',
    `background:${escapeCss(background)}`,
    borderStyle,
    cell['is-header'] ? 'color:#FFFFFF' : '',
    (cell['is-header'] || isResultRow || isResultCol) ? 'font-weight:bold' : '',
    'text-align:center',
  ].filter(Boolean).join(';');

  const content = renderParagraphList(cell.paragraphs || []);

  return `<${tag}${attrs.length ? ' ' + attrs.join(' ') : ''} style="${style}">${content}</${tag}>`;
}

function renderTable(table) {
  const css      = positioningToCss(table);
  const styleAttr = css ? ` style="${css}"` : '';

  // col-def is the new spec field; colWidths is the old field (backward compat)
  const colDef   = table['col-def'] || (table.colWidths || []).map(w => ({ width: w }));
  const colgroup = colDef.map(c => `<col style="width:${c.width}px">`).join('');

  const rows = (table.rows || []).map((row, rowIndex) => {
    const cells = (row.cells || []).map((cell, colIndex) =>
      renderCell(cell, rowIndex, colIndex, table)
    ).join('');
    const rowStyle = typeof row.height === 'number' ? ` style="height:${row.height}px"` : '';
    return `<tr${rowStyle}>${cells}</tr>`;
  }).join('\n');

  return `<div class="slide-table"${styleAttr}>
<table>
<colgroup>${colgroup}</colgroup>
<tbody>
${rows}
</tbody>
</table>
</div>`;
}

module.exports = { renderTable };
