const { escapeCss } = require('./escape');
const { renderParagraphList, positioningToCss } = require('./text');

function verticalAlign(anchor) {
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b') return 'bottom';
  return 'top';
}

function renderCell(cell, isHeader) {
  if (cell.hMerge === true || cell.vMerge === true) return '';

  const tag = isHeader ? 'th' : 'td';
  const attrs = [];

  if (cell.colspan && cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
  if (cell.rowspan && cell.rowspan > 1) attrs.push(`rowspan="${cell.rowspan}"`);

  const style = [
    `vertical-align:${verticalAlign(cell.anchor)}`,
    `padding:${cell.marT || 0}px ${cell.marR || 0}px ${cell.marB || 0}px ${cell.marL || 0}px`,
    'overflow:hidden',
    'box-sizing:border-box',
    cell.fill ? `background:${escapeCss(cell.fill)}` : '',
  ].filter(Boolean).join(';');

  const content = renderParagraphList(cell.paragraphs || []);

  return `<${tag}${attrs.length ? ' ' + attrs.join(' ') : ''} style="${style}">${content}</${tag}>`;
}

function renderTable(table) {
  const css = positioningToCss(table);
  const styleAttr = css ? ` style="${css}"` : '';

  const colgroup = (table.colWidths || [])
    .map(w => `<col style="width:${w}px">`)
    .join('');

  const rows = (table.rows || []).map((row, rowIndex) => {
    const isHeader = table.firstRow === true && rowIndex === 0;
    const cells = (row.cells || []).map(cell => renderCell(cell, isHeader)).join('');
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