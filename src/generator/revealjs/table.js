const { escapeCss } = require('./escape');
const { renderParagraphList, positioningToCss } = require('./text');

function verticalAlign(anchor) {
  if (anchor === 'ctr') return 'middle';
  if (anchor === 'b') return 'bottom';
  return 'top';
}

function renderCell(cell, isHeader, rowIndex, table) {
    if (cell.hMerge === true || cell.vMerge === true) return '';
  
    const tag = isHeader ? 'th' : 'td';
    const attrs = [];
  
    if (cell.colspan && cell.colspan > 1) attrs.push(`colspan="${cell.colspan}"`);
    if (cell.rowspan && cell.rowspan > 1) attrs.push(`rowspan="${cell.rowspan}"`);
  
    const isBandRow = table.bandRow === true && rowIndex > 0 && rowIndex % 2 === 1;
    const isTotalRow = rowIndex === (table.rows || []).length - 1;
  
    const background =
      cell.fill ||
      (isHeader ? '#4472C4' : isBandRow ? '#D9E1F2' : '#E9EDF7');
  
    const style = [
      `vertical-align:${verticalAlign(cell.anchor)}`,
      `padding:${cell.marT ?? 2}px ${cell.marR ?? 6}px ${cell.marB ?? 2}px ${cell.marL ?? 6}px`,
      'overflow:hidden',
      'box-sizing:border-box',
      `background:${escapeCss(background)}`,
      'border:1px solid #FFFFFF',
      isHeader ? 'color:#FFFFFF' : '',
      isHeader || isTotalRow ? 'font-weight:bold' : '',
      'text-align:center',
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
    const cells = (row.cells || []).map(cell => renderCell(cell, isHeader, rowIndex, table)).join('');
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