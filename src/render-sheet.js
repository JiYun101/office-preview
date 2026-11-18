import { formatSpreadsheetValue } from './spreadsheet-format.js';

function node(tag, className = '', text) {
  const element = document.createElement(tag); element.className = className;
  if (text !== undefined) element.textContent = String(text);
  return element;
}
function columnName(index) {
  let result = '';
  for (let value = index + 1; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result;
  return result;
}
function indexAt(offsets, position) {
  let low = 0, high = offsets.length - 1;
  while (low < high) {
    const middle = Math.floor((low + high + 1) / 2);
    if (offsets[middle] <= position) low = middle; else high = middle - 1;
  }
  return Math.min(low, offsets.length - 2);
}

function layoutAxis(count, hidden, sizeAt) {
  const indices = [], offsets = [0], positions = new Uint32Array(count + 1);
  for (let index = 0; index < count; index++) {
    positions[index] = indices.length;
    if (hidden?.[index]) continue;
    const size = sizeAt(index);
    if (size === 0) continue;
    indices.push(index);
    offsets.push(offsets.at(-1) + size);
  }
  positions[count] = indices.length;
  return { indices, offsets, positions };
}

// Only the visible rows/columns exist as cells. The spacer retains the complete
// worksheet geometry, so scrolling never changes the file's row/column sizes.
export function renderSheet(model, viewport, content, auxiliary, options = {}) {
  const { runStyle, fontStack } = options;
  if (typeof runStyle !== 'function' || typeof fontStack !== 'function') throw new TypeError('Sheet rendering requires the shared font/style helpers.');
  const ROW_HEADER = 45, COLUMN_HEADER = 31;
  const sheets = model.sheets?.length ? model.sheets : [{ name: '数据', rows: [] }];
  const formatted = model.format === 'xls';
  const measure = document.createElement('canvas').getContext('2d');
  const layouts = new Map();
  let selected = 0, disposed = false, frame = 0, windowKey = '', selectedAddress = '';
  content.style.padding = '0';
  const spacer = node('div', 'op-sheet-spacer');
  Object.assign(spacer.style, { position: 'relative', background: '#fff' });
  content.replaceChildren(spacer); auxiliary.replaceChildren();
  let tabButtons = [];
  if (sheets.length > 1) {
    const tabs = node('div', 'op-tabs'); tabs.setAttribute('role', 'tablist'); tabs.setAttribute('aria-label', '工作表');
    tabButtons = sheets.map((sheet, index) => {
      const tab = node('button', '', sheet.name); tab.type = 'button'; tab.setAttribute('role', 'tab');
      tab.onclick = () => { selected = index; selectedAddress = ''; windowKey = ''; viewport.scrollTo(0, 0); draw(); };
      tabs.append(tab); return tab;
    });
    auxiliary.append(tabs);
  }

  function layoutFor(index) {
    if (layouts.has(index)) return layouts.get(index);
    const sheet = sheets[index];
    let rowCount = sheet.rows.length, colCount = sheet.rows.reduce((max, row) => Math.max(max, row?.length || 0), 1);
    for (const merge of sheet.merges || []) { rowCount = Math.max(rowCount, merge.endRow + 1); colCount = Math.max(colCount, merge.endCol + 1); }
    let digit = 7;
    if (formatted && sheet.defaultFont) {
      measure.font = `${sheet.defaultFont.fontSize || 14}px ${fontStack(sheet.defaultFont.fontFamily || 'Arial')}`;
      digit = Math.max(...Array.from('0123456789', character => measure.measureText(character).width));
    }
    const columnAxis = layoutAxis(colCount, sheet.hiddenColumns, col => {
      const units = sheet.columnWidthUnits?.[col] ?? sheet.defaultColumnWidthUnits;
      const size = units !== undefined && sheet.defaultFont ? units / 256 * digit + 5 : sheet.columnWidths?.[col] ?? sheet.defaultColumnWidth ?? 105;
      return Math.max(1, Math.min(2000, size));
    });
    const rowAxis = layoutAxis(rowCount, sheet.hiddenRows, row => {
      let height = sheet.rowHeights?.[row] ?? sheet.defaultRowHeight ?? (formatted ? 20 : 31);
      // CSV/XLSX values can contain deliberate line breaks. Preserve those
      // without allowing content-driven table growth to corrupt scroll offsets.
      if (!formatted && sheet.rowHeights?.[row] === undefined && sheet.defaultRowHeight === undefined) {
        for (const cell of sheet.rows[row] || []) {
          if (cell?.value !== undefined) height = Math.max(height, String(cell.value).split(/\r\n|\r|\n/).length * 19.5 + 10);
        }
      }
      return Math.max(0, height);
    });
    const result = { sheet, rowAxis, columnAxis };
    layouts.set(index, result); return result;
  }
  const styleFor = (sheet, row, col) => sheet.rows[row]?.[col]?.style || sheet.rowStyles?.[row] || sheet.columnStyles?.[col] || sheet.defaultStyle || {};
  function border(element, side, value) {
    if (value?.width > 0) element.style[`border${side[0].toUpperCase()}${side.slice(1)}`] = `${value.width}px ${value.style || 'solid'} ${value.color || '#000'}`;
  }
  function draw() {
    if (disposed) return;
    const { sheet, rowAxis, columnAxis } = layoutFor(selected);
    const { indices: rowIndices, offsets: rows, positions: rowPositions } = rowAxis;
    const { indices: colIndices, offsets: columns, positions: colPositions } = columnAxis;
    const rowCount = rowIndices.length, colCount = colIndices.length;
    tabButtons.forEach((tab, index) => tab.setAttribute('aria-selected', String(index === selected)));
    spacer.style.width = `${ROW_HEADER + columns[colCount]}px`;
    spacer.style.height = `${COLUMN_HEADER + rows[rowCount]}px`;
    const scale = parseFloat(getComputedStyle(content).zoom) || 1;
    const top = viewport.scrollTop / scale, left = viewport.scrollLeft / scale;
    const height = (viewport.clientHeight || 400) / scale, width = (viewport.clientWidth || 800) / scale;
    const startRow = rowCount ? Math.max(0, indexAt(rows, Math.max(0, top - COLUMN_HEADER)) - 12) : 0;
    const endRow = rowCount ? Math.min(rowCount, indexAt(rows, top + height) + 14) : 0;
    const startCol = Math.max(0, indexAt(columns, Math.max(0, left - ROW_HEADER)) - 3);
    const endCol = Math.min(colCount, indexAt(columns, left + width) + 5);
    const key = [selected, startRow, endRow, startCol, endCol].join(':');
    if (key === windowKey) return;
    windowKey = key;
    const table = node('table', 'op-grid op-sheet-formatted op-sheet-window'); table.setAttribute('aria-label', `${sheet.name} 工作表`);
    table.dataset.startRow = String(rowIndices[startRow] ?? 0); table.dataset.endRow = String(rowCount ? rowIndices[endRow - 1] + 1 : 0);
    table.dataset.startCol = String(colIndices[startCol] ?? 0); table.dataset.endCol = String(colCount ? colIndices[endCol - 1] + 1 : 0);
    Object.assign(table.style, { position: 'absolute', top: `${rows[startRow]}px`, left: `${columns[startCol]}px`, width: `${ROW_HEADER + columns[endCol] - columns[startCol]}px`, tableLayout: 'fixed', borderCollapse: formatted ? 'collapse' : 'separate', borderSpacing: '0' });
    const group = node('colgroup'), rowNumberCol = node('col'); rowNumberCol.style.width = `${ROW_HEADER}px`; group.append(rowNumberCol);
    for (let col = startCol; col < endCol; col++) { const element = node('col'); element.style.width = `${columns[col + 1] - columns[col]}px`; group.append(element); }
    table.append(group);
    const thead = node('thead'), head = node('tr'); head.style.height = `${COLUMN_HEADER}px`;
    const corner = node('th'); corner.style.width = `${ROW_HEADER}px`; head.append(corner);
    for (let col = startCol; col < endCol; col++) { const th = node('th', '', columnName(colIndices[col])); th.scope = 'col'; th.style.whiteSpace = 'nowrap'; th.style.overflow = 'hidden'; head.append(th); }
    thead.append(head); table.append(thead);
    const merged = new Map();
    for (const range of sheet.merges || []) {
      // Convert the physical range to visible coordinates before clipping to the window.
      const firstRow = rowPositions[range.startRow], afterRow = rowPositions[range.endRow + 1];
      const firstCol = colPositions[range.startCol], afterCol = colPositions[range.endCol + 1];
      const top = Math.max(startRow, firstRow), bottom = Math.min(endRow - 1, afterRow - 1);
      const left = Math.max(startCol, firstCol), right = Math.min(endCol - 1, afterCol - 1);
      const merge = { ...range, top, bottom, left, right, firstRow, afterRow, firstCol, afterCol };
      for (let row = top; row <= bottom; row++) for (let col = left; col <= right; col++) {
        if (merged.has(`${row}:${col}`)) throw new Error('工作表包含重叠的合并区域，无法正确预览。');
        merged.set(`${row}:${col}`, merge);
      }
    }
    const tbody = node('tbody');
    for (let row = startRow; row < endRow; row++) {
      const tr = node('tr'); tr.style.height = `${rows[row + 1] - rows[row]}px`; tr.dataset.row = String(rowIndices[row]);
      const th = node('th'), number = node('span', '', rowIndices[row] + 1); th.scope = 'row'; th.style.top = 'auto'; th.style.position = 'sticky';
      Object.assign(number.style, { position: 'absolute', inset: '0', display: 'flex', alignItems: 'center', justifyContent: 'center', lineHeight: 'normal' }); th.append(number); tr.append(th);
      for (let col = startCol; col < endCol; col++) {
        const merge = merged.get(`${row}:${col}`);
        if (merge && (row !== merge.top || col !== merge.left)) continue;
        const sourceRow = merge?.startRow ?? rowIndices[row], sourceCol = merge?.startCol ?? colIndices[col];
        const cell = sheet.rows[sourceRow]?.[sourceCol], value = cell?.value ?? '', style = cell?.style || styleFor(sheet, sourceRow, sourceCol);
        const display = formatted || (typeof value === 'number' && style.format) ? formatSpreadsheetValue(value, style.format, { date1904: sheet.date1904 }) : value;
        const td = node('td'); td.dataset.row = String(sourceRow); td.dataset.col = String(sourceCol);
        if (merge) { td.rowSpan = merge.bottom - merge.top + 1; td.colSpan = merge.right - merge.left + 1; }
        runStyle(td, style); td.style.textAlign = style.align || (typeof value === 'number' ? 'right' : 'left');
        if (style.verticalAlign) td.style.verticalAlign = style.verticalAlign;
        for (const side of ['top', 'right', 'bottom', 'left']) border(td, side, style.borders?.[side]);
        if (formatted) {
          const nextRow = rowIndices[merge?.afterRow ?? row + 1] ?? rowPositions.length - 1;
          const nextCol = colIndices[merge?.afterCol ?? col + 1] ?? colPositions.length - 1;
          if (!style.borders?.bottom?.width) border(td, 'bottom', styleFor(sheet, nextRow, sourceCol).borders?.top);
          if (!style.borders?.right?.width) border(td, 'right', styleFor(sheet, sourceRow, nextCol).borders?.left);
        }
        const box = node('div', 'op-cell-content'), text = node('span', 'op-cell-text', display);
        box.style.alignItems = ({ top: 'flex-start', middle: 'center', bottom: 'flex-end' }[style.verticalAlign]) || 'flex-end';
        box.style.justifyContent = ({ center: 'center', right: 'flex-end' }[td.style.textAlign]) || 'flex-start';
        text.style.whiteSpace = style.wrapText ? 'pre-wrap' : 'pre';
        if (!formatted && sheet.rowHeights?.[sourceRow] === undefined && sheet.defaultRowHeight === undefined) text.style.lineHeight = '1.5';
        if (style.wrapText) text.style.width = '100%';
        if (style.rotation && style.rotation !== 255) text.style.transform = `rotate(${-style.rotation}deg)`;
        if (style.rotation === 255) text.style.writingMode = 'vertical-rl';
        if (style.indent) box.style.paddingLeft = `${style.indent * (style.fontSize || 14) * 0.75}px`;
        if (merge) Object.assign(box.style, { left: `${columns[merge.firstCol] - columns[merge.left] + 4}px`, top: `${rows[merge.firstRow] - rows[merge.top] + 2}px`, right: 'auto', bottom: 'auto', width: `${Math.max(0, columns[merge.afterCol] - columns[merge.firstCol] - 8)}px`, height: `${Math.max(0, rows[merge.afterRow] - rows[merge.firstRow] - 4)}px` });
        box.append(text); td.append(box);
        const address = `${columnName(sourceCol)}${sourceRow + 1}`;
        td.title = cell?.formula ? `${address}: =${cell.formula}\n${value}` : `${address}: ${value}`;
        td.dataset.address = address;
        if (address === selectedAddress) td.classList.add('op-selected');
        tr.append(td);
      }
      tbody.append(tr);
    }
    table.append(tbody);
    table.onclick = event => {
      const cell = event.target.closest('td');
      if (!cell || !table.contains(cell)) return;
      table.querySelector('.op-selected')?.classList.remove('op-selected'); cell.classList.add('op-selected'); selectedAddress = cell.dataset.address;
    };
    spacer.replaceChildren(table);
    if (!rowCount) { const empty = node('div', 'op-empty', '空工作表'); Object.assign(empty.style, { position: 'absolute', left: '60px', top: '40px' }); spacer.append(empty); }
  }
  function schedule() {
    if (disposed || frame) return;
    frame = requestAnimationFrame(() => { frame = 0; draw(); });
  }
  viewport.addEventListener('scroll', schedule, { passive: true });
  const resize = new ResizeObserver(schedule); resize.observe(viewport);
  const zoom = new MutationObserver(schedule); zoom.observe(content, { attributes: true, attributeFilter: ['style'] });
  draw();
  return { fitWidth: 0, dispose() { disposed = true; if (frame) cancelAnimationFrame(frame); viewport.removeEventListener('scroll', schedule); resize.disconnect(); zoom.disconnect(); layouts.clear(); spacer.remove(); auxiliary.replaceChildren(); } };
}
