import { TWIPS_PER_PIXEL, children, first, attr, numberAttr, onOff, rgb, warn } from './ooxml-common.js';

function count(node, name, fallback = 0) {
  const value = numberAttr(node, name, fallback);
  if (!Number.isInteger(value) || value < 0 || value > 1024) throw new Error('Word 表格列数或合并列数无效。');
  return value;
}

function length(node, fallbackType = 'dxa', signed = false) {
  if (!node) return {};
  const type = attr(node, 'type') || fallbackType, raw = attr(node, 'w');
  if (type === 'auto' || type === 'nil') return {};
  const value = type === 'pct' && raw.endsWith('%') ? Number(raw.slice(0, -1)) * 50 : numberAttr(node, 'w');
  if (!Number.isFinite(value) || value < 0 && !signed) throw new Error('Word 表格尺寸无效。');
  return type === 'dxa' ? { width: value / TWIPS_PER_PIXEL } : type === 'pct' ? { widthPercent: value / 50 } : {};
}

function margins(node) {
  if (!node) return undefined;
  const result = {};
  for (const side of children(node)) {
    const key = ({ start: 'left', end: 'right' })[side.localName] || side.localName;
    if (!['top', 'right', 'bottom', 'left'].includes(key)) continue;
    const value = length(side);
    if (value.width !== undefined) result[key] = value.width;
  }
  return result;
}

function borders(node) {
  if (!node) return undefined;
  const result = {};
  for (const side of children(node)) {
    const key = ({ start: 'left', end: 'right' })[side.localName] || side.localName;
    if (!['top', 'right', 'bottom', 'left', 'insideH', 'insideV'].includes(key)) continue;
    const style = attr(side, 'val') || 'single';
    const width = numberAttr(side, 'sz', 4) / 6;
    if (width < 0) throw new Error('Word 表格边框宽度无效。');
    result[key] = { style: ['none', 'nil'].includes(style) ? 'none' : style, width: ['none', 'nil'].includes(style) ? 0 : width, color: rgb(attr(side, 'color')) || '#000000' };
  }
  return result;
}

export function parseWordTable(node, { blocks, sectionIndex = 0, warnings = [] }) {
  const properties = first(node, 'tblPr');
  const table = {
    type: 'table', rows: [], columnWidths: children(first(node, 'tblGrid'), 'gridCol').map(column => {
      const value = numberAttr(column, 'w');
      if (value < 0) throw new Error('Word 表格列宽无效。');
      return value / TWIPS_PER_PIXEL;
    }), rowProperties: [], sectionIndex,
    ...length(first(properties, 'tblW')),
  };
  if (table.columnWidths.length > 1024) throw new Error('Word 表格列数超过当前上限。');
  const alignment = attr(first(properties, 'jc'), 'val');
  if (alignment) table.align = ({ start: 'left', end: 'right' })[alignment] || alignment;
  const indent = length(first(properties, 'tblInd'), 'dxa', true);
  if (indent.width !== undefined) table.indent = indent.width;
  const layout = attr(first(properties, 'tblLayout'), 'type');
  if (layout) table.layout = layout;
  const tableBorders = borders(first(properties, 'tblBorders'));
  if (tableBorders) table.borders = tableBorders;
  const tableMargins = margins(first(properties, 'tblCellMar'));
  if (tableMargins) table.cellMargins = tableMargins;
  const spacing = length(first(properties, 'tblCellSpacing'));
  if (spacing.width !== undefined) table.cellSpacing = spacing.width;

  // Only cells from the immediately preceding row may continue a vMerge.
  // Physical grid columns, including a horizontal span, identify the anchor.
  let active = new Map();
  for (const row of children(node, 'tr')) {
    const rowProperties = first(row, 'trPr'), height = first(rowProperties, 'trHeight'), settings = {};
    if (height) {
      const value = numberAttr(height, 'val');
      if (value < 0) throw new Error('Word 表格行高无效。');
      settings.height = value / TWIPS_PER_PIXEL; settings.heightRule = attr(height, 'hRule') || 'atLeast';
    }
    const before = count(first(rowProperties, 'gridBefore'), 'val'), after = count(first(rowProperties, 'gridAfter'), 'val');
    if (before) settings.gridBefore = before;
    if (after) settings.gridAfter = after;
    const cells = [], continued = new Map();
    let column = before;
    for (const cell of children(row, 'tc')) {
      const own = first(cell, 'tcPr'), span = Math.max(1, count(first(own, 'gridSpan'), 'val', 1));
      if (column + span + after > 1024) throw new Error('Word 表格列数超过当前上限。');
      const merge = first(own, 'vMerge'), continuation = merge && attr(merge, 'val') !== 'restart', previous = active.get(column);
      if (continuation && previous && previous.span === span) {
        previous.cell.rowSpan += 1;
        const ownBorders = borders(first(own, 'tcBorders'));
        if (ownBorders?.bottom) previous.cell.borders = { ...previous.cell.borders, bottom: ownBorders.bottom };
        continued.set(column, previous); column += span; continue;
      }
      if (continuation) warn(warnings, 'Word 表格的纵向合并没有相邻且同宽的起始单元格，按独立单元格显示。');
      const value = { blocks: blocks(cell), column, ...(span > 1 ? { colSpan: span } : {}), ...length(first(own, 'tcW')) };
      const ownBorders = borders(first(own, 'tcBorders')), ownMargins = margins(first(own, 'tcMar'));
      if (ownBorders) value.borders = ownBorders;
      if (ownMargins) value.margins = ownMargins;
      const vertical = attr(first(own, 'vAlign'), 'val');
      if (vertical) value.verticalAlign = vertical;
      const shading = first(own, 'shd'), fill = rgb(attr(shading, 'fill'));
      if (shading && attr(shading, 'val') === 'nil') value.shading = 'transparent';
      else if (fill) value.shading = fill;
      const nowrap = first(own, 'noWrap'); if (nowrap) value.noWrap = onOff(nowrap);
      if (merge && !continuation) { value.rowSpan = 1; continued.set(column, { cell: value, span }); }
      cells.push(value); column += span;
    }
    active = continued;
    table.rows.push(cells); table.rowProperties.push(settings);
  }
  return table;
}
