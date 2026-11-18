import { elements, children, first, attr, relationshipId, numberAttr, warn, xmlPart, relationships, relatedPath, onOff, rgb, richText, themeColors } from './ooxml-common.js';

const MAX_ROWS = 50000;
const MAX_COLUMNS = 1024;
const MAX_CELLS = 1000000;
const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'mm-dd-yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 48: '##0.0E+0', 49: '@',
};
const INDEXED_COLORS = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF'];
const BORDER_STYLES = new Map([
  ['hair', ['solid', 0.5]], ['thin', ['solid', 1]], ['medium', ['solid', 2]], ['thick', ['solid', 3]],
  ['dotted', ['dotted', 1]], ['dashed', ['dashed', 1]], ['double', ['double', 3]],
  ['dashDot', ['dashed', 1]], ['dashDotDot', ['dashed', 1]], ['mediumDashed', ['dashed', 2]],
  ['mediumDashDot', ['dashed', 2]], ['mediumDashDotDot', ['dashed', 2]], ['slantDashDot', ['dashed', 2]],
]);

const xmlTrue = value => value === '1' || value === 'true';

function coordinate(reference) {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(reference);
  if (!match) throw new Error(`无效的 Excel 单元格坐标：${reference}`);
  let col = 0;
  for (const character of match[1].toUpperCase()) col = col * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[2]) - 1;
  col -= 1;
  if (row >= MAX_ROWS || col >= MAX_COLUMNS) throw new Error(`工作表超出当前预览范围：最多 ${MAX_ROWS} 行、${MAX_COLUMNS} 列。`);
  return { row, col };
}

export function parseWorkbook(parts, path, warnings) {
  const workbook = xmlPart(parts, path);
  const date1904 = xmlTrue(attr(first(workbook.documentElement, 'workbookPr'), 'date1904'));
  const rels = relationships(parts, path, warnings);
  const stringsPath = relatedPath(rels, 'sharedStrings');
  const sharedStrings = stringsPath ? elements(xmlPart(parts, stringsPath), 'si').map(richText) : [];
  const stylesPath = relatedPath(rels, 'styles');
  const stylesXml = stylesPath ? xmlPart(parts, stylesPath) : null;
  const theme = themeColors(parts, rels);
  const themeOrder = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  const formats = { ...BUILTIN_FORMATS };
  for (const format of elements(stylesXml, 'numFmt')) formats[attr(format, 'numFmtId')] = attr(format, 'formatCode');

  function color(node) {
    const direct = rgb(attr(node, 'rgb'));
    if (direct) return direct;
    if (attr(node, 'theme') !== '') return theme[themeOrder[numberAttr(node, 'theme')]];
    const index = numberAttr(node, 'indexed', -1);
    return rgb(INDEXED_COLORS[index < 8 ? index : index < 16 ? index - 8 : -1] || '');
  }

  const fonts = children(first(stylesXml?.documentElement, 'fonts'), 'font').map(font => ({
    bold: onOff(first(font, 'b')), italic: onOff(first(font, 'i')), color: color(first(font, 'color')),
    fontFamily: attr(first(font, 'name'), 'val') || undefined,
    fontSize: numberAttr(first(font, 'sz'), 'val') * 4 / 3 || undefined,
    underline: onOff(first(font, 'u')) && attr(first(font, 'u'), 'val') !== 'none', strike: onOff(first(font, 'strike')),
  }));
  const fills = children(first(stylesXml?.documentElement, 'fills'), 'fill').map(fill => {
    const pattern = first(fill, 'patternFill');
    return attr(pattern, 'patternType') === 'solid' ? color(first(pattern, 'fgColor')) : undefined;
  });
  const borders = children(first(stylesXml?.documentElement, 'borders'), 'border').map(border => {
    const result = {};
    for (const side of ['top', 'right', 'bottom', 'left']) {
      const edge = first(border, side);
      const line = BORDER_STYLES.get(attr(edge, 'style'));
      if (line) result[side] = { style: line[0], width: line[1], color: color(first(edge, 'color')) };
    }
    return result;
  });
  const styles = children(first(stylesXml?.documentElement, 'cellXfs'), 'xf').map(xf => {
    const alignment = first(xf, 'alignment');
    const horizontal = attr(alignment, 'horizontal');
    const vertical = attr(alignment, 'vertical');
    const rotation = numberAttr(alignment, 'textRotation');
    return {
      ...fonts[numberAttr(xf, 'fontId')],
      background: fills[numberAttr(xf, 'fillId')],
      format: formats[numberAttr(xf, 'numFmtId')],
      borders: borders[numberAttr(xf, 'borderId')],
      align: ['left', 'center', 'right', 'justify'].includes(horizontal) ? horizontal : horizontal === 'centerContinuous' ? 'center' : undefined,
      verticalAlign: vertical === 'center' ? 'middle' : ['top', 'bottom'].includes(vertical) ? vertical : undefined,
      wrapText: xmlTrue(attr(alignment, 'wrapText')),
      rotation: rotation >= 0 && (rotation <= 90 || rotation === 255) ? rotation : rotation > 90 && rotation <= 180 ? 90 - rotation : undefined,
      indent: Math.max(0, Math.min(250, numberAttr(alignment, 'indent'))),
    };
  });
  let totalCells = 0;
  const sheets = children(first(workbook.documentElement, 'sheets'), 'sheet').map((sheet, sheetIndex) => {
    if (['hidden', 'veryHidden'].includes(attr(sheet, 'state'))) return null;
    const target = rels.get(relationshipId(sheet));
    if (!target?.target) throw new Error('工作表关系缺失或指向外部资源。');
    if (!target.type.endsWith('/worksheet')) {
      warn(warnings, '图表工作表和对话框工作表尚不支持。');
      return null;
    }
    const doc = xmlPart(parts, target.target);
    if (doc.documentElement.localName !== 'worksheet') throw new Error('工作表 XML 类型无效。');
    const formatProperties = first(doc.documentElement, 'sheetFormatPr');
    const height = numberAttr(formatProperties, 'defaultRowHeight', -1);
    const defaultRowHeight = height > 0 ? Math.min(2000, height * 4 / 3) : undefined;
    const rows = [], rowHeights = [], hiddenRows = [], rowStyles = [];
    let implicitRow = 0;
    for (const row of children(first(doc.documentElement, 'sheetData'), 'row')) {
      const rowIndex = numberAttr(row, 'r', implicitRow + 1) - 1;
      if (!Number.isInteger(rowIndex) || rowIndex < 0 || rowIndex >= MAX_ROWS) throw new Error(`工作表行号无效或超过 ${MAX_ROWS} 行。`);
      implicitRow = rowIndex + 1;
      if (rows[rowIndex]) throw new Error('工作表出现重复行。');
      rows[rowIndex] = [];
      if (xmlTrue(attr(row, 'hidden'))) hiddenRows[rowIndex] = true;
      const height = numberAttr(row, 'ht', -1);
      if (height >= 0) rowHeights[rowIndex] = Math.min(2000, height * 4 / 3);
      if (xmlTrue(attr(row, 'customFormat'))) rowStyles[rowIndex] = styles[numberAttr(row, 's')];
      let implicitCol = 0;
      for (const cell of children(row, 'c')) {
        totalCells += 1;
        if (totalCells > MAX_CELLS) throw new Error(`工作簿超过当前 ${MAX_CELLS} 个数据单元格的预览上限。`);
        const reference = attr(cell, 'r');
        const position = reference ? coordinate(reference) : { row: rowIndex, col: implicitCol };
        if (position.row !== rowIndex || position.col >= MAX_COLUMNS) throw new Error('单元格坐标与所在行不一致，或超出列上限。');
        if (rows[rowIndex][position.col]) throw new Error('工作表出现重复单元格。');
        implicitCol = position.col + 1;
        const type = attr(cell, 't');
        const cached = first(cell, 'v');
        let value = cached?.textContent ?? '';
        if (type === 's') {
          const index = Number(value);
          if (!cached || !/^\d+$/.test(value) || index >= sharedStrings.length) throw new Error('Excel 共享字符串索引无效。');
          value = sharedStrings[index];
        } else if (type === 'inlineStr') value = richText(first(cell, 'is'));
        else if (type === 'b') {
          if (value !== '' && !['0', '1'].includes(value)) throw new Error('Excel 布尔单元格内容无效。');
          value = value === '' ? '' : value === '1';
        }
        else if (!type || type === 'n') {
          if (value !== '' && !Number.isFinite(Number(value))) throw new Error('Excel 数值单元格内容无效。');
          value = value === '' ? '' : Number(value);
        }
        const item = { value };
        const formula = first(cell, 'f');
        if (formula) {
          item.formula = formula.textContent;
          warn(warnings, '公式仅显示文件内缓存结果，当前不重新计算；缓存可能与实际结果不同。');
          if (!cached) warn(warnings, '部分公式没有缓存结果，显示为空白。');
          if (!formula.textContent) warn(warnings, '共享公式/数组公式的展开表达式尚未还原。');
        }
        const styleIndex = numberAttr(cell, 's', -1);
        if (styleIndex >= 0) {
          if (!Number.isInteger(styleIndex) || !styles[styleIndex]) throw new Error('Excel 单元格样式索引无效。');
          item.style = { ...styles[styleIndex] };
        }
        rows[rowIndex][position.col] = item;
      }
    }
    const merges = elements(doc, 'mergeCell').map(merge => {
      const range = attr(merge, 'ref').split(':');
      const start = coordinate(range[0]);
      const end = coordinate(range[1] || range[0]);
      if (end.row < start.row || end.col < start.col) throw new Error('Excel 合并区域无效。');
      return { startRow: start.row, startCol: start.col, endRow: end.row, endCol: end.col };
    });
    const columnWidths = [], hiddenColumns = [], columnStyles = [];
    for (const column of children(doc.documentElement, 'cols').flatMap(group => children(group, 'col'))) {
      const min = numberAttr(column, 'min');
      const max = numberAttr(column, 'max');
      // Formatting may describe the whole native 16,384-column grid; only visible parsed columns matter.
      if (!Number.isInteger(min) || !Number.isInteger(max) || min < 1 || max < min || max > 16384) throw new Error('Excel 列定义无效。');
      const width = numberAttr(column, 'width', -1);
      const hidden = xmlTrue(attr(column, 'hidden'));
      const style = styles[numberAttr(column, 'style', -1)];
      for (let index = min - 1; index < Math.min(max, MAX_COLUMNS); index += 1) {
        if (width >= 0) columnWidths[index] = width * 7 + 5;
        if (hidden || hiddenColumns[index]) hiddenColumns[index] = hidden;
        if (style) columnStyles[index] = style;
      }
    }
    if (elements(doc, 'drawing').length || elements(doc, 'legacyDrawing').length) warn(warnings, '工作表中的图表、图片、批注绘图尚未显示。');
    if (elements(doc, 'conditionalFormatting').length) warn(warnings, '条件格式尚未计算。');
    if (elements(doc, 'autoFilter').length) warn(warnings, '自动筛选保留文件中已隐藏的行，预览中不可调整筛选条件。');
    return { name: attr(sheet, 'name') || `Sheet${sheetIndex + 1}`, rows, merges, columnWidths, rowHeights, defaultRowHeight,
      hiddenRows, hiddenColumns, rowStyles, columnStyles, defaultStyle: styles[0], date1904 };
  }).filter(Boolean);
  if (!sheets.length) throw new Error('工作簿中没有可预览的普通工作表。');
  warn(warnings, 'XLSX 仅还原基本单元格样式；主题色变换、复杂数字格式、打印分页和样式继承可能不同。');
  return { kind: 'spreadsheet', sheets };
}
