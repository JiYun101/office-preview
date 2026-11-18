import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLegacy } from '../src/parsers/legacy.js';
import { readCompound } from '../src/containers/cfb.js';
import { makeCompound } from '../scripts/generate-fixtures.mjs';
import { buildStyledXls } from '../scripts/legacy-style-fixtures.mjs';

function editRecords(bytes, edit) {
  const stream = Buffer.from(readCompound(bytes).get('Workbook'));
  let offset = 0;
  while (offset + 4 <= stream.length) {
    const type = stream.readUInt16LE(offset), length = stream.readUInt16LE(offset + 2);
    edit(type, stream.subarray(offset + 4, offset + 4 + length));
    offset += 4 + length;
  }
  return makeCompound({ Workbook: stream });
}

test('BIFF8 FONT/XF restores title size, font, alignment and merged heading', () => {
  const sheet = parseLegacy(buildStyledXls(), 'xls').sheets[0];
  const title = sheet.rows[0][0];
  assert.equal(title.style.fontFamily, '宋体');
  assert.equal(title.style.fontSize, 32); // 24pt at CSS 96dpi.
  assert.equal(title.style.color, '#184c4c');
  assert.equal(title.style.align, 'center');
  assert.equal(title.style.verticalAlign, 'middle');
  assert.deepEqual(sheet.merges[0], { startRow: 0, endRow: 0, startCol: 0, endCol: 3 });
  assert.equal(sheet.defaultFont.fontSize, 16);
});

test('BIFF8 palette, solid fill and individual borders survive cell style resolution', () => {
  const model = parseLegacy(buildStyledXls(), 'xls'), header = model.sheets[0].rows[1][0].style;
  assert.equal(header.bold, true);
  assert.equal(header.background, '#e8f1ff');
  assert.equal(header.wrapText, true);
  for (const side of ['left','right','top','bottom']) assert.deepEqual(header.borders[side], { style: 'solid', width: 1, color: '#000000' });
});

test('BIFF8 cell XF inherits unmodified font/number groups while keeping its own fill/alignment', () => {
  const style = parseLegacy(buildStyledXls(), 'xls').sheets[0].rows[3][0].style;
  assert.equal(style.fontFamily, '宋体');
  assert.equal(style.fontSize, 16);
  assert.equal(style.format, 'General');
  assert.equal(style.background, '#e8f1ff');
  assert.equal(style.align, 'center');
  assert.equal(style.verticalAlign, 'middle');
});

test('BIFF8 FONT index 5 skips reserved index 4', () => {
  const style = parseLegacy(buildStyledXls(), 'xls').sheets[0].rows[3][1].style;
  assert.equal(style.fontFamily, 'Arial');
  assert.equal(style.fontSize, 14 * 96 / 72);
  assert.equal(style.italic, true);
  assert.equal(style.color, '#c43b44');
  assert.equal(style.align, 'left');
  assert.equal(style.indent, 1);
});

test('BIFF8 row heights, column widths and defaults retain their file units', () => {
  const sheet = parseLegacy(buildStyledXls(), 'xls').sheets[0];
  assert.deepEqual(sheet.rowHeights.slice(0,5), [48,40,32,30,30]);
  assert.deepEqual(sheet.columnWidthUnits.slice(0,4), [10,18,22,16].map(value => value * 256));
  assert.deepEqual(sheet.columnWidths.slice(0,4), [85,149,181,133]);
  assert.equal(sheet.defaultRowHeight, 24);
  assert.equal(sheet.defaultColumnWidthUnits, 9 * 256);
  assert.equal(sheet.defaultColumnWidth, 77);
});

test('BIFF8 FORMAT preserves raw numeric values and date system for display formatting', () => {
  const bytes = editRecords(buildStyledXls(), (type,data) => { if (type === 0x0022) data.writeUInt16LE(1); });
  const sheet = parseLegacy(bytes, 'xls').sheets[0];
  assert.equal(sheet.rows[2][1].value, 42);
  assert.equal(sheet.rows[2][1].style.format, '000000');
  assert.equal(sheet.rows[2][2].value, 46000);
  assert.equal(sheet.rows[2][2].style.format, 'yyyy-mm-dd');
  assert.equal(sheet.rows[2][3].value, 0.125);
  assert.equal(sheet.rows[2][3].style.format, '0.00%');
  assert.equal(sheet.date1904, true);
});

test('BIFF8 BLANK/MULBLANK keep formatted empty cells and resolved row/column defaults', () => {
  const bytes = editRecords(buildStyledXls(), (type,data) => {
    if (type === 0x0208 && data.readUInt16LE(0) === 4) { data[12] |= 0x80; data.writeUInt16LE(9,14); }
  });
  const sheet = parseLegacy(bytes, 'xls').sheets[0];
  for (const col of [2,3]) {
    assert.equal(sheet.rows[3][col].value, '');
    assert.equal(sheet.rows[3][col].style.background, '#e8f1ff');
    assert.equal(sheet.rows[3][col].style.bold, true);
  }
  assert.equal(sheet.rows[4][0].value, '');
  assert.equal(sheet.rows[4][0].style.fontFamily, 'Arial');
  assert.equal(sheet.rowStyles[4].fontFamily, 'Arial');
  assert.equal(sheet.columnStyles[1].fontFamily, '宋体');
  assert.equal(sheet.defaultStyle.fontFamily, '宋体');
});

test('BIFF8 hidden rows and column ranges retain their cell values and coordinates', () => {
  const { sheets: [sheet] } = parseLegacy(buildStyledXls({ hidden: true }), 'xls');
  assert.equal(sheet.hiddenRows[1], true);
  assert.equal(sheet.hiddenRows[2], false);
  assert.equal(sheet.hiddenColumns[1], true);
  assert.equal(sheet.rows[2][1].value, 42);
  assert.equal(sheet.rows[2][2].value, 46000);
  assert.equal(sheet.rowHeights[1], 40);
  assert.equal(sheet.merges[0].endCol, 3);
});

test('BIFF8 default zero-height hides empty row gaps while formatted rows remain visible', () => {
  const bytes = editRecords(buildStyledXls(), (type, data) => {
    if (type === 0x0225) data.writeUInt16LE(data.readUInt16LE(0) | 2);
    if (type === 0x0200) data.writeUInt32LE(7, 4);
    if ((type === 0x0208 || type === 0x0201) && data.readUInt16LE(0) === 4) data.writeUInt16LE(6);
  });
  const { sheets: [sheet] } = parseLegacy(bytes, 'xls');
  assert.deepEqual(sheet.hiddenRows, [false, false, false, false, true, true, false]);
  assert.equal(sheet.rows[6][0].value, '');
});

test('BIFF8 truncated font strings fail rather than reading neighboring records', () => {
  let changed = false;
  const bytes = editRecords(buildStyledXls(), (type,data) => {
    if (type === 0x0031 && !changed) { data[14] = 250; changed = true; }
  });
  assert.throws(() => parseLegacy(bytes, 'xls'), /越界|损坏/);
});
