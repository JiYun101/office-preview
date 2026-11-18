import test from 'node:test';
import assert from 'node:assert/strict';
import { capabilities, parseDocument } from '../src/parse.js';
import { readZip } from '../src/containers/zip.js';
import { buildDocx, buildPptx, buildXlsx, makeZip } from '../scripts/generate-fixtures.mjs';
import { formatSpreadsheetValue } from '../src/spreadsheet-format.js';

async function workbookWithMerges(ranges) {
  const parts = await readZip(buildXlsx());
  const name = 'xl/worksheets/sheet1.xml';
  const original = new TextDecoder().decode(parts.get(name));
  const merges = `<mergeCells count="${ranges.length}">${ranges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;
  const modified = original.replace(/<mergeCells\b[\s\S]*?<\/mergeCells>/, merges);
  assert.notEqual(modified, original, 'development fixture merge replacement failed');
  parts.set(name, new TextEncoder().encode(modified));
  return makeZip(Object.fromEntries(parts));
}

test('parseDocument rejects overlapping XLSX merges beyond the first two visible segments', async () => {
  const bytes = await workbookWithMerges(['A401:B402', 'B402:C403']);
  await assert.rejects(parseDocument(bytes, { name: 'overlapping.xlsx' }), /重叠.*合并|合并.*重叠/);
});

test('parseDocument accepts adjacent, non-overlapping XLSX merges beyond row 400', async () => {
  const bytes = await workbookWithMerges(['A401:A402', 'B401:B402']);
  const model = await parseDocument(bytes, { name: 'adjacent.xlsx' });
  assert.equal(model.kind, 'spreadsheet');
  assert.deepEqual(model.sheets[0].merges, [
    { startRow: 400, startCol: 0, endRow: 401, endCol: 0 },
    { startRow: 400, startCol: 1, endRow: 401, endCol: 1 },
  ]);
});

test('public parsing accepts OOXML macro, template and slide-show extensions as static documents', async () => {
  for (const [kind, build, formats] of [
    ['document', buildDocx, ['docm', 'dotx', 'dotm']],
    ['spreadsheet', buildXlsx, ['xlsm', 'xltx', 'xltm']],
    ['presentation', buildPptx, ['pptm', 'ppsx', 'ppsm', 'potx', 'potm']],
  ]) {
    const parts = await readZip(build());
    const macroPath = { document: 'word', spreadsheet: 'xl', presentation: 'ppt' }[kind] + '/vbaProject.bin';
    const macroFile = makeZip({ ...Object.fromEntries(parts), [macroPath]: new Uint8Array([0, 1, 2]) });
    for (const format of formats) {
      const model = await parseDocument(format.endsWith('m') ? macroFile : build(), { name: `example.${format.toUpperCase()}` });
      assert.equal(model.kind, kind);
      assert.equal(model.format, format);
      assert.ok(capabilities.some(group => group.formats.includes(format)));
      if (format.endsWith('m')) assert.ok(model.warnings.some(warning => warning.includes('宏代码')));
    }
  }
  await assert.rejects(parseDocument(buildDocx(), { name: 'wrong.xlsm' }), /不一致/);
  await assert.rejects(parseDocument(new Uint8Array([0xd0, 0xcf]), { name: 'encrypted.dotx' }), /加密/);
});

test('XLSM applies XLSX styles and visibility metadata while preserving formula caches', async () => {
  const parts = await readZip(buildXlsx());
  const update = (path, fn) => parts.set(path, new TextEncoder().encode(fn(new TextDecoder().decode(parts.get(path)))));
  update('xl/workbook.xml', xml => xml.replace('<sheets>', '<workbookPr date1904="true"/><sheets>').replace('sheetId="2"', 'sheetId="2" state="veryHidden"'));
  update('xl/styles.xml', xml => xml
    .replace('<fonts', '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts')
    .replace('<borders count="1"><border/></borders>', '<borders count="2"><border/><border><bottom style="thin"><color rgb="FF113355"/></bottom></border></borders>')
    .replace('cellXfs count="2"', 'cellXfs count="6"')
    .replace('</cellXfs>', '<xf numFmtId="164"/><xf numFmtId="10"/><xf numFmtId="4" borderId="1"><alignment horizontal="center" vertical="center" wrapText="true"/></xf><xf numFmtId="999"/></cellXfs>'));
  update('xl/worksheets/sheet1.xml', xml => xml
    .replace('<cols>', '<sheetFormatPr defaultRowHeight="18"/><cols>')
    .replace('min="2" max="4"', 'min="2" max="2" hidden="1"')
    .replace('<row r="2">', '<row r="2" hidden="true">')
    .replace('<row r="3">', '<row r="3" ht="30">')
    .replace('<c r="B3"><v>12</v>', '<c r="B3" s="3"><v>0.125</v>')
    .replace('<c r="C3">', '<c r="C3" s="5">')
    .replace('<c r="D3">', '<c r="D3" s="4">')
    .replace('<c r="D4"><v>42</v>', '<c r="D4" s="2"><v>0</v>'));
  const model = await parseDocument(makeZip(Object.fromEntries(parts)), { name: 'sample.xlsm' });
  assert.deepEqual(model.sheets.map(sheet => sheet.name), ['产品清单']);
  const [sheet] = model.sheets;
  assert.equal(sheet.defaultRowHeight, 24);
  assert.equal(sheet.rowHeights[2], 40);
  assert.equal(sheet.hiddenRows[1], true);
  assert.equal(sheet.hiddenColumns[1], true);
  const display = ({ value, style }) => formatSpreadsheetValue(value, style.format, { date1904: sheet.date1904 });
  assert.equal(display(sheet.rows[2][1]), '12.50%');
  assert.equal(display(sheet.rows[2][2]), '100.5');
  assert.equal(display(sheet.rows[3][3]), '1904-01-01');
  const formula = sheet.rows[2][3];
  assert.equal(formula.value, 1206);
  assert.equal(formula.formula, 'B3*C3');
  assert.equal(display(formula), '1,206.00');
  assert.equal(formula.style.align, 'center');
  assert.equal(formula.style.verticalAlign, 'middle');
  assert.equal(formula.style.wrapText, true);
  assert.deepEqual(formula.style.borders.bottom, { style: 'solid', width: 1, color: '#113355' });
});
