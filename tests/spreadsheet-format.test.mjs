import test from 'node:test';
import assert from 'node:assert/strict';
import { formatSpreadsheetValue as display } from '../src/spreadsheet-format.js';

test('XLS number display preserves identifiers and formats percentages and decimals', () => {
  assert.equal(display('00001001401', '0'), '00001001401');
  assert.equal(display(123, '000000'), '000123');
  assert.equal(display(0.125, '0.0%'), '12.5%');
  assert.equal(display(1234.5, '#,##0.00'), '1,234.50');
  assert.equal(display(-1234.5, '#,##0.00;[Red](#,##0.00)'), '(1,234.50)');
  assert.equal(display(0, '0.0;-0.0;"—"'), '—');
  assert.equal(display(1234567, '0.00E+00'), '1.23E+06');
});
test('XLS dates respect both workbook date systems and Excel serial 60', () => {
  assert.equal(display(1, 'yyyy-mm-dd'), '1900-01-01');
  assert.equal(display(60, 'yyyy-mm-dd'), '1900-02-29');
  assert.equal(display(61, 'yyyy-mm-dd'), '1900-03-01');
  assert.equal(display(0, 'yyyy-mm-dd', { date1904: true }), '1904-01-01');
  assert.equal(display(0.5, 'h:mm:ss'), '12:00:00');
  assert.equal(display(0.5, 'h:mm AM/PM'), '12:00 PM');
  assert.equal(display(1.5, '[h]:mm:ss'), '36:00:00');
});
test('XLS display distinguishes quoted literals and falls back for unsupported conditions', () => {
  assert.equal(display(42, '0" meters"'), '42 meters');
  assert.equal(display(1234.5, '"¥"#,##0.00'), '¥1,234.50');
  assert.equal(display(1234, '[>=100]0.0'), '1234');
  assert.equal(display(12, 'General'), '12');
});
