import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { parseCsv } from '../src/parsers/csv.js';

const values = (model) => model.sheets[0].rows.map((row) => row.map((cell) => cell.value));
const parse = (text, options) => parseCsv(Buffer.from(text), 'csv', 'input.csv', options);

test('CSV preserves quoted commas, escaped quotes and embedded newlines', () => {
  assert.deepEqual(values(parse('名称,备注\r\n样本,"逗号,双引号""内容"""\r\n多行,"第一行\n第二行"\r\n')), [
    ['名称', '备注'], ['样本', '逗号,双引号"内容"'], ['多行', '第一行\n第二行'],
  ]);
});

test('CSV detects semicolons outside quotes and preserves empty trailing fields', () => {
  const model = parse('A;B;C\n1;"a,b,c";\n');
  assert.equal(model.delimiter, ';');
  assert.deepEqual(values(model), [['A', 'B', 'C'], ['1', 'a,b,c', '']]);
});

test('CSV respects an explicit delimiter and never interprets formula-looking text', () => {
  const model = parse('=1+1|<script>alert(1)</script>|00042', { delimiter: '|' });
  assert.deepEqual(values(model), [['=1+1', '<script>alert(1)</script>', '00042']]);
});

test('CSV reads UTF-8 and UTF-16LE BOM development samples', async () => {
  const utf8 = parseCsv(await readFile(new URL('../fixtures/sample.csv', import.meta.url)));
  assert.equal(utf8.sheets[0].rows[0][0].value, '项目');
  const utf16 = parseCsv(await readFile(new URL('../fixtures/sample-utf16.csv', import.meta.url)));
  assert.equal(utf16.delimiter, '\t');
  assert.deepEqual(values(utf16), [['名称', '数量'], ['中文编码', '12']]);
});

test('CSV rejects malformed quoting and invalid delimiters', () => {
  assert.throws(() => parse('A,"unclosed'), /未闭合/);
  assert.throws(() => parse('A,"closed"x'), /引号/);
  assert.throws(() => parse('A,B', { delimiter: '||' }), /分隔符/);
});

test('empty CSV produces an empty sheet instead of a spurious data row', () => {
  assert.deepEqual(values(parse('')), []);
});
