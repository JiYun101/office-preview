import test from 'node:test';
import assert from 'node:assert/strict';
import { parseOoxml } from '../src/parsers/ooxml.js';
import { readZip } from '../src/containers/zip.js';
import { buildDocx, buildXlsx, buildPptx, makeZip } from '../scripts/generate-fixtures.mjs';

const encode = source => new TextEncoder().encode(source);
const decode = bytes => new TextDecoder().decode(bytes);
async function rewrite(bytes, update) {
  const entries = await readZip(bytes);
  update(entries);
  return makeZip(Object.fromEntries(entries));
}

test('DOCX reads paragraphs, direct formatting, table, internal image and page units without DOMParser', async () => {
  assert.equal(typeof globalThis.DOMParser, 'undefined');
  const model = await parseOoxml(buildDocx(), '.DOCX', 'sample.docx');
  assert.equal(model.kind, 'document');
  assert.equal(model.format, 'docx');
  assert.equal(model.blocks[0].align, 'center');
  assert.equal(model.blocks[0].runs[0].text, '自主研发 Office 预览套件');
  assert.equal(model.blocks[0].runs[0].bold, true);
  assert.equal(model.blocks[0].runs[0].color, '#184C4C');
  assert.equal(model.blocks[0].runs[0].fontSize, 40 * 2 / 3);
  const table = model.blocks.find(block => block.type === 'table');
  assert.equal(table.rows.length, 2);
  assert.equal(table.rows[1][0].blocks[0].runs[0].text, 'DOCX');
  const image = model.blocks.flatMap(block => block.runs || []).find(run => run.image).image;
  assert.equal(image.mime, 'image/png');
  assert.equal(image.width, 96);
  assert.equal(image.height, 96);
  assert.ok(image.bytes.length > 0);
  assert.equal(model.page.width, 11906 / 15);
  assert.ok(model.warnings.some(value => value.includes('精确分页')));
});

test('XLSX preserves sheet names, sparse coordinates, strings, merges and cached formulas', async () => {
  const model = await parseOoxml(buildXlsx(), 'xlsx');
  assert.equal(model.kind, 'spreadsheet');
  assert.deepEqual(model.sheets.map(sheet => sheet.name), ['产品清单', '说明']);
  const sheet = model.sheets[0];
  assert.equal(sheet.rows[0][0].value, 'Office 预览开发样本');
  assert.equal(sheet.rows[2][0].value, '自研内核');
  assert.equal(sheet.rows[2][3].value, 1206);
  assert.equal(sheet.rows[2][3].formula, 'B3*C3');
  assert.equal(sheet.rows[3][1].value, true);
  assert.equal(sheet.rows[3][2], undefined);
  assert.equal(sheet.rows[1][0].style.bold, true);
  assert.equal(sheet.rows[1][0].style.background, '#24645B');
  assert.deepEqual(sheet.merges[0], { startRow: 0, startCol: 0, endRow: 0, endCol: 3 });
  assert.ok(model.warnings.some(value => value.includes('缓存')));
});

test('PPTX uses relationship IDs instead of numeric IDs and honors the presentation ordering', async () => {
  const bytes = await rewrite(buildPptx(), entries => {
    const source = decode(entries.get('ppt/presentation.xml')).replace('r:id="first"', 'r:id="temp"').replace('r:id="second"', 'r:id="first"').replace('r:id="temp"', 'r:id="second"');
    entries.set('ppt/presentation.xml', encode(source));
  });
  const model = await parseOoxml(bytes, 'pptx');
  assert.equal(model.width, 1280);
  assert.equal(model.height, 720);
  assert.equal(model.slides.length, 2);
  assert.equal(model.slides[0].elements[0].runs[0].text, '真实读取基础结构');
  assert.equal(model.slides[1].elements[0].runs[0].text, 'Office，嵌入你的软件');
  assert.equal(model.slides[0].elements[0].x, 81.6); // DrawingML default left inset: 0.1 inch.
  assert.equal(model.slides[0].background, '#F6F3EC');
  assert.equal(model.slides[0].elements.find(item => item.type === 'shape').fill, '#B55A3A');
});

test('PPTX text uses body insets, vertical anchoring and paragraph spacing', async () => {
  const bytes = await rewrite(buildPptx(), entries => {
    const path = 'ppt/slides/slide1.xml';
    const source = decode(entries.get(path)).replace('<a:bodyPr/>', '<a:bodyPr lIns="19050" rIns="57150" tIns="38100" bIns="76200" anchor="ctr"/>');
    entries.set(path, encode(source));
  });
  const model = await parseOoxml(bytes, 'pptx');
  const text = model.slides[0].elements.find(element => element.type === 'text');
  assert.equal(text.x, 74);
  assert.equal(text.verticalAlign, 'middle');
  assert.ok(text.paragraphs.length > 0);
  assert.equal(text.runs.map(run => run.text).join(''), text.paragraphs.flatMap(paragraph => paragraph.runs).map(run => run.text).join(''));
});

test('DOCX resolves main part and media relative paths and skips external image relationships', async () => {
  const relocated = await rewrite(buildDocx(), entries => {
    entries.set('_rels/.rels', encode(decode(entries.get('_rels/.rels')).replace('word/document.xml', './docs/main.xml')));
    entries.set('docs/main.xml', entries.get('word/document.xml'));
    entries.set('docs/_rels/main.xml.rels', encode(decode(entries.get('word/_rels/document.xml.rels')).replace('media/pixel.png', '../word/media/pixel.png')));
    entries.delete('word/document.xml');
    entries.delete('word/_rels/document.xml.rels');
  });
  const model = await parseOoxml(relocated, 'docx');
  assert.ok(model.blocks.some(block => block.runs?.some(run => run.image)));
  const external = await rewrite(buildDocx(), entries => {
    entries.set('word/_rels/document.xml.rels', encode(decode(entries.get('word/_rels/document.xml.rels')).replace('Target="media/pixel.png"', 'TargetMode="External" Target="https://example.invalid/image.png"')));
  });
  const safe = await parseOoxml(external, 'docx');
  assert.equal(safe.blocks.some(block => block.runs?.some(run => run.image)), false);
  assert.ok(safe.warnings.some(value => value.includes('外部')));
});

test('OOXML rejects malformed XML, wrong formats, missing required parts and oversized sheet coordinates', async () => {
  await assert.rejects(parseOoxml(buildDocx(), 'xlsx'), /不一致/);
  const broken = await rewrite(buildDocx(), entries => entries.set('word/document.xml', encode('<document><body></document>')));
  await assert.rejects(parseOoxml(broken, 'docx'), /XML/);
  const missing = await rewrite(buildPptx(), entries => entries.delete('ppt/slides/slide1.xml'));
  await assert.rejects(parseOoxml(missing, 'pptx'), /缺少/);
  const oversized = await rewrite(buildXlsx(), entries => {
    const source = decode(entries.get('xl/worksheets/sheet1.xml')).replace('<row r="1"><c r="A1"', '<row r="50001"><c r="A50001"');
    entries.set('xl/worksheets/sheet1.xml', encode(source));
  });
  await assert.rejects(parseOoxml(oversized, 'xlsx'), /50000/);
});
