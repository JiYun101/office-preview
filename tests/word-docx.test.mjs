import test from 'node:test';
import assert from 'node:assert/strict';
import { parseOoxml } from '../src/parsers/ooxml.js';
import { readZip } from '../src/containers/zip.js';
import { makeZip } from '../scripts/generate-fixtures.mjs';
import { buildStyledDocx, buildWordPageAreasDocx } from '../scripts/word-fixtures.mjs';

const encoder = new TextEncoder(), decoder = new TextDecoder();
const text = paragraph => paragraph.runs.map(run => run.text || '').join('');
async function updated(edit) {
  const parts = await readZip(buildStyledDocx());
  const change = (path, transform) => parts.set(path, encoder.encode(transform(decoder.decode(parts.get(path)))));
  edit(change, parts);
  return parseOoxml(makeZip(Object.fromEntries(parts)), 'docx');
}

test('DOCX resolves document defaults and inherited TOC paragraph/character styles', async () => {
  const model = await parseOoxml(buildStyledDocx(), 'docx');
  const heading = model.blocks[0];
  assert.equal(heading.isTocHeading, true);
  assert.equal(heading.style.fontSize, 32 * 2 / 3);
  assert.equal(heading.runs[0].fontFamily, '黑体');
  assert.equal(heading.runs[0].bold, true);
  assert.equal(heading.runs[0].color, '#2F75B5');
  assert.equal(heading.keepNext, true);
  assert.equal(model.defaultStyle.fontFamily, '宋体');
  assert.equal(model.defaultStyle.fontSize, 14);
  const emphasis = model.blocks.flatMap(block => block.runs || []).find(run => run.text === '字符样式继承');
  assert.equal(emphasis.italic, true);
  assert.equal(emphasis.color, '#C0504D');
});

test('DOCX TOC preserves cached field results, custom tab leaders and inherited indentation', async () => {
  const model = await parseOoxml(buildStyledDocx(), 'docx');
  const toc = model.blocks.filter(block => block.isToc);
  assert.deepEqual(toc.map(block => block.tocLevel), [1, 2, 3]);
  assert.deepEqual(toc.map(block => block.indentLeft), [24, 48, 72]);
  assert.deepEqual(toc.map(block => block.tabs), Array.from({ length: 3 }, () => [{ position: 600, align: 'right', leader: 'dot' }]));
  assert.deepEqual(toc.map(block => block.runs.filter(run => run.type === 'tab').length), [1, 1, 1]);
  assert.deepEqual(toc.map(block => block.runs.at(-1).text), ['5', '12', '23']);
  assert.equal(toc[0].runs.at(-1).fontFamily, 'Times New Roman');
  assert.equal(toc[0].runs.find(run => /办公/.test(run.text)).fontFamily, '宋体');
  assert.equal(toc[0].lineHeightRatio, 1.5);
  assert.equal(toc[0].spaceAfter, 8);
  assert.ok(!JSON.stringify(toc).includes('PAGEREF'));
  assert.ok(!JSON.stringify(toc).includes('TOC \\o'));
});

test('DOCX direct formatting overrides styles without applying paragraph-mark formatting to runs', async () => {
  const model = await updated(change => {
    change('word/document.xml', xml => xml.replace('<w:pStyle w:val="TOCHeading"/>', '<w:pStyle w:val="TOCHeading"/><w:rPr><w:sz w:val="10"/></w:rPr>')
      .replace('<w:t xml:space="preserve">目录</w:t>', '<w:rPr><w:b w:val="0"/><w:sz w:val="24"/><w:color w:val="123456"/></w:rPr><w:t xml:space="preserve">目录</w:t>'));
  });
  const heading = model.blocks[0];
  assert.equal(heading.runs[0].bold, false);
  assert.equal(heading.runs[0].fontSize, 16);
  assert.equal(heading.runs[0].color, '#123456');
  assert.equal(heading.paragraphMarkStyle.fontSize, 10 * 2 / 3);
});

test('DOCX keeps saved and explicit page breaks plus page-before and keep flags', async () => {
  const model = await parseOoxml(buildStyledDocx(), 'docx');
  const breaks = model.blocks.flatMap(block => block.runs || []).filter(run => run.type === 'pageBreak');
  assert.equal(model.hasSavedPageBreaks, true);
  assert.deepEqual(breaks.map(run => run.cached), [false, true]);
  assert.ok(model.blocks.find(block => text(block).includes('正文第一页')).pageBreakBefore);
  assert.ok(model.blocks.find(block => text(block).includes('正文第一页')).keepNext);
});

test('DOCX applies numbering indents, text and start overrides without rewriting document text', async () => {
  const model = await updated(change => {
    change('word/numbering.xml', xml => xml.replace('<w:abstractNumId w:val="3"/>', '<w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/></w:lvlOverride>'));
  });
  const list = model.blocks.find(block => block.numberingLabel);
  assert.equal(list.numberingLabel, '4.');
  assert.equal(list.numberingSuffix, 'tab');
  assert.equal(list.indentLeft, 48);
  assert.equal(list.firstLineIndent, -24);
  assert.equal(text(list), '编号列表项目');
});

test('DOCX merges and clears inherited tab stops and rejects cyclic style inheritance', async () => {
  const model = await updated(change => change('word/document.xml', xml => xml.replace('<w:pStyle w:val="TOC2"/>', '<w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="clear" w:pos="9000"/><w:tab w:val="right" w:leader="hyphen" w:pos="8400"/></w:tabs>')));
  assert.deepEqual(model.blocks.find(block => block.tocLevel === 2).tabs, [{ position: 560, align: 'right', leader: 'hyphen' }]);
  await assert.rejects(updated(change => change('word/styles.xml', xml => xml.replace('<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>', '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:basedOn w:val="TOC1"/>'))), /样式继承/);
});

test('DOCX section geometry, line grid and snap override remain local to the correct section', async () => {
  const model = await updated(change => change('word/document.xml', xml => xml
    .replace('<w:pStyle w:val="TOCHeading"/>', '<w:pStyle w:val="TOCHeading"/><w:sectPr><w:pgSz w:w="10000" w:h="15000"/><w:pgMar w:top="900" w:right="1000" w:bottom="900" w:left="1000"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>')
    .replace('<w:pStyle w:val="TOC1"/>', '<w:pStyle w:val="TOC1"/><w:snapToGrid w:val="0"/>')));
  assert.equal(model.sections.length, 2);
  assert.equal(model.page.width, 10000 / 15);
  assert.equal(model.page.lineGrid, 20.8);
  assert.equal(model.blocks[0].sectionIndex, 0);
  assert.equal(model.blocks[1].sectionIndex, 1);
  assert.equal(model.blocks[1].snapToGrid, false);
  assert.equal(model.sections[1].width, 11906 / 15);
});

test('DOCX inline images retain native dimensions and tables remain structural', async () => {
  const model = await parseOoxml(buildStyledDocx(), 'docx');
  const image = model.blocks.flatMap(block => block.runs || []).find(run => run.image).image;
  assert.equal(image.width, 100); assert.equal(image.height, 50); assert.equal(image.mime, 'image/png');
  assert.equal(image.anchor, undefined);
  const table = model.blocks.find(block => block.type === 'table');
  assert.equal(table.rows.length, 2); assert.equal(table.rows[0].length, 2);
  assert.deepEqual(table.columnWidths, [300, 300]);
});

test('DOCX reads page areas per section, their own media relationships and cached fields', async () => {
  const model = await parseOoxml(buildWordPageAreasDocx(), 'docx');
  const [first, last] = model.sections;
  assert.equal(model.evenAndOddHeaders, true);
  assert.equal(first.titlePage, true);
  assert.equal(last.titlePage, false);
  assert.deepEqual([first.pageNumberStart, last.pageNumberStart], [4, 8]);
  assert.deepEqual([first.headerDistance, first.footerDistance], [32, 24]);
  assert.equal(text(first.headers.first[0]), 'First header');
  assert.equal(text(last.headers.default[0]), 'Next odd header');
  assert.equal(last.headers.even, first.headers.even);
  const image = first.headers.even[0].runs.find(run => run.image).image;
  assert.deepEqual([image.width, image.height, image.mime], [32, 24, 'image/png']);
  assert.equal(last.footers.default, first.footers.default);
  assert.equal(text(last.footers.default[0].rows[0][0].blocks[0]), 'Cached page');
  assert.deepEqual(model.blocks.map(block => block.sectionIndex), [0, 0, 0, 1, 1]);
  assert.match(model.warnings.join('\n'), /页眉页脚中的域使用文件保存的结果/);
});

test('a damaged DOCX page area leaves the body and other areas readable', async () => {
  const parts = await readZip(buildWordPageAreasDocx());
  parts.set('word/even.xml', encoder.encode('<broken>'));
  const model = await parseOoxml(makeZip(Object.fromEntries(parts)), 'docx');
  assert.equal(text(model.blocks[0]), 'First body');
  assert.equal(text(model.sections[0].headers.first[0]), 'First header');
  assert.deepEqual(model.sections[1].headers.even, []);
  assert.match(model.warnings.join('\n'), /页眉页脚部件无效或缺失/);
});

test('DOCX print TOC suppresses implicit hyperlink color but preserves direct formatting and ordinary links', async () => {
  const model = await updated(change => {
    change('word/styles.xml', xml => xml.replace('</w:styles>', '<w:style w:type="character" w:styleId="Hyper"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>'));
    change('word/document.xml', xml => xml.replace('<w:t xml:space="preserve">一、办公文档测试</w:t>', '<w:rPr><w:rStyle w:val="Hyper"/></w:rPr><w:t xml:space="preserve">一、办公文档测试</w:t>')
      .replace('<w:t xml:space="preserve">1. 二级目录示例</w:t>', '<w:rPr><w:rStyle w:val="Hyper"/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve">1. 二级目录示例</w:t>')
      .replace('<w:t xml:space="preserve">这是自主生成的 Word 样式与分页开发样本。</w:t>', '<w:rPr><w:rStyle w:val="Hyper"/></w:rPr><w:t>普通超链接</w:t>'));
  });
  assert.equal(model.blocks[1].runs[0].color, '#000000');
  assert.equal(model.blocks[1].runs[0].underline, undefined);
  assert.equal(model.blocks[2].runs[0].color, '#FF0000');
  const ordinary = model.blocks.flatMap(block => block.runs || []).find(run => run.text === '普通超链接');
  assert.equal(ordinary.color, '#0563C1'); assert.equal(ordinary.underline, true);
});
