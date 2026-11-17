import test from 'node:test';
import assert from 'node:assert/strict';
import { parseXml } from '../src/xml.js';
import { children, elements } from '../src/parsers/ooxml-common.js';
import { parseWordTable } from '../src/parsers/ooxml-word-table.js';
import { buildWordTableFixtureEntries } from '../scripts/word-table-fixture.mjs';

const ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const cell = (properties = '', text = 'x') => `<w:tc><w:tcPr>${properties}</w:tcPr><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:tc>`;
const row = (cells, properties = '') => `<w:tr><w:trPr>${properties}</w:trPr>${cells}</w:tr>`;
function parse(body, properties = '', grid = [900, 1800, 2700, 3600]) {
  const warnings = [];
  const document = parseXml(`<w:tbl xmlns:w="${ns}"><w:tblPr>${properties}</w:tblPr><w:tblGrid>${grid.map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${body}</w:tbl>`);
  const blocks = node => children(node, 'p').map(paragraph => ({ type: 'paragraph', runs: [{ text: elements(paragraph, 't').map(node => node.textContent).join('') }] }));
  return { table: parseWordTable(document.documentElement, { blocks, sectionIndex: 2, warnings }), warnings };
}

test('DOCX tables retain unequal grid widths, table sizing, alignment and explicit row heights', () => {
  const { table } = parse(row(cell().repeat(4), '<w:trHeight w:val="600" w:hRule="exact"/>') + row(cell().repeat(4), '<w:trHeight w:val="900" w:hRule="atLeast"/>'), '<w:tblW w:w="9000" w:type="dxa"/><w:jc w:val="center"/><w:tblInd w:w="300" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellSpacing w:w="30" w:type="dxa"/>');
  assert.deepEqual(table.columnWidths, [60, 120, 180, 240]);
  assert.equal(table.width, 600); assert.equal(table.align, 'center'); assert.equal(table.indent, 20); assert.equal(table.layout, 'fixed'); assert.equal(table.cellSpacing, 2);
  assert.equal(table.sectionIndex, 2);
  assert.deepEqual(table.rowProperties, [{ height: 40, heightRule: 'exact' }, { height: 60, heightRule: 'atLeast' }]);
  assert.deepEqual(table.rows.map(row => row.map(cell => cell.column)), [[0, 1, 2, 3], [0, 1, 2, 3]]);
});

test('DOCX table percentage widths and auto cell widths preserve their distinct sizing semantics', () => {
  for (const amount of ['5000', '100%']) {
    const { table } = parse(row(cell('<w:tcW w:w="0" w:type="auto"/>') + cell('<w:tcW w:w="2500" w:type="pct"/>') + cell('<w:tcW w:w="1800" w:type="dxa"/>')), `<w:tblW w:w="${amount}" w:type="pct"/><w:tblLayout w:type="autofit"/>`);
    assert.equal(table.widthPercent, 100); assert.equal(table.width, undefined); assert.equal(table.layout, 'autofit');
    assert.equal(table.rows[0][0].width, undefined); assert.equal(table.rows[0][1].widthPercent, 50); assert.equal(table.rows[0][2].width, 120);
  }
  const { table } = parse(row(cell()), '<w:tblInd w:w="-150" w:type="dxa"/>');
  assert.equal(table.indent, -10);
});

test('DOCX cell formatting keeps individual border colors, none overrides, margins, shading and vertical alignment', () => {
  const { table } = parse(row(cell('<w:tcBorders><w:left w:val="single" w:sz="6" w:color="937264"/><w:top w:val="dashed" w:sz="12" w:color="112233"/><w:bottom w:val="nil"/></w:tcBorders><w:tcMar><w:top w:type="dxa" w:w="60"/><w:start w:type="dxa" w:w="45"/></w:tcMar><w:vAlign w:val="center"/><w:shd w:val="clear" w:fill="DDEEFF"/><w:noWrap/>')), '<w:tblBorders><w:top w:val="single" w:sz="6" w:color="000000"/><w:insideH w:val="dotted" w:sz="3" w:color="auto"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>');
  assert.deepEqual(table.borders.top, { style: 'single', width: 1, color: '#000000' });
  assert.deepEqual(table.borders.insideH, { style: 'dotted', width: .5, color: '#000000' });
  assert.deepEqual(table.cellMargins, { top: 0, left: 7.2, right: 7.2 });
  const formatted = table.rows[0][0];
  assert.deepEqual(formatted.borders.left, { style: 'single', width: 1, color: '#937264' });
  assert.deepEqual(formatted.borders.top, { style: 'dashed', width: 2, color: '#112233' });
  assert.deepEqual(formatted.borders.bottom, { style: 'none', width: 0, color: '#000000' });
  assert.deepEqual(formatted.margins, { top: 4, left: 3 });
  assert.equal(formatted.verticalAlign, 'center'); assert.equal(formatted.shading, '#DDEEFF'); assert.equal(formatted.noWrap, true);
  const cleared = parse(row(cell('<w:shd w:val="nil" w:fill="112233"/><w:noWrap w:val="0"/>'))).table.rows[0][0];
  assert.equal(cleared.shading, 'transparent'); assert.equal(cleared.noWrap, false);
});

test('DOCX gridSpan and vMerge combine only adjacent matching physical columns and keep the last outer border', () => {
  const merged = '<w:gridSpan w:val="2"/>';
  const { table, warnings } = parse(
    row(cell(merged + '<w:vMerge w:val="restart"/><w:tcBorders><w:top w:val="single" w:sz="6"/><w:bottom w:val="none"/></w:tcBorders>', 'anchor') + cell() + cell()) +
    row(cell(merged + '<w:vMerge/><w:tcBorders><w:bottom w:val="single" w:sz="12" w:color="123456"/></w:tcBorders>', 'ignored') + cell() + cell()) +
    row(cell(merged, 'separate') + cell() + cell()) +
    row(cell(merged + '<w:vMerge/>', 'orphan') + cell() + cell())
  );
  const anchor = table.rows[0][0];
  assert.equal(anchor.colSpan, 2); assert.equal(anchor.rowSpan, 2); assert.equal(anchor.column, 0);
  assert.equal(anchor.blocks[0].runs[0].text, 'anchor');
  assert.deepEqual(anchor.borders.bottom, { style: 'single', width: 2, color: '#123456' });
  assert.deepEqual(table.rows[1].map(cell => cell.column), [2, 3]);
  assert.equal(table.rows[3][0].rowSpan, undefined); assert.equal(table.rows[3][0].blocks[0].runs[0].text, 'orphan');
  assert.equal(warnings.length, 1);
});

test('DOCX omitted grid columns and mismatched vertical spans do not attach to an unrelated anchor', () => {
  const { table, warnings } = parse(
    row(cell('<w:vMerge w:val="restart"/>') + cell('<w:gridSpan w:val="2"/>'), '<w:gridBefore w:val="1"/>') +
    row(cell('<w:gridSpan w:val="2"/><w:vMerge/>') + cell(), '<w:gridBefore w:val="1"/>') +
    row(cell(), '<w:gridBefore w:val="2"/><w:gridAfter w:val="1"/>')
  );
  assert.equal(table.rows[0][0].column, 1); assert.equal(table.rows[0][0].rowSpan, 1);
  assert.equal(table.rows[1][0].column, 1); assert.equal(table.rows[1][0].colSpan, 2); assert.equal(table.rows[1][0].rowSpan, undefined);
  assert.deepEqual(table.rowProperties[2], { gridBefore: 2, gridAfter: 1 });
  assert.equal(warnings.length, 1);
});

test('DOCX table dimensions and spans reject negative or excessive geometry', () => {
  assert.throws(() => parse(row(cell('<w:gridSpan w:val="2048"/>'))), /表格/);
  assert.throws(() => parse(row(cell('<w:gridSpan w:val="1.5"/>'))), /表格/);
  assert.throws(() => parse(row(cell(), '<w:trHeight w:val="-1"/>')), /行高/);
  assert.throws(() => parse(row(cell()), '<w:tblW w:w="-10" w:type="dxa"/>'), /尺寸/);
  assert.throws(() => parse(row(cell()), '', [-15]), /列宽/);
});

test('the self-authored browser fixture contains real unequal widths, merged cells and row geometry', () => {
  const document = parseXml(buildWordTableFixtureEntries()['word/document.xml']);
  const table = parseWordTable(elements(document, 'tbl')[0], { blocks: () => [], warnings: [] });
  assert.equal(table.width, 600); assert.deepEqual(table.columnWidths, [60, 120, 180, 240]);
  assert.deepEqual(table.rowProperties, [{ height: 40, heightRule: 'exact' }, { height: 60, heightRule: 'atLeast' }, { height: 30, heightRule: 'atLeast' }]);
  assert.equal(table.rows[0][0].colSpan, 2); assert.equal(table.rows[0][1].rowSpan, 2);
  assert.equal(table.rows[0][0].shading, '#DDEEFF');
});
