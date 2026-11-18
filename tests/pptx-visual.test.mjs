import test from 'node:test';
import assert from 'node:assert/strict';
import { readZip } from '../src/containers/zip.js';
import { parseXml } from '../src/xml.js';
import { parseOoxml } from '../src/parsers/ooxml.js';
import { readDrawingColor, cssColor } from '../src/parsers/ooxml-paint.js';
import { buildPptx, makeZip } from '../scripts/generate-fixtures.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const gradient = (attrs = 'ang="0" scaled="1"') => `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="E0153D"/></a:gs><a:gs pos="100000"><a:srgbClr val="2ED3E8"/></a:gs></a:gsLst><a:lin ${attrs}/></a:gradFill>`;
function shape(properties, text = '', width = 12192000, height = 76200) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Visual regression"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>${properties}</p:spPr>${text ? `<p:txBody><a:bodyPr/><a:lstStyle/>${text}</p:txBody>` : ''}</p:sp>`;
}
async function model(content, background = '', slideParts = {}) {
  const parts = await readZip(buildPptx());
  const source = `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:cSld>${background}<p:spTree>${content}</p:spTree></p:cSld></p:sld>`;
  parts.set('ppt/slides/slide1.xml', new TextEncoder().encode(source));
  return parseOoxml(makeZip({ ...Object.fromEntries(parts), ...slideParts }), 'pptx');
}

test('PPTX gradient-only shapes retain both colors, positions and a scaled Office angle', async () => {
  const parsed = await model(shape(`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${gradient()}`));
  const strip = parsed.slides[0].elements[0];
  assert.equal(strip.type, 'shape');
  assert.equal(strip.width, 1280);
  assert.equal(strip.height, 8);
  assert.deepEqual(strip.fillPaint, { type: 'linear', angle: 0, scaled: true, stops: [{ offset: 0, color: '#E0153D' }, { offset: 1, color: '#2ED3E8' }] });
  assert.ok(!parsed.warnings.some(warning => warning.includes('渐变') && warning.includes('尚未')));
});

test('PPTX linear gradients preserve unscaled diagonal directions and reject bad stop positions', async () => {
  const parsed = await model(shape(gradient('ang="2700000" scaled="0"')));
  assert.equal(parsed.slides[0].elements[0].fillPaint.angle, 45);
  assert.equal(parsed.slides[0].elements[0].fillPaint.scaled, false);
  await assert.rejects(model(shape(gradient().replace('pos="100000"', 'pos="100001"'))), /停靠点位置/);
});

test('PPTX run character spacing and baseline are converted from real DrawingML units', async () => {
  const text = '<a:p><a:pPr><a:defRPr sz="1500" spc="150"/></a:pPr><a:r><a:rPr spc="200"/><a:t>ACE</a:t></a:r><a:r><a:rPr spc="-75"/><a:t>tight</a:t></a:r><a:r><a:rPr spc="0" baseline="30000"/><a:t>zero</a:t></a:r></a:p>';
  const parsed = await model(shape('<a:solidFill><a:srgbClr val="101010"/></a:solidFill>', text));
  const runs = parsed.slides[0].elements.find(element => element.type === 'text').runs;
  assert.equal(runs[0].letterSpacing, 200 / 75);
  assert.equal(runs[1].letterSpacing, -1);
  assert.equal(runs[2].letterSpacing, 0);
  assert.equal(runs[2].baselineShift, 6);
  assert.ok(runs.every(run => run.fontSize === 20));
});

test('PPTX paragraph spacing retains point distances and percentage ratios independently', async () => {
  const text = '<a:p><a:pPr><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="1200"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="1500"/><a:t>First</a:t></a:r></a:p><a:p><a:pPr><a:spcBef><a:spcPct val="100000"/></a:spcBef><a:spcAft><a:spcPct val="50000"/></a:spcAft></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>';
  const parsed = await model(shape('', text));
  const paragraphs = parsed.slides[0].elements.find(element => element.type === 'text').paragraphs;
  assert.equal(paragraphs[0].spaceBefore, 16);
  assert.equal(paragraphs[0].spaceAfter, 8);
  assert.equal(paragraphs[0].lineHeight, 1.5);
  assert.equal(paragraphs[1].spaceBeforeRatio, 1);
  assert.equal(paragraphs[1].spaceAfterRatio, 0.5);
});

test('PPTX rounded corners and outlines follow explicit adjustment and EMU width', async () => {
  const properties = '<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 6000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="111C2D"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="22354F"/></a:solidFill></a:ln>';
  const parsed = await model(shape(properties, '', 2540000, 1092200));
  const card = parsed.slides[0].elements[0];
  assert.equal(card.cornerRadius, 6.88);
  assert.equal(card.strokeWidth, 4 / 3);
  assert.equal(card.stroke, '#22354F');
});

test('DrawingML color modifiers apply ordered luminance, saturation and alpha changes', () => {
  const color = xml => readDrawingColor(parseXml(`<a:solidFill xmlns:a="${A}">${xml}</a:solidFill>`).documentElement, { accent1: '#0000FF' });
  assert.deepEqual(color('<a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr>'), { color: '#000080' });
  assert.deepEqual(color('<a:schemeClr val="accent1"><a:satMod val="0"/></a:schemeClr>'), { color: '#808080' });
  assert.deepEqual(color('<a:srgbClr val="FF0000"><a:alpha val="75000"/><a:alphaMod val="50000"/></a:srgbClr>'), { color: '#FF0000', opacity: 0.375 });
  assert.equal(cssColor(color('<a:srgbClr val="000000"><a:tint val="50000"/><a:alpha val="50000"/></a:srgbClr>')), 'rgba(128,128,128,0.5)');
});

test('PPTX slide background gradients use the same paint model as shape gradients', async () => {
  const parsed = await model('', `<p:bg><p:bgPr>${gradient('ang="5400000" scaled="0"')}</p:bgPr></p:bg>`);
  assert.equal(parsed.slides[0].backgroundPaint.type, 'linear');
  assert.equal(parsed.slides[0].backgroundPaint.angle, 90);
  assert.equal(parsed.slides[0].backgroundPaint.stops.length, 2);
});

test('PPTX picture crop and object rotation use DrawingML units and share the outline center with text', async () => {
  const bounds = '<a:off x="95250" y="190500"/><a:ext cx="1905000" cy="952500"/>';
  const picture = `<p:pic><p:blipFill><a:blip r:embed="picture"/><a:srcRect l="25000" t="-10%" b="10000"/></p:blipFill><p:spPr><a:xfrm rot="5400000" flipH="true" flipV="1">${bounds}</a:xfrm></p:spPr></p:pic>`;
  const label = shape('<a:solidFill><a:srgbClr val="111C2D"/></a:solidFill>', '<a:p><a:r><a:t>Label</a:t></a:r></a:p>', 1905000, 952500)
    .replace('<a:xfrm>', '<a:xfrm rot="1800000" flipH="1">').replace('<a:bodyPr/>', '<a:bodyPr lIns="190500" tIns="95250"/>');
  const gif = Buffer.from('R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', 'base64');
  const parsed = await model(picture + label, '', {
    'ppt/slides/_rels/slide1.xml.rels': `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="picture" Type="${R}/image" Target="../media/image.gif"/></Relationships>`,
    'ppt/media/image.gif': gif,
  });
  const [image, outline, text] = parsed.slides[0].elements;
  assert.deepEqual(image.crop, { left: 0.25, top: -0.1, right: 0, bottom: 0.1 });
  assert.equal(image.rotation, 90);
  assert.equal(image.flipH, true);
  assert.equal(image.flipV, true);
  assert.deepEqual(Buffer.from(image.bytes), gif);
  assert.equal(outline.flipH, true);
  assert.equal(text.rotation, outline.rotation);
  assert.deepEqual(text.transformOrigin, { x: 80, y: 40 });
  assert.ok(!parsed.warnings.some(warning => /裁剪|旋转和翻转对象/.test(warning)));
});

test('PPTX static table keeps its physical grid, cell content, spanning merge and outside borders', async () => {
  const border = (side, color) => `<a:ln${side} w="19050"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:prstDash val="dash"/></a:ln${side}>`;
  const cell = (attrs, name, properties = '') => `<a:tc ${attrs}><a:txBody><a:bodyPr/><a:p><a:r><a:t>${name}</a:t></a:r></a:p></a:txBody><a:tcPr marL="19050" anchor="ctr">${properties}</a:tcPr></a:tc>`;
  const table = `<p:graphicFrame><p:xfrm><a:off x="95250" y="190500"/><a:ext cx="2857500" cy="762000"/></p:xfrm><a:graphic><a:graphicData uri="${A}/table"><a:tbl><a:tblGrid><a:gridCol w="952500"/><a:gridCol w="1428750"/><a:gridCol w="476250"/></a:tblGrid><a:tr h="285750">${cell('gridSpan="2" rowSpan="2"', 'Merged', '<a:solidFill><a:srgbClr val="DEEFFA"/></a:solidFill>')}${cell('rowSpan="2" hMerge="1"', 'Covered', border('R', '123456'))}${cell('', 'A')}</a:tr><a:tr h="476250">${cell('gridSpan="2" vMerge="1"', 'Covered', border('B', '654321'))}${cell('hMerge="1" vMerge="1"', 'Covered')}${cell('', 'B')}</a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>`;
  const parsed = await model(table);
  const [item] = parsed.slides[0].elements;
  assert.equal(item.type, 'table');
  assert.deepEqual([item.x, item.y, item.width, item.height], [10, 20, 300, 80]);
  assert.deepEqual(item.columns, [100, 150, 50]);
  assert.deepEqual(item.rows.map(row => row.height), [30, 50]);
  assert.deepEqual(item.rows.flatMap(row => row.cells.map(cell => cell.paragraphs[0].runs[0].text)), ['Merged', 'A', 'B']);
  const [merged] = item.rows[0].cells;
  assert.deepEqual([merged.column, merged.colSpan, merged.rowSpan], [0, 2, 2]);
  assert.deepEqual(merged.borders, { right: '2px dashed #123456', bottom: '2px dashed #654321' });
  assert.equal(merged.fill, '#DEEFFA');
  assert.equal(merged.padding[3], 2);
  assert.equal(merged.verticalAlign, 'middle');
});
