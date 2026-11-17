import test from 'node:test';
import assert from 'node:assert/strict';
import { readZip } from '../src/containers/zip.js';
import { parseXml } from '../src/xml.js';
import { parseOoxml } from '../src/parsers/ooxml.js';
import { readDrawingColor, cssColor } from '../src/parsers/ooxml-paint.js';
import { buildPptx, makeZip } from '../scripts/generate-fixtures.mjs';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const gradient = (attrs = 'ang="0" scaled="1"') => `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="E0153D"/></a:gs><a:gs pos="100000"><a:srgbClr val="2ED3E8"/></a:gs></a:gsLst><a:lin ${attrs}/></a:gradFill>`;
function shape(properties, text = '', width = 12192000, height = 76200) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Visual regression"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>${properties}</p:spPr>${text ? `<p:txBody><a:bodyPr/><a:lstStyle/>${text}</p:txBody>` : ''}</p:sp>`;
}
async function model(content, background = '') {
  const parts = await readZip(buildPptx());
  const source = `<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:p="${P}" xmlns:a="${A}"><p:cSld>${background}<p:spTree>${content}</p:spTree></p:cSld></p:sld>`;
  parts.set('ppt/slides/slide1.xml', new TextEncoder().encode(source));
  return parseOoxml(makeZip(Object.fromEntries(parts)), 'pptx');
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
