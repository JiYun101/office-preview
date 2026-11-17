import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLegacy } from '../src/parsers/legacy.js';
import { enrichPptRoundtrip } from '../src/parsers/legacy-ppt-art.js';
import { makeCompound, makeZip } from '../scripts/generate-fixtures.mjs';

const u16 = value => { const bytes = Buffer.alloc(2); bytes.writeUInt16LE(value & 65535); return bytes; };
const u32 = value => { const bytes = Buffer.alloc(4); bytes.writeUInt32LE(value >>> 0); return bytes; };
const record = (type, bytes, instance = 0, version = 0) => Buffer.concat([u16(instance * 16 + version), u16(type), u32(bytes.length), bytes]);
const container = (type, items, instance = 0) => record(type, Buffer.concat(items), instance, 15);

function options(properties) {
  const entries = [...properties].sort((a, b) => a[0] - b[0]);
  const head = [], data = [];
  for (const [id, value] of entries) {
    const complex = value instanceof Uint8Array;
    head.push(u16(id | (complex ? 0x8000 : 0)), u32(complex ? value.length : value));
    if (complex) data.push(Buffer.from(value));
  }
  return record(0xf00b, Buffer.concat([...head, ...data]), entries.length, 3);
}

function visualPpt({ shape = 1, properties = [], text = '', paragraphBefore = 0, paragraphAfter = 0 } = {}) {
  const nativeText = text ? [
    record(0x0f9f, u32(4)),
    record(0x0fa0, Buffer.from(text, 'utf16le')),
    record(0x0fa1, Buffer.concat([
      u32(text.length + 1), u16(0), u32(0x2000 | 0x4000), u16(paragraphBefore), u16(paragraphAfter),
      u32(text.length + 1), u32(0x20000), u16(12),
    ])),
  ] : [];
  const drawingShape = container(0xf004, [
    record(0xf00a, Buffer.concat([u32(1025), u32(0x0a00)]), shape, 2),
    options([[0x01bf, 0x00100010], [0x01ff, 0x00080008], ...properties]),
    record(0xf010, Buffer.concat([u32(120), u32(240), u32(1440), u32(720)])),
    ...(text ? [record(0xf00d, Buffer.concat(nativeText))] : []),
  ]);
  const documentAtom = Buffer.alloc(40); documentAtom.writeUInt32LE(5760); documentAtom.writeUInt32LE(3240, 4);
  const slidePersist = Buffer.alloc(20); slidePersist.writeUInt32LE(2); slidePersist.writeUInt32LE(256, 12);
  const document = container(0x03e8, [record(0x03e9, documentAtom), container(0x0ff0, [record(0x03f3, slidePersist)])]);
  const slide = container(0x03ee, [record(0x03ef, Buffer.alloc(24)), container(0x040c, [container(0xf002, [drawingShape])])]);
  const persistOffset = document.length + slide.length;
  const persist = record(0x1772, Buffer.concat([u32((2 << 20) | 1), u32(0), u32(document.length)]));
  const editOffset = persistOffset + persist.length;
  const editData = Buffer.alloc(28); editData.writeUInt32LE(persistOffset, 12); editData.writeUInt32LE(1, 16);
  const edit = record(0x0ff5, editData);
  const userData = Buffer.alloc(20); userData.writeUInt32LE(0xe391c05f, 4); userData.writeUInt32LE(editOffset, 8);
  return makeCompound({ 'PowerPoint Document': Buffer.concat([document, slide, persist, edit]), 'Current User': record(0x0ff6, userData) });
}

const firstShape = model => model.slides[0].elements.find(element => element.type === 'shape');
const firstText = model => model.slides[0].elements.find(element => element.type === 'text');
const metro = ({ text = 'Visual title', size = 1150, spacing = 200 } = {}) => makeZip({
  'drs/shapexml.xml': `<?xml version="1.0"?><p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:spPr><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr" lIns="0" rIns="0" tIns="0" bIns="0"/><a:p><a:pPr algn="r"><a:lnSpc><a:spcPct val="130000"/></a:lnSpc><a:spcBef><a:spcPts val="300"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r><a:rPr sz="${size}" spc="${spacing}" b="1"><a:solidFill><a:srgbClr val="2ED3E8"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`,
});

test('PPT native linear fill preserves angle, endpoint colors and shape scaling', () => {
  const model = parseLegacy(visualPpt({ properties: [[0x0180, 7], [0x0181, 0x003d15e0], [0x0183, 0x00e8d32e], [0x018b, 90 * 65536], [0x018c, 100]] }), 'ppt');
  assert.deepEqual(firstShape(model).fillPaint, { type: 'linear', angle: 0, scaled: true, stops: [{ offset: 0, color: '#e0153d', opacity: 1 }, { offset: 1, color: '#2ed3e8', opacity: 1 }] });
});

test('PPT shade arrays retain intermediate colors and opacity, including reflected focus', () => {
  const shades = Buffer.concat([u16(3), u16(3), u16(8), u32(0x000000ff), u32(0), u32(0x0000ff00), u32(32768), u32(0x00ff0000), u32(65536)]);
  const model = parseLegacy(visualPpt({ properties: [[0x0180, 4], [0x0182, 32768], [0x0184, 65536], [0x018c, 100], [0x0197, shades]] }), 'ppt');
  const paint = firstShape(model).fillPaint;
  assert.equal(paint.scaled, false);
  assert.deepEqual(paint.stops, [{ offset: 0, color: '#ff0000', opacity: 0.5 }, { offset: 0.5, color: '#00ff00', opacity: 0.75 }, { offset: 1, color: '#0000ff', opacity: 1 }]);
  const reflected = parseLegacy(visualPpt({ properties: [[0x0180, 7], [0x0181, 0x000000ff], [0x0183, 0x00ff0000], [0x018c, 50]] }), 'ppt');
  assert.deepEqual(firstShape(reflected).fillPaint.stops.map(stop => [stop.offset, stop.color]), [[0, '#ff0000'], [0.5, '#0000ff'], [1, '#ff0000']]);
});

test('PPT native round-rectangle adjustments and stroke widths remain in file units', () => {
  const model = parseLegacy(visualPpt({ shape: 2, properties: [[0x0147, 1296], [0x01cb, 12700], [0x0182, 32768]] }), 'ppt');
  const shape = firstShape(model);
  assert.equal(shape.width, 200);
  assert.equal(shape.height, 100);
  assert.equal(shape.cornerRadius, 6);
  assert.equal(shape.strokeWidth, 4 / 3);
  assert.equal(shape.fillOpacity, 0.5);
});

test('PPT paragraph spacing reads exact master-unit and line-relative distances', () => {
  const model = parseLegacy(visualPpt({ shape: 202, text: 'Visual title', paragraphBefore: -24, paragraphAfter: 200 }), 'ppt');
  const paragraph = firstText(model).paragraphs[0];
  assert.equal(paragraph.spaceBefore, 4);
  assert.equal(paragraph.spaceAfter, 38.4);
});

test('PPT roundtrip metadata restores fractional sizes and character tracking through the async reader', async () => {
  const model = parseLegacy(visualPpt({ shape: 202, text: 'Visual title', properties: [[0x03a9, metro()]] }), 'ppt');
  assert.equal(firstText(model).runs[0].fontSize, 16);
  await enrichPptRoundtrip(model);
  const element = firstText(model), run = element.runs[0];
  assert.equal(run.text, 'Visual title');
  assert.equal(run.fontSize, 1150 / 75);
  assert.equal(run.letterSpacing, 200 / 75);
  assert.equal(run.fontFamily, 'Arial');
  assert.equal(run.color, '#2ED3E8');
  assert.equal(element.verticalAlign, 'middle');
  assert.equal(element.align, 'right');
  assert.equal(element.x, 40);
  assert.equal(element.y, 20);
  assert.equal(element.paragraphs[0].lineHeight, 1.3);
  assert.equal(element.paragraphs[0].spaceBefore, 4);
  assert.equal(element.paragraphs[0].spaceAfter, 8);
});

test('PPT stale compatibility text cannot replace the current binary body', async () => {
  const model = parseLegacy(visualPpt({ shape: 202, text: 'Current body', properties: [[0x03a9, metro({ text: 'Previous body' })]] }), 'ppt');
  await enrichPptRoundtrip(model);
  assert.equal(firstText(model).runs[0].text, 'Current body');
  assert.equal(firstText(model).runs[0].fontSize, 16);
  assert.ok(model.warnings.some(message => message.includes('当前正文')));
});

test('PPT malformed optional compatibility ZIP falls back while invalid shade arrays fail clearly', async () => {
  const model = parseLegacy(visualPpt({ shape: 202, text: 'Visual title', properties: [[0x03a9, Buffer.from('not a ZIP')]] }), 'ppt');
  await enrichPptRoundtrip(model);
  assert.equal(firstText(model).runs[0].text, 'Visual title');
  assert.ok(model.warnings.some(message => message.includes('兼容属性无法读取')));
  const shades = Buffer.concat([u16(2), u16(2), u16(8), u32(0), u32(65536), u32(0xffffff), u32(0)]);
  assert.throws(() => parseLegacy(visualPpt({ properties: [[0x0180, 7], [0x0197, shades]] }), 'ppt'), /渐变颜色位置/);
});
