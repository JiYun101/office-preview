import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLegacy } from '../src/parsers/legacy.js';
import { readCompound } from '../src/containers/cfb.js';
import { makeCompound, buildLegacyPpt } from '../scripts/generate-fixtures.mjs';
import { buildStyledPpt } from '../scripts/legacy-style-fixtures.mjs';

const paragraphText = element => (element.runs || []).map(run => run.text || '').join('');
function streams(bytes) {
  const compound = readCompound(bytes);
  return Object.fromEntries(compound.names.map(name => [name, compound.get(name).slice()]));
}
function findRecords(bytes, type) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), matches = [];
  const visit = (start, end) => {
    for (let offset = start; offset + 8 <= end;) {
      const flags = view.getUint16(offset, true), recordType = view.getUint16(offset + 2, true), length = view.getUint32(offset + 4, true);
      const next = offset + 8 + length;
      if (next > end) break;
      if (recordType === type) matches.push({ offset, data: offset + 8, length, instance: flags >>> 4 });
      if ((flags & 15) === 15) visit(offset + 8, next);
      offset = next;
    }
  };
  visit(0, bytes.length); return matches;
}

test('PPT restores native canvas size, slide colors and inherited master geometry', () => {
  const model = parseLegacy(buildStyledPpt(), 'ppt');
  assert.equal(model.width, 960);
  assert.equal(model.height, 540);
  assert.equal(model.slides.length, 2);
  assert.equal(model.slides[0].background, '#f3f7fc');
  assert.equal(model.slides[1].background, '#fff3e9');
  assert.ok(model.slides[0].elements.some(element => element.type === 'shape' && element.y === 500 && element.height === 40), 'master decoration is inherited at its real coordinates');
  assert.ok(!model.slides[1].elements.some(element => element.type === 'shape' && element.y === 500 && element.height === 40), 'slide master objects are controlled by SlideFlags');
});

test('PPT keeps per-shape fills, positions, font family, size, color and paragraph alignment', () => {
  const model = parseLegacy(buildStyledPpt(), 'ppt');
  const first = model.slides[0];
  const banner = first.elements.find(element => element.type === 'shape' && element.fill === '#184c4c' && element.width === 840);
  assert.deepEqual([banner.x, banner.y, banner.width, banner.height], [60, 52, 840, 100]);
  const title = first.elements.find(element => element.type === 'text' && paragraphText(element).includes('旧版 PPT 图片与颜色'));
  assert.equal(title.align, 'center');
  const titleRun = title.runs.find(run => run.text.includes('旧版 PPT'));
  assert.equal(titleRun.fontFamily, '微软雅黑');
  assert.equal(titleRun.bold, true);
  assert.equal(titleRun.color, '#ffffff');
  assert.ok(Math.abs(titleRun.fontSize - 26 * 96 / 72) < 1e-8);
  const subtitle = first.elements.find(element => element.type === 'text' && paragraphText(element).includes('按原始位置显示'));
  assert.equal(subtitle.paragraphs.length, 2);
  assert.ok(subtitle.x > title.x + 400, 'separate text boxes must retain separate positions');
});

test('PPT reconstructs a marked OfficeArt table from native child anchors and keeps its separate borders', () => {
  const { elements } = parseLegacy(buildStyledPpt(), 'ppt').slides[0];
  const table = elements.find(element => element.type === 'table');
  assert.deepEqual([table.x, table.y, table.width, table.height], [520, 340, 360, 136]);
  assert.deepEqual(table.columns, [80, 130, 150]);
  assert.deepEqual(table.rows.map(row => row.height), [40, 52, 44]);
  const heading = table.rows[0].cells[0], left = table.rows[1].cells[0], green = table.rows[1].cells[1];
  assert.deepEqual([heading.column, heading.colSpan, heading.rowSpan], [0, 2, 1]);
  assert.equal(heading.paragraphs[0].runs[0].text, '合并表头');
  assert.deepEqual([left.column, left.colSpan, left.rowSpan], [0, 1, 2]);
  assert.equal(green.fill, '#22aa44');
  assert.deepEqual(table.rows[1].cells[2].paragraphs, []);
  assert.ok(elements.some(element => element.shape === 'line' && element.stroke === '#331155' && element.x === 600 && element.y === 380));
  assert.ok(!elements.some(element => element.type === 'text' && paragraphText(element) === '合并表头'));
});

test('PPT keeps its grouped visual fallback when table child anchors overlap', () => {
  const parts = streams(buildStyledPpt());
  const ppt = parts['PowerPoint Document'];
  const anchor = findRecords(ppt, 0xf00f)[1];
  new DataView(ppt.buffer, ppt.byteOffset, ppt.byteLength).setInt32(anchor.data, 0, true);
  const model = parseLegacy(makeCompound(parts), 'ppt');
  assert.ok(!model.slides[0].elements.some(element => element.type === 'table'));
  assert.ok(model.slides[0].elements.some(element => element.type === 'text' && paragraphText(element) === '合并表头'));
  assert.ok(model.warnings.some(warning => warning.includes('表格的网格无法恢复')));
});

test('PPT resolves pib through BSE foDelay into the Pictures stream and retains PNG bytes', () => {
  const model = parseLegacy(buildStyledPpt(), 'ppt');
  const image = model.slides[0].elements.find(element => element.type === 'image');
  assert.equal(image.mime, 'image/png');
  assert.deepEqual([...image.bytes.subarray(0, 8)], [137, 80, 78, 71, 13, 10, 26, 10]);
  assert.deepEqual([image.x, image.y, image.width, image.height], [70, 200, 400, 240]);
});

test('PPT resolves embedded BLIP records inside BSE entries without a Pictures stream', () => {
  const bytes = buildStyledPpt({ embeddedBlip: true });
  const compound = readCompound(bytes);
  // The fixture option must actually exercise embedded storage.
  assert.equal(compound.get('Pictures'), undefined);
  const model = parseLegacy(bytes, 'ppt');
  assert.ok(model.slides[0].elements.some(element => element.type === 'image' && element.mime === 'image/png'));
});

test('PPT rejects broken image payloads instead of passing unrelated bytes to the browser', () => {
  const parts = streams(buildStyledPpt());
  parts.Pictures[25] = 0;
  assert.throws(() => parseLegacy(makeCompound(parts), 'ppt'), /PNG.*标识/);
});

test('PPT validates BSE delayed offsets and anchor record boundaries', () => {
  const parts = streams(buildStyledPpt());
  const ppt = parts['PowerPoint Document'], view = new DataView(ppt.buffer, ppt.byteOffset, ppt.byteLength);
  const bse = findRecords(ppt, 0xf007)[0];
  view.setUint32(bse.data + 28, parts.Pictures.length + 100, true);
  assert.throws(() => parseLegacy(makeCompound(parts), 'ppt'), /越界/);
  const anchorParts = streams(buildStyledPpt());
  const anchorBytes = anchorParts['PowerPoint Document'];
  const anchor = findRecords(anchorBytes, 0xf010)[0];
  new DataView(anchorBytes.buffer, anchorBytes.byteOffset, anchorBytes.byteLength).setUint32(anchor.offset + 4, 12, true);
  assert.throws(() => parseLegacy(makeCompound(anchorParts), 'ppt'), /长度|容器|记录/);
});

test('PPT checks JPEG dimensions before handing an image to the native decoder', () => {
  const parts = streams(buildStyledPpt());
  const jpeg = new Uint8Array(25 + 13);
  const view = new DataView(jpeg.buffer);
  view.setUint16(0, 0x46a << 4, true); view.setUint16(2, 0xf01d, true); view.setUint32(4, jpeg.length - 8, true);
  jpeg[24] = 255;
  jpeg.set([255,216,255,192,0,9,8,255,255,255,255,1,0],25);
  parts.Pictures = jpeg;
  assert.throws(() => parseLegacy(makeCompound(parts), 'ppt'), /JPEG.*像素/);
});

test('PPT fallback is explicit for old text-only structures and does not restore deleted pages', () => {
  const model = parseLegacy(buildLegacyPpt(), 'ppt');
  assert.equal(model.slides.length, 2);
  assert.match(JSON.stringify(model.slides[0].elements), /中文文本/);
  assert.doesNotMatch(JSON.stringify(model.slides), /ORPHAN DELETED/);
  assert.ok(model.warnings.some(warning => warning.includes('阅读模式')));
});
