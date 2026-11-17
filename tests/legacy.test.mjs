import test from 'node:test';
import assert from 'node:assert/strict';
import { parseLegacy } from '../src/parsers/legacy.js';
import { buildLegacyDoc, buildLegacyXls, buildLegacyPpt, makeCompound } from '../scripts/generate-fixtures.mjs';

const paragraphText = (model) => model.blocks.map((block) => (block.runs || []).map((run) => run.text).join('')).join('\n');
const slideText = (slide) => slide.elements.map((element) => (element.runs || []).map((run) => run.text).join('')).join('\n');

test('DOC follows FIB/CLX piece table to read UTF-16 Chinese paragraphs', () => {
  const model = parseLegacy(buildLegacyDoc(), 'doc', 'sample.doc');
  assert.equal(model.kind, 'document');
  assert.match(paragraphText(model), /自主研发 Office 预览套件/);
  assert.match(paragraphText(model), /这是旧版 DOC 中文正文。/);
  assert.ok(model.warnings.length > 0, 'basic text preview must expose fidelity limitations');
});

test('DOC joins Unicode and compressed ANSI pieces in logical character order', () => {
  const model = parseLegacy(buildLegacyDoc({ mixed: true }), 'doc');
  assert.match(paragraphText(model), /中文正文。/);
  assert.match(paragraphText(model), /ASCII tail/);
  assert.ok(paragraphText(model).indexOf('ASCII tail') > paragraphText(model).indexOf('中文正文'));
});

for (const continuation of [false, true]) {
  test(`XLS BIFF8 reads SST${continuation ? ' with CONTINUE' : ''}, NUMBER, RK and booleans`, () => {
    const model = parseLegacy(buildLegacyXls({ continuation }), 'xls');
    const sheet = model.sheets[0];
    assert.equal(sheet.name, '数据');
    assert.equal(sheet.rows[0][0].value, '名称');
    assert.equal(sheet.rows[1][0].value, '自研解析');
    assert.equal(sheet.rows[1][1].value, 12);
    assert.equal(sheet.rows[1][2].value, 1280.5);
    assert.equal(sheet.rows[2][1].value, true);
    assert.equal(sheet.rows[2][2].value, 42);
  });
}

test('PPT reads active persist objects and excludes unreferenced deleted slides', () => {
  const model = parseLegacy(buildLegacyPpt(), 'ppt');
  assert.equal(model.slides.length, 2);
  assert.match(slideText(model.slides[0]), /旧版 PPT 中文文本/);
  assert.match(slideText(model.slides[1]), /Local parsing/);
  assert.doesNotMatch(model.slides.map(slideText).join('\n'), /ORPHAN DELETED/);
});

test('PPT follows logical slide list order rather than file record order', () => {
  const model = parseLegacy(buildLegacyPpt({ reverseSlides: true }), 'ppt');
  assert.match(slideText(model.slides[0]), /Local parsing/);
  assert.match(slideText(model.slides[1]), /中文文本/);
});

test('encrypted legacy Office documents are explicitly rejected', () => {
  for (const [format, builder] of [['doc', buildLegacyDoc], ['xls', buildLegacyXls], ['ppt', buildLegacyPpt]]) {
    assert.throws(() => parseLegacy(builder({ encrypted: true }), format), /加密|密码/, format);
  }
});

test('a CFB file with no matching Office document stream is rejected', () => {
  const bytes = makeCompound({ Unrelated: Buffer.from('This is not an Office document') });
  for (const format of ['doc', 'xls', 'ppt']) assert.throws(() => parseLegacy(bytes, format));
});
