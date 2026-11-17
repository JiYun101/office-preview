import test from 'node:test';
import assert from 'node:assert/strict';
import { preparePdfFontProgram } from '../src/parsers/pdf-font-program.js';
import { loadPdfFont } from '../src/parsers/pdf-font.js';
import { buildPdfFontFixture } from '../scripts/pdf-font-fixture.mjs';

const name = name => ({ name });
const stream = (bytes, dictionary = {}) => ({ stream: bytes, dictionary });
const unicodeMap = text => stream(new TextEncoder().encode(`1 begincodespacerange <0000> <ffff> endcodespacerange ${text.length} beginbfchar ${Array.from(text, (character, index) => `<${(index + 1).toString(16).padStart(4, '0')}> <${character.charCodeAt(0).toString(16).padStart(4, '0')}>`).join(' ')} endbfchar`));
function document() {
  return { warnings: [], budget: { tokens: 0 }, totalDecoded: 0, resolve: async value => value, stream: async value => value.stream, warn(message) { if (!this.warnings.includes(message)) this.warnings.push(message); } };
}
function cidFont(bytes, extra = {}) {
  return { Subtype: name('Type0'), BaseFont: name('ABCDEF+Fixture-Regular'), Encoding: name('Identity-H'), ToUnicode: unicodeMap('AB'), DescendantFonts: [{ Subtype: name('CIDFontType2'), DW: 1000, W: [1, [900, 1000]], CIDToGIDMap: name('Identity'), FontDescriptor: { FontFile2: stream(bytes) }, ...extra }] };
}
function tables(bytes) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength), entries = new Map();
  for (let index = 0; index < view.getUint16(4); index += 1) {
    const at = 12 + index * 16, tag = String.fromCharCode(...bytes.subarray(at, at + 4)), offset = view.getUint32(at + 8), length = view.getUint32(at + 12);
    entries.set(tag, { at, offset, bytes: bytes.subarray(offset, offset + length) });
  }
  return entries;
}
function checksum(bytes) {
  let result = 0;
  for (let at = 0; at < bytes.length; at += 4) result = (result + (bytes[at] || 0) * 0x1000000 + ((bytes[at + 1] || 0) << 16) + ((bytes[at + 2] || 0) << 8) + (bytes[at + 3] || 0)) >>> 0;
  return result;
}

test('sfnt rebuild preserves real outlines, metrics and hints while adding direct private GID mappings', () => {
  const original = buildPdfFontFixture({ weight: 700, italic: true });
  const prepared = preparePdfFontProgram(original);
  assert.equal(prepared.weight, 700);
  assert.equal(prepared.italic, true);
  assert.equal(prepared.glyphIndex(65), 1);
  assert.equal(prepared.glyphIndex(66), 2);
  const rebuilt = preparePdfFontProgram(prepared.bytes);
  assert.equal(rebuilt.glyphIndex(prepared.renderText(1).codePointAt(0)), 1);
  assert.equal(rebuilt.glyphIndex(prepared.renderText(2).codePointAt(0)), 2);
  const before = tables(original), after = tables(prepared.bytes);
  for (const tag of ['glyf', 'loca', 'hmtx', 'hhea', 'maxp']) assert.deepEqual(after.get(tag).bytes, before.get(tag).bytes, tag);
  assert.equal(checksum(prepared.bytes), 0xb1b0afba, 'The complete sfnt checksum must match OpenType checkSumAdjustment.');
  const view = new DataView(prepared.bytes.buffer);
  for (const [tag, entry] of after) {
    const data = entry.bytes.slice();
    if (tag === 'head') new DataView(data.buffer).setUint32(8, 0);
    assert.equal(checksum(data), view.getUint32(entry.at + 4), `${tag} table checksum`);
  }
  assert.deepEqual(original, buildPdfFontFixture({ weight: 700, italic: true }), 'The source font is never modified.');
});

test('Identity-H uses embedded original GIDs even when a subset has no Unicode cmap', async () => {
  const doc = document(), font = await loadPdfFont(doc, cidFont(buildPdfFontFixture({ weight: 700, includeCmap: false })));
  const glyphs = font.decode(Uint8Array.of(0, 1, 0, 2));
  assert.equal(font.weight, 700);
  assert.equal(font.style, 'normal');
  assert.equal(font.program.format, 'truetype');
  assert.equal(font.family, font.id);
  assert.deepEqual(glyphs.map(({ code, text, width }) => ({ code, text, width })), [{ code: 1, text: 'A', width: 900 }, { code: 2, text: 'B', width: 1000 }]);
  assert.deepEqual(glyphs.map(glyph => glyph.renderText.codePointAt(0)), [0xf0001, 0xf0002]);
  assert.deepEqual(doc.warnings, []);
});

test('CIDToGIDMap streams select the embedded glyph and missing entries retain semantic text with a warning', async () => {
  const doc = document(), font = await loadPdfFont(doc, cidFont(buildPdfFontFixture(), { CIDToGIDMap: stream(Uint8Array.of(0, 0, 0, 2, 0, 1)) }));
  assert.deepEqual(font.decode(Uint8Array.of(0, 1, 0, 2)).map(glyph => glyph.renderText.codePointAt(0)), [0xf0002, 0xf0001]);
  const missing = font.decode(Uint8Array.of(0, 3))[0];
  assert.equal(missing.renderText, undefined);
  assert.ok(doc.warnings.some(message => message.includes('没有可用的嵌入字形')));
});

test('simple TrueType and symbolic encoding select the original cmap glyph without relying on system fallback', async () => {
  for (const symbolic of [false, true]) {
    const doc = document(), font = await loadPdfFont(doc, { Subtype: name('TrueType'), BaseFont: name('Fixture'), Encoding: name('WinAnsiEncoding'), FirstChar: 65, Widths: [900, 1000], FontDescriptor: { Flags: symbolic ? 4 : 32, FontFile2: stream(buildPdfFontFixture({ symbolic })) } });
    assert.deepEqual(font.decode(Uint8Array.of(65, 66)).map(glyph => glyph.renderText.codePointAt(0)), [0xf0001, 0xf0002]);
    assert.equal(doc.warnings.length, 0);
  }
});

test('FontFile3 OpenType streams expose font resources, while unsupported Type1C explicitly falls back', async () => {
  for (const type of ['OpenType', 'Type1C']) {
    const doc = document(), font = await loadPdfFont(doc, { Subtype: name('TrueType'), BaseFont: name('Fixture-BoldItalic'), Encoding: name('WinAnsiEncoding'), FirstChar: 65, Widths: [900], FontDescriptor: { FontFile3: stream(buildPdfFontFixture(), { Subtype: name(type) }) } });
    assert.equal(!!font.program, type === 'OpenType');
    if (type === 'Type1C') { assert.equal(font.weight, 700); assert.equal(font.style, 'italic'); assert.ok(doc.warnings.some(message => message.includes('Type1/CFF'))); }
  }
});

test('font metadata preserves bold/italic and descriptor weights on the system-font fallback path', async () => {
  const cases = [
    ['ABCDEF+Helvetica-BoldOblique', {}, 700, 'italic'],
    ['Fixture', { FontWeight: 600, Flags: 64 }, 600, 'italic'],
    ['Fixture', { Flags: 262144 }, 700, 'normal'],
    ['Fixture', { StemV: 140 }, 700, 'normal'],
    ['Fixture-Regular', { StemV: 140 }, 400, 'normal'],
    ['Fixture-ExtraLight', {}, 200, 'normal'],
  ];
  for (const [base, descriptor, weight, style] of cases) {
    const font = await loadPdfFont(document(), { Subtype: name('Type1'), BaseFont: name(base), FirstChar: 65, Widths: [900], FontDescriptor: descriptor });
    assert.equal(font.weight, weight, base); assert.equal(font.style, style, base); assert.equal(font.program, undefined);
  }
});

test('font table bounds, overlapping data, malformed cmap and rebuild budget are checked before use', () => {
  const valid = buildPdfFontFixture(), entries = tables(valid);
  const truncate = valid.slice(0, -1);
  assert.throws(() => preparePdfFontProgram(truncate), /表范围/);
  const overlap = valid.slice();
  new DataView(overlap.buffer).setUint32(entries.get('name').at + 8, entries.get('glyf').offset);
  assert.throws(() => preparePdfFontProgram(overlap), /重叠|表范围/);
  const cmap = valid.slice();
  new DataView(cmap.buffer).setUint32(entries.get('cmap').offset + 8, 0xffffffff);
  assert.throws(() => preparePdfFontProgram(cmap), /cmap/);
  assert.throws(() => preparePdfFontProgram(valid, 100), /128MB/);
  assert.throws(() => preparePdfFontProgram(Uint8Array.of(0, 1)), /截断/);
});
