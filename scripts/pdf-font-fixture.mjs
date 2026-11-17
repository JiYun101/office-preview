// Self-authored font fixture: only an empty .notdef, a triangle and a square.
// No system/user font data is included. This module also runs in a browser.
const padded = length => Math.ceil(length / 4) * 4;
function sum(bytes) {
  let result = 0;
  for (let at = 0; at < bytes.length; at += 4) result = (result + (bytes[at] || 0) * 0x1000000 + ((bytes[at + 1] || 0) << 16) + ((bytes[at + 2] || 0) << 8) + (bytes[at + 3] || 0)) >>> 0;
  return result;
}
function table(length, fill) { const bytes = new Uint8Array(length); fill(new DataView(bytes.buffer), bytes); return bytes; }
function contour(points) {
  return table(padded(14 + points.length * 5), view => {
    view.setInt16(0, 1); view.setInt16(6, 800); view.setInt16(8, 800); view.setUint16(10, points.length - 1);
    let x = 0, y = 0;
    points.forEach(([nextX, nextY], index) => {
      view.setUint8(14 + index, 1);
      view.setInt16(14 + points.length + index * 2, nextX - x);
      view.setInt16(14 + points.length * 3 + index * 2, nextY - y);
      x = nextX; y = nextY;
    });
  });
}

export function buildPdfFontFixture({ weight = 400, italic = false, symbolic = false, includeCmap = true } = {}) {
  const triangle = contour([[0, 0], [800, 0], [400, 800]]), square = contour([[0, 0], [800, 0], [800, 800], [0, 800]]);
  const glyf = new Uint8Array(triangle.length + square.length); glyf.set(triangle); glyf.set(square, triangle.length);
  const style = `${weight >= 600 ? 'Bold' : 'Regular'}${italic ? 'Italic' : ''}`;
  const names = ['OfficePreviewFixture', style, `OfficePreviewFixture-${style}`, `OfficePreviewFixture ${style}`, `OfficePreviewFixture-${style}`];
  const namesLength = names.reduce((length, name) => length + name.length * 2, 0);
  const tables = new Map([
    ['head', table(54, view => { view.setUint32(0, 0x00010000); view.setUint32(4, 0x00010000); view.setUint32(12, 0x5f0f3cf5); view.setUint16(18, 1000); view.setInt16(40, 800); view.setInt16(42, 800); view.setUint16(44, (weight >= 600 ? 1 : 0) | (italic ? 2 : 0)); view.setUint16(46, 8); view.setInt16(48, 2); view.setInt16(50, 1); })],
    ['hhea', table(36, view => { view.setUint32(0, 0x00010000); view.setInt16(4, 800); view.setInt16(6, -200); view.setUint16(10, 1000); view.setInt16(16, 800); view.setInt16(18, 1); view.setUint16(34, 3); })],
    ['maxp', table(32, view => { view.setUint32(0, 0x00010000); view.setUint16(4, 3); view.setUint16(6, 4); view.setUint16(8, 1); view.setUint16(14, 2); })],
    ['hmtx', table(12, view => { [500, 900, 1000].forEach((advance, index) => view.setUint16(index * 4, advance)); })],
    ['loca', table(16, view => { [0, 0, triangle.length, glyf.length].forEach((offset, index) => view.setUint32(index * 4, offset)); })],
    ['glyf', glyf],
    ['OS/2', table(78, (view, bytes) => { view.setInt16(2, 900); view.setUint16(4, weight); view.setUint16(6, 5); view.setUint32(42, 1); bytes.set([79, 80, 82, 86], 58); view.setUint16(62, (weight >= 600 ? 32 : 64) | (italic ? 1 : 0)); view.setUint16(64, 65); view.setUint16(66, 66); view.setInt16(68, 800); view.setInt16(70, -200); view.setUint16(74, 800); view.setUint16(76, 200); })],
    ['post', table(32, view => { view.setUint32(0, 0x00030000); view.setInt16(8, -100); view.setInt16(10, 50); })],
    ['name', table(6 + names.length * 12 + namesLength, view => {
      view.setUint16(2, names.length); view.setUint16(4, 6 + names.length * 12);
      let offset = 0;
      names.forEach((name, index) => {
        const at = 6 + index * 12; view.setUint16(at, 3); view.setUint16(at + 2, 1); view.setUint16(at + 4, 0x409); view.setUint16(at + 6, [1, 2, 3, 4, 6][index]); view.setUint16(at + 8, name.length * 2); view.setUint16(at + 10, offset);
        Array.from(name).forEach((character, index) => view.setUint16(6 + names.length * 12 + offset + index * 2, character.charCodeAt(0)));
        offset += name.length * 2;
      });
    })],
  ]);
  if (includeCmap) tables.set('cmap', table(44, view => {
    view.setUint16(2, 1); view.setUint16(4, 3); view.setUint16(6, symbolic ? 0 : 1); view.setUint32(8, 12);
    view.setUint16(12, 4); view.setUint16(14, 32); view.setUint16(18, 4); view.setUint16(20, 4); view.setUint16(22, 1);
    view.setUint16(26, symbolic ? 0xf042 : 66); view.setUint16(28, 65535); view.setUint16(32, symbolic ? 0xf041 : 65); view.setUint16(34, 65535); view.setInt16(36, symbolic ? 0x1000 - 64 : -64); view.setInt16(38, 1);
  }));
  const entries = Array.from(tables).sort(([a], [b]) => a < b ? -1 : a > b ? 1 : 0), count = entries.length;
  const output = new Uint8Array(12 + count * 16 + entries.reduce((length, [, bytes]) => length + padded(bytes.length), 0)), view = new DataView(output.buffer);
  const power = 2 ** Math.floor(Math.log2(count));
  view.setUint32(0, 0x00010000); view.setUint16(4, count); view.setUint16(6, power * 16); view.setUint16(8, Math.log2(power)); view.setUint16(10, count * 16 - power * 16);
  let offset = 12 + count * 16, head;
  entries.forEach(([tag, bytes], index) => {
    const at = 12 + index * 16; output.set(Array.from(tag, character => character.charCodeAt(0)), at); view.setUint32(at + 4, sum(bytes)); view.setUint32(at + 8, offset); view.setUint32(at + 12, bytes.length); output.set(bytes, offset); if (tag === 'head') head = offset; offset += padded(bytes.length);
  });
  view.setUint32(head + 8, (0xb1b0afba - sum(output)) >>> 0);
  return output;
}
