// Original sfnt reader/repacker. Table layout, cmap 4/12 and checksum rules:
// https://learn.microsoft.com/en-us/typography/opentype/spec/otff
// https://learn.microsoft.com/en-us/typography/opentype/spec/cmap
const align4 = value => Math.ceil(value / 4) * 4;
const privateCode = glyph => glyph <= 65533 ? 0xf0000 + glyph : 0x100000 + glyph - 65534;

function checksum(bytes) {
  let sum = 0;
  for (let index = 0; index < bytes.length; index += 4) {
    const word = ((bytes[index] || 0) * 0x1000000) + ((bytes[index + 1] || 0) << 16) + ((bytes[index + 2] || 0) << 8) + (bytes[index + 3] || 0);
    sum = (sum + word) >>> 0;
  }
  return sum;
}

function cmapReader(table, glyphCount) {
  if (!table) return () => 0;
  const view = new DataView(table.buffer, table.byteOffset, table.byteLength);
  const range = (offset, length) => {
    if (!Number.isInteger(offset) || offset < 0 || offset + length > table.length) throw new Error('PDF 嵌入字体 cmap 范围无效。');
  };
  range(0, 4);
  const count = view.getUint16(2), candidates = [];
  range(4, count * 8);
  for (let index = 0; index < count; index += 1) {
    const record = 4 + index * 8;
    const platform = view.getUint16(record), encoding = view.getUint16(record + 2), offset = view.getUint32(record + 4);
    if (!(platform === 0 || platform === 3 && [0, 1, 10].includes(encoding))) continue;
    range(offset, 2);
    const format = view.getUint16(offset);
    if (![0, 4, 6, 12].includes(format)) continue;
    range(offset, format === 12 ? 16 : 6);
    const length = format === 12 ? view.getUint32(offset + 4) : view.getUint16(offset + 2);
    range(offset, length);
    const sub = new DataView(table.buffer, table.byteOffset + offset, length);
    const check = (position, size) => { if (position + size > length) throw new Error('PDF 嵌入字体 cmap 子表被截断。'); };
    let lookup;
    if (format === 12) {
      const groups = sub.getUint32(12);
      if (groups > 100000) throw new Error('PDF 嵌入字体 cmap 分组超过当前上限。');
      check(16, groups * 12);
      let previous = -1;
      for (let group = 0; group < groups; group += 1) {
        const p = 16 + group * 12, start = sub.getUint32(p), end = sub.getUint32(p + 4), first = sub.getUint32(p + 8);
        if (start <= previous || end < start || end > 0x10ffff || first + end - start >= glyphCount) throw new Error('PDF 嵌入字体 cmap 字形范围无效。');
        previous = end;
      }
      lookup = code => {
        let low = 0, high = groups - 1;
        while (low <= high) {
          const middle = (low + high) >>> 1, p = 16 + middle * 12;
          if (code < sub.getUint32(p)) high = middle - 1;
          else if (code > sub.getUint32(p + 4)) low = middle + 1;
          else return sub.getUint32(p + 8) + code - sub.getUint32(p);
        }
        return 0;
      };
    } else if (format === 4) {
      check(0, 16);
      const segments = sub.getUint16(6) / 2;
      if (!Number.isInteger(segments) || !segments) throw new Error('PDF 嵌入字体 cmap 分段数量无效。');
      check(16, segments * 8);
      const starts = 16 + segments * 2, deltas = starts + segments * 2, offsets = deltas + segments * 2;
      let previous = -1;
      for (let segment = 0; segment < segments; segment += 1) {
        const start = sub.getUint16(starts + segment * 2), end = sub.getUint16(14 + segment * 2);
        if (start <= previous || end < start) throw new Error('PDF 嵌入字体 cmap 分段范围无效。');
        const jump = sub.getUint16(offsets + segment * 2);
        if (jump) check(offsets + segment * 2 + jump, (end - start + 1) * 2);
        previous = end;
      }
      lookup = code => {
        if (code > 65535) return 0;
        let low = 0, high = segments - 1;
        while (low < high) { const middle = (low + high) >>> 1; if (code > sub.getUint16(14 + middle * 2)) low = middle + 1; else high = middle; }
        if (code < sub.getUint16(starts + low * 2) || code > sub.getUint16(14 + low * 2)) return 0;
        const delta = sub.getInt16(deltas + low * 2), at = offsets + low * 2, jump = sub.getUint16(at);
        if (!jump) return (code + delta) & 65535;
        const glyph = sub.getUint16(at + jump + (code - sub.getUint16(starts + low * 2)) * 2);
        return glyph ? (glyph + delta) & 65535 : 0;
      };
    } else if (format === 6) {
      check(0, 10);
      const first = sub.getUint16(6), size = sub.getUint16(8);
      check(10, size * 2);
      lookup = code => code >= first && code < first + size ? sub.getUint16(10 + (code - first) * 2) : 0;
    } else {
      check(6, 256);
      lookup = code => code <= 255 ? sub.getUint8(6 + code) : 0;
    }
    candidates.push({ symbol: platform === 3 && encoding === 0, priority: format === 12 ? 2 : 1, lookup });
  }
  candidates.sort((a, b) => b.priority - a.priority);
  return (code, symbol = false) => {
    for (const candidate of candidates) {
      if (candidate.symbol !== symbol) continue;
      const glyph = candidate.lookup(code);
      if (glyph > 0 && glyph < glyphCount) return glyph;
    }
    return 0;
  };
}

function privateCmap(glyphCount) {
  // PDF already selects glyphs and positions them; avoid browser shaping or
  // ligatures by mapping each selected original GID to one private scalar.
  const groups = [[0xf0000, 0xf0000 + Math.min(glyphCount - 1, 65533), 0]];
  if (glyphCount > 65534) groups.push([0x100000, 0x100000 + glyphCount - 65535, 65534]);
  const bytes = new Uint8Array(20 + 16 + groups.length * 12), view = new DataView(bytes.buffer);
  view.setUint16(2, 2);
  view.setUint16(4, 0); view.setUint16(6, 4); view.setUint32(8, 20);
  view.setUint16(12, 3); view.setUint16(14, 10); view.setUint32(16, 20);
  view.setUint16(20, 12); view.setUint32(24, bytes.length - 20); view.setUint32(32, groups.length);
  groups.forEach((group, index) => group.forEach((value, field) => view.setUint32(36 + index * 12 + field * 4, value)));
  return bytes;
}

export function preparePdfFontProgram(input, maxBytes = 128 * 1024 * 1024) {
  if (!(input instanceof Uint8Array) || input.length < 12) throw new Error('PDF 嵌入字体程序被截断。');
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength), signature = view.getUint32(0);
  if (![0x00010000, 0x74727565, 0x4f54544f].includes(signature)) throw new Error('PDF 嵌入字体尚不是支持的 TrueType/OpenType sfnt。');
  const count = view.getUint16(4);
  if (!count || count > 256 || 12 + count * 16 > input.length) throw new Error('PDF 嵌入字体表目录无效。');
  const tables = new Map(), spans = [];
  for (let index = 0; index < count; index += 1) {
    const at = 12 + index * 16, tag = String.fromCharCode(...input.subarray(at, at + 4));
    const offset = view.getUint32(at + 8), length = view.getUint32(at + 12);
    if (tables.has(tag) || offset < 12 + count * 16 || offset + length > input.length) throw new Error('PDF 嵌入字体表范围或名称无效。');
    if (length) spans.push([offset, offset + length]);
    tables.set(tag, input.subarray(offset, offset + length));
  }
  spans.sort((a, b) => a[0] - b[0]);
  if (spans.some((span, index) => index && span[0] < spans[index - 1][1])) throw new Error('PDF 嵌入字体表范围重叠。');
  if (!tables.has('head') || tables.get('head').length < 54 || !tables.has('maxp') || tables.get('maxp').length < 6) throw new Error('PDF 嵌入字体缺少有效的 head/maxp。');
  const glyphCount = new DataView(tables.get('maxp').buffer, tables.get('maxp').byteOffset).getUint16(4);
  if (!glyphCount) throw new Error('PDF 嵌入字体没有字形。');
  const glyphIndex = cmapReader(tables.get('cmap'), glyphCount);
  const os2 = tables.get('OS/2'), head = tables.get('head').slice();
  const weight = os2?.length >= 6 ? new DataView(os2.buffer, os2.byteOffset).getUint16(4) : undefined;
  const macStyle = new DataView(head.buffer).getUint16(44);
  new DataView(head.buffer).setUint32(8, 0);
  tables.set('head', head); tables.set('cmap', privateCmap(glyphCount)); tables.delete('DSIG');
  // PDF text has already been shaped and positioned into explicit GIDs. Browser
  // substitution/positioning tables must not change those selected glyphs again.
  for (const tag of ['GSUB', 'GPOS', 'kern', 'kerx', 'mort', 'morx']) tables.delete(tag);
  const entries = Array.from(tables).sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0);
  const length = 12 + entries.length * 16 + entries.reduce((sum, [, data]) => sum + align4(data.length), 0);
  if (length > maxBytes) throw new Error('PDF 嵌入字体重建超过当前 128MB 总量上限。');
  const bytes = new Uint8Array(length), out = new DataView(bytes.buffer), power = 2 ** Math.floor(Math.log2(entries.length));
  out.setUint32(0, signature === 0x74727565 ? 0x00010000 : signature); out.setUint16(4, entries.length);
  out.setUint16(6, power * 16); out.setUint16(8, Math.log2(power)); out.setUint16(10, entries.length * 16 - power * 16);
  let offset = 12 + entries.length * 16, headOffset;
  entries.forEach(([tag, data], index) => {
    const at = 12 + index * 16;
    bytes.set(Array.from(tag, character => character.charCodeAt(0)), at);
    out.setUint32(at + 4, checksum(data)); out.setUint32(at + 8, offset); out.setUint32(at + 12, data.length);
    bytes.set(data, offset); if (tag === 'head') headOffset = offset;
    offset += align4(data.length);
  });
  out.setUint32(headOffset + 8, (0xb1b0afba - checksum(bytes)) >>> 0);
  return { bytes, format: signature === 0x4f54544f ? 'opentype' : 'truetype', glyphCount, glyphIndex, weight: weight >= 1 && weight <= 1000 ? weight : macStyle & 1 ? 700 : undefined, italic: !!(macStyle & 2), renderText: glyph => String.fromCodePoint(privateCode(glyph)) };
}
