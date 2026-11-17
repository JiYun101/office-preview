import { binary, utf16, warn } from './legacy-binary.js';

const fixedSizes = [1, 1, 2, 4, 2, 2, -1, 3];
const alignments = ['left', 'center', 'right', 'justify', 'justify', 'justify', 'left', 'left', 'justify'];
const palette = ['#000000', '#000000', '#0000ff', '#00ffff', '#00ff00', '#ff00ff', '#ff0000', '#ffff00', '#ffffff', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#c0c0c0'];
const signed16 = value => value & 0x8000 ? value - 65536 : value;

export function readDocPrls(bytes) {
  const data = binary(bytes), result = [];
  for (let cursor = 0; cursor < bytes.length;) {
    // Style property arrays may end with a zero alignment byte.
    if (cursor + 1 === bytes.length && bytes[cursor] === 0) break;
    const id = data.u16(cursor); cursor += 2;
    let length = fixedSizes[id >>> 13];
    if (length < 0) {
      if (id === 0xd608 || id === 0xd606) { length = data.u16(cursor); cursor += 2; }
      else { length = data.u8(cursor++); }
    }
    data.need(cursor, length);
    result.push({ id, bytes: bytes.subarray(cursor, cursor + length) });
    if (result.length > 10000) throw new Error('DOC 属性记录数量超过当前上限。');
    cursor += length;
  }
  return result;
}

function changeTabs(previous = [], bytes, extended) {
  const data = binary(bytes); let cursor = 0;
  const deleted = data.u8(cursor++), stride = extended ? 4 : 2;
  const removals = [];
  for (let i = 0; i < deleted; i++) { removals.push(signed16(data.u16(cursor)) / 15); cursor += stride; }
  const count = data.u8(cursor++);
  data.need(cursor, count * 3);
  const tabs = previous.filter(tab => !removals.includes(tab.position));
  for (let i = 0; i < count; i++) {
    const position = signed16(data.u16(cursor + i * 2)) / 15;
    const descriptor = data.u8(cursor + count * 2 + i);
    const tab = { position, align: ['left', 'center', 'right', 'decimal', 'bar'][descriptor & 7] || 'left', leader: ['none', 'dot', 'hyphen', 'underscore', 'heavy', 'middleDot'][descriptor >>> 3 & 7] || 'none' };
    const old = tabs.findIndex(item => item.position === position);
    if (old >= 0) tabs.splice(old, 1);
    tabs.push(tab);
  }
  return tabs.sort((a, b) => a.position - b.position);
}

function toggle(previous, value) { return value === 0x80 ? previous : value === 0x81 ? !previous : value === 1; }

export function applyDocPrls(base, prls, context = {}) {
  const result = { ...base };
  for (const prl of prls) {
    const data = binary(prl.bytes), byte = () => data.u8(0), short = () => data.u16(0), signed = () => signed16(short());
    switch (prl.id) {
      case 0x0835: result.bold = toggle(result.bold, byte()); break;
      case 0x0836: result.italic = toggle(result.italic, byte()); break;
      case 0x0837: result.strike = toggle(result.strike, byte()); break;
      case 0x083c: result._hidden = toggle(result._hidden, byte()); break;
      case 0x2a3e: result.underline = byte() !== 0; break;
      case 0x4a43: if (short() > 0) result.fontSize = short() * 2 / 3; break;
      case 0x4a4f: result._fontAscii = short(); break;
      case 0x4a50: result._fontEastAsia = short(); break;
      case 0x4a51: result._fontOther = short(); break;
      case 0x4a30: result._characterStyle = short(); break;
      case 0x8840: result.letterSpacing = signed() / 15; break;
      case 0x2a42: if (palette[byte()]) result.color = palette[byte()]; break;
      case 0x6870: { const value = data.u32(0); if (value !== 0xff000000) result.color = `#${[value & 255, value >>> 8 & 255, value >>> 16 & 255].map(value => value.toString(16).padStart(2, '0')).join('')}`; break; }
      case 0x286f: result._fontHint = byte(); break;
      case 0x6a03: result._pictureLocation = data.u32(0); break;
      case 0x0806: result._pictureIsData = byte() === 1; break;
      case 0x2403: case 0x2461: result.align = alignments[byte()] || 'left'; break;
      case 0x840e: case 0x845d: result.indentRight = signed() / 15; break;
      case 0x840f: case 0x845e: result.indentLeft = signed() / 15; break;
      case 0x8411: case 0x8460: result.firstLineIndent = signed() / 15; break;
      case 0xa413: result.spaceBefore = short() / 15; break;
      case 0xa414: result.spaceAfter = short() / 15; break;
      case 0x6412: { const value = signed16(data.u16(0)); if (data.u16(2)) { result.lineHeightRatio = value / 240; delete result.lineHeight; } else { result.lineHeight = Math.abs(value) / 15; result.lineHeightRule = value < 0 ? 'exact' : 'atLeast'; delete result.lineHeightRatio; } break; }
      case 0x2405: result.keepLines = byte() === 1; break;
      case 0x2406: result.keepNext = byte() === 1; break;
      case 0x2407: result.pageBreakBefore = byte() === 1; break;
      case 0x2431: result.snapToGrid = byte() === 1; break;
      case 0x260a: result._listLevel = byte(); break;
      case 0x460b: result._listId = short(); break;
      case 0xc60d: result.tabs = changeTabs(result.tabs, prl.bytes, false); break;
      case 0xc615: result.tabs = changeTabs(result.tabs, prl.bytes, true); break;
      case 0xb01f: result.width = short() / 15; break;
      case 0xb020: result.height = short() / 15; break;
      case 0xb021: result.marginLeft = short() / 15; break;
      case 0xb022: result.marginRight = short() / 15; break;
      case 0x9023: result.marginTop = Math.abs(signed()) / 15; break;
      case 0x9024: result.marginBottom = Math.abs(signed()) / 15; break;
      case 0x9031: result.linePitch = signed() / 15; break;
      case 0x6646: case 0x6645: {
        if (context.data && (context.depth || 0) < 4) {
          const input = binary(context.data), position = data.u32(0), length = input.u16(position);
          Object.assign(result, applyDocPrls(result, readDocPrls(input.slice(position + 2, length)), { ...context, depth: (context.depth || 0) + 1 }));
        }
        break;
      }
    }
  }
  return result;
}

export function createDocStyleReader(wordBytes, tableBytes, pair, dataBytes, warnings) {
  const word = binary(wordBytes), table = binary(tableBytes);
  const fonts = [];
  const fontPart = pair(15);
  if (fontPart.length) {
    table.need(fontPart.offset, fontPart.length);
    const end = fontPart.offset + fontPart.length, count = table.u16(fontPart.offset);
    if (count > 4096 || table.u16(fontPart.offset + 2)) throw new Error('DOC 字体表头无效。');
    let cursor = fontPart.offset + 4;
    for (let i = 0; i < count; i++) {
      const length = table.u8(cursor++);
      if (length < 39 || cursor + length > end) throw new Error('DOC 字体记录不完整。');
      fonts.push(utf16.decode(table.slice(cursor + 39, length - 39)).split('\0')[0]);
      cursor += length;
    }
  }
  const defaults = { fontSize: 40 / 3, _fontAscii: 0, _fontEastAsia: 0, _fontOther: 0, color: '#000000' };
  const styles = [];
  const sheet = pair(1);
  if (sheet.length) {
    table.need(sheet.offset, sheet.length);
    const end = sheet.offset + sheet.length, headerSize = table.u16(sheet.offset), count = table.u16(sheet.offset + 2), baseSize = table.u16(sheet.offset + 4);
    if (headerSize < 18 || count > 4096 || baseSize < 10 || sheet.offset + 2 + headerSize > end) throw new Error('DOC 样式表头无效。');
    defaults._fontAscii = table.u16(sheet.offset + 14);
    defaults._fontEastAsia = table.u16(sheet.offset + 16);
    defaults._fontOther = table.u16(sheet.offset + 18);
    let cursor = sheet.offset + 2 + headerSize;
    for (let index = 0; index < count; index++) {
      const length = table.u16(cursor); cursor += 2;
      if (!length) { styles.push(null); continue; }
      const start = cursor, limit = start + length;
      if (length < baseSize + 4 || limit > end) throw new Error('DOC 样式描述长度无效。');
      const kind = table.u16(start + 2) & 15, parent = table.u16(start + 2) >>> 4, cupx = table.u16(start + 4) & 15;
      const countName = table.u16(start + baseSize), nameLength = countName * 2;
      if (start + baseSize + 4 + nameLength > limit) throw new Error('DOC 样式名称不完整。');
      const name = utf16.decode(table.slice(start + baseSize + 2, nameLength));
      const entry = { index, kind, parent, name, sti: table.u16(start) & 4095, paragraph: [], character: [] };
      let position = start + baseSize + 4 + nameLength;
      for (let upx = 0; upx < cupx; upx++) {
        if ((position - start) & 1) position++;
        if (position + 2 > limit) throw new Error('DOC 样式属性长度不完整。');
        const bytes = table.u16(position); position += 2;
        if (position + bytes > limit) throw new Error('DOC 样式属性越过记录边界。');
        if (kind === 1 && upx === 0) entry.paragraph = bytes >= 2 ? readDocPrls(table.slice(position + 2, bytes - 2)) : [];
        if (kind === 1 && upx === 1 || kind === 2 && upx === 0) entry.character = readDocPrls(table.slice(position, bytes));
        position += bytes;
      }
      styles.push(entry); cursor = limit;
    }
  }
  const context = { data: dataBytes };
  const cache = new Map();
  const resolve = (index, seen = new Set()) => {
    if (cache.has(index)) return cache.get(index);
    const entry = styles[index];
    if (!entry) return { paragraph: {}, character: { ...defaults } };
    if (seen.has(index) || seen.size > 64) throw new Error('DOC 样式继承包含循环或嵌套过深。');
    seen.add(index);
    const base = entry.parent !== 4095 ? resolve(entry.parent, seen) : { paragraph: {}, character: { ...defaults } };
    const result = { paragraph: applyDocPrls(base.paragraph, entry.paragraph, context), character: applyDocPrls(base.character, entry.character, context), name: entry.name, sti: entry.sti };
    cache.set(index, result); return result;
  };
  const readFkps = (pairIndex, paragraph) => {
    const part = pair(pairIndex);
    if (!part.length) return [];
    table.need(part.offset, part.length);
    if (part.length < 4 || (part.length - 4) % 8) throw new Error('DOC 格式页索引长度无效。');
    const count = (part.length - 4) / 8, runs = [];
    for (let i = 0; i < count; i++) {
      const page = (table.u32(part.offset + (count + 1) * 4 + i * 4) & 0x3fffff) * 512;
      word.need(page, 512);
      const n = word.u8(page + 511), positionsEnd = page + (n + 1) * 4;
      if (!n || n > (paragraph ? 29 : 101)) throw new Error('DOC 格式页运行数量无效。');
      for (let j = 0; j < n; j++) {
        const start = word.u32(page + j * 4), end = word.u32(page + (j + 1) * 4);
        if (start >= end) throw new Error('DOC 格式运行范围无效。');
        const offset = word.u8(positionsEnd + j * (paragraph ? 13 : 1)) * 2;
        let prls = [], style = 0;
        if (offset) {
          if (offset < (n + 1) * 4 + n * (paragraph ? 13 : 1) || offset >= 511) throw new Error('DOC 属性与格式页索引重叠。');
          let position = page + offset, length = word.u8(position++);
          if (paragraph) {
            if (!length) length = word.u8(position++) * 2; else length = length * 2 - 1;
            if (length < 2) throw new Error('DOC 段落属性缺少样式索引。');
            if (position + length > page + 511) throw new Error('DOC 段落属性超出格式页。');
            style = word.u16(position); position += 2; length -= 2;
          }
          if (position + length > page + 511) throw new Error('DOC 字符属性超出格式页。');
          prls = readDocPrls(word.slice(position, length));
        }
        runs.push({ start, end, style, prls });
        if (runs.length > 200000) throw new Error('DOC 格式运行数量超过当前上限。');
      }
    }
    return runs.sort((a, b) => a.start - b.start);
  };
  const characterRuns = readFkps(12, false), paragraphRuns = readFkps(13, true);
  const characterStyle = (index, base, suppressHyperlink, seen = new Set()) => {
    const entry = styles[index];
    if (!entry || entry.kind !== 2) return { ...base };
    if (seen.has(index) || seen.size > 64) throw new Error('DOC 字符样式继承包含循环。');
    seen.add(index);
    const inherited = entry.parent !== 4095 ? characterStyle(entry.parent, base, suppressHyperlink, seen) : { ...base };
    const hyperlink = entry.sti === 85 || entry.sti === 86 || /^(?:hyperlink|followedhyperlink|超链接|已访问的超链接)$/i.test(entry.name);
    // Character styles contribute deltas; they must not reset the paragraph's
    // font/size to document defaults when the character style omits them.
    return suppressHyperlink && hyperlink ? inherited : applyDocPrls(inherited, entry.character, context);
  };
  const find = (runs, fc) => {
    let low = 0, high = runs.length;
    while (low < high) { const middle = (low + high) >>> 1; if (runs[middle].start <= fc) low = middle + 1; else high = middle; }
    const run = runs[low - 1]; return run && fc < run.end ? run : null;
  };
  const publicCharacter = (properties, text = '') => {
    const eastAsia = properties._fontHint === 1 || /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(text);
    const family = fonts[eastAsia ? properties._fontEastAsia : properties._fontAscii] || fonts[properties._fontOther] || fonts[properties._fontEastAsia];
    const result = { ...(family ? { fontFamily: family } : {}) };
    for (const key of ['fontSize', 'bold', 'italic', 'underline', 'strike', 'color', 'letterSpacing']) if (properties[key] !== undefined) result[key] = properties[key];
    return result;
  };
  const sections = [];
  const sectionPart = pair(6);
  if (sectionPart.length) {
    const n = (sectionPart.length - 4) / 16;
    if (!Number.isInteger(n) || n < 1 || n > 1000) throw new Error('DOC 节索引长度无效。');
    table.need(sectionPart.offset, sectionPart.length);
    for (let i = 0; i < n; i++) {
      const start = table.u32(sectionPart.offset + i * 4), end = table.u32(sectionPart.offset + (i + 1) * 4);
      const position = table.u32(sectionPart.offset + (n + 1) * 4 + i * 12 + 2);
      let page = { width: 816, height: 1056, marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96 };
      if (position !== 0xffffffff) {
        const length = word.u16(position);
        page = applyDocPrls(page, readDocPrls(word.slice(position + 2, length)), context);
      }
      sections.push({ start, end, page });
    }
  }
  return {
    fonts, sections,
    paragraph(fc) {
      const run = find(paragraphRuns, fc), index = run?.style || 0, inherited = resolve(index);
      return { ...inherited, paragraph: applyDocPrls(inherited.paragraph, run?.prls || [], context), baseParagraph: inherited.paragraph, direct: run?.prls || [], index };
    },
    character(fc, inherited, suppressHyperlink = false) {
      const run = find(characterRuns, fc), prls = run?.prls || [];
      const index = prls.findLast(prl => prl.id === 0x4a30);
      const base = index ? characterStyle(binary(index.bytes).u16(0), inherited, suppressHyperlink) : inherited;
      return { properties: applyDocPrls(base, prls, context), end: run?.end ?? Number.MAX_SAFE_INTEGER };
    },
    publicCharacter,
    defaultStyle: publicCharacter(resolve(0).character, '中'),
    stats: { fonts: fonts.length, styles: styles.filter(Boolean).length, characterRuns: characterRuns.length, paragraphRuns: paragraphRuns.length }
  };
}
