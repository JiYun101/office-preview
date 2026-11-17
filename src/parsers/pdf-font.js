import { PdfLexer, asName, isDictionary, isString, numeric, integer } from './pdf-syntax.js';
import { preparePdfFontProgram } from './pdf-font-program.js';

let nextFontId = 0;

const HEX = bytes => Array.from(bytes, byte => byte.toString(16).padStart(2, '0')).join('');
const bytesInteger = bytes => bytes.reduce((result, value) => result * 256 + value, 0);
const GLYPHS = { space: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&', quotesingle: "'", quoteright: '’', quoteleft: '‘', parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-', period: '.', slash: '/', colon: ':', semicolon: ';', less: '<', equal: '=', greater: '>', question: '?', at: '@', bracketleft: '[', backslash: '\\', bracketright: ']', asciicircum: '^', underscore: '_', grave: '`', braceleft: '{', bar: '|', braceright: '}', asciitilde: '~', bullet: '•', endash: '–', emdash: '—', Euro: '€', fi: 'fi', fl: 'fl', zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9' };

function unicodeString(bytes) {
  if (bytes.length % 2) throw new Error('PDF ToUnicode 的 UTF-16BE 字符串长度无效。');
  return new TextDecoder('utf-16be', { fatal: true }).decode(bytes);
}

function glyphName(name) {
  if (name.length === 1) return name;
  if (Object.hasOwn(GLYPHS, name)) return GLYPHS[name];
  if (/^uni(?:[0-9a-f]{4})+$/i.test(name)) return unicodeString(Uint8Array.from(name.slice(3).match(/../g), pair => Number.parseInt(pair, 16)));
  if (/^u[0-9a-f]{4,6}$/i.test(name)) {
    const code = Number.parseInt(name.slice(1), 16);
    if (code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff)) return String.fromCodePoint(code);
  }
  return undefined;
}

function parseCmap(bytes, budget) {
  const lexer = new PdfLexer(bytes, 0, budget);
  const mapping = new Map();
  const lengths = new Set();
  let previous;
  function add(source, destination) {
    if (!isString(source) || !isString(destination) || source.bytes.length < 1 || source.bytes.length > 4) throw new Error('PDF ToUnicode 映射无效。');
    lengths.add(source.bytes.length);
    mapping.set(HEX(source.bytes), unicodeString(destination.bytes));
    if (mapping.size > 100000) throw new Error('PDF ToUnicode 映射数量超过当前上限。');
  }
  while (true) {
    lexer.skip();
    if (lexer.position >= bytes.length) break;
    const token = lexer.value(0, false);
    if (['begincodespacerange', 'beginbfchar', 'beginbfrange'].includes(token?.operator)) {
      const count = integer(previous, 'CMap 映射组数量');
      for (let index = 0; index < count; index += 1) {
        const source = lexer.value(0, false);
        const destination = lexer.value(0, false);
        if (token.operator === 'begincodespacerange') {
          if (!isString(source) || !isString(destination) || source.bytes.length !== destination.bytes.length || source.bytes.length < 1 || source.bytes.length > 4) throw new Error('PDF CMap 字符代码范围无效。');
          lengths.add(source.bytes.length);
        } else if (token.operator === 'beginbfchar') add(source, destination);
        else {
          const values = lexer.value(0, false);
          if (!isString(source) || !isString(destination) || source.bytes.length !== destination.bytes.length || source.bytes.length > 4) throw new Error('PDF CMap 连续映射范围无效。');
          const start = bytesInteger(source.bytes), end = bytesInteger(destination.bytes);
          if (end < start || end - start > 65536 || !isString(values) && !Array.isArray(values)) throw new Error('PDF CMap 连续映射超限或无效。');
          for (let code = start; code <= end; code += 1) {
            const key = new Uint8Array(source.bytes.length);
            let remaining = code;
            for (let byte = key.length - 1; byte >= 0; byte -= 1) { key[byte] = remaining & 255; remaining = Math.floor(remaining / 256); }
            let target = values[code - start];
            if (isString(values)) {
              const out = values.bytes.slice();
              let increment = code - start;
              for (let byte = out.length - 1; byte >= 0 && increment; byte -= 1) { const sum = out[byte] + increment; out[byte] = sum & 255; increment = Math.floor(sum / 256); }
              if (increment) throw new Error('PDF CMap Unicode 范围溢出。');
              target = { bytes: out };
            }
            add({ bytes: key }, target);
          }
        }
      }
      const expected = { begincodespacerange: 'endcodespacerange', beginbfchar: 'endbfchar', beginbfrange: 'endbfrange' }[token.operator];
      if (lexer.value(0, false)?.operator !== expected) throw new Error('PDF CMap 映射组未正确结束。');
    }
    previous = token;
  }
  return { mapping, lengths: Array.from(lengths).sort((a, b) => b - a) };
}

export async function loadPdfFont(document, value) {
  const font = await document.resolve(value);
  if (!isDictionary(font)) throw new Error('PDF 字体资源无效。');
  const subtype = asName(font.Subtype);
  const composite = subtype === 'Type0';
  const encoding = await document.resolve(font.Encoding);
  const baseName = asName(font.BaseFont).replace(/^[A-Z]{6}\+/, '');
  const fallbackFamily = /YaHei/i.test(baseName) ? '"Microsoft YaHei", "PingFang SC", sans-serif' : /Courier|Mono/i.test(baseName) ? 'monospace' : /Times|Serif/i.test(baseName) ? 'serif' : /Song|SimSun|Ming/i.test(baseName) ? '"Songti SC", "SimSun", serif' : '"Arial", "PingFang SC", "Microsoft YaHei", sans-serif';
  if (subtype === 'Type3') document.warn('PDF Type3 字体的自定义字形尚未绘制，使用 Unicode 字符替代。');
  const toUnicode = font.ToUnicode ? parseCmap(await document.stream(await document.resolve(font.ToUnicode)), document.budget) : null;
  if (composite && !toUnicode) document.warn('部分复合字体没有可用的 ToUnicode，无法可靠映射的字符显示为替代符。');
  if (composite && /-V$/.test(asName(encoding))) document.warn('PDF 竖排字体尚未排版。');
  let descendant = null;
  if (composite) {
    const descendants = await document.resolve(font.DescendantFonts);
    if (!Array.isArray(descendants) || !descendants.length) throw new Error('PDF Type0 字体没有后代字体。');
    descendant = await document.resolve(descendants[0]);
    if (!isDictionary(descendant)) throw new Error('PDF CID 字体字典无效。');
  }
  const descriptor = await document.resolve((descendant || font).FontDescriptor);
  const defaultWidth = numeric(composite ? descendant.DW : descriptor?.MissingWidth, '默认字形宽度', composite ? 1000 : /Courier/i.test(baseName) ? 600 : 500);
  const widths = new Map();
  if (composite) {
    const list = await document.resolve(descendant.W) || [];
    if (!Array.isArray(list)) throw new Error('PDF CID 字宽表无效。');
    for (let index = 0; index < list.length;) {
      const start = integer(list[index++], 'CID 起始代码', 0xffffffff);
      const next = list[index++];
      if (Array.isArray(next)) next.forEach((width, offset) => widths.set(start + offset, numeric(width, 'CID 字宽')));
      else {
        const end = integer(next, 'CID 结束代码', 0xffffffff);
        const width = numeric(list[index++], 'CID 字宽');
        if (end < start || end - start > 100000) throw new Error('PDF CID 字宽范围过大。');
        for (let code = start; code <= end; code += 1) widths.set(code, width);
      }
      if (widths.size > 100000) throw new Error('PDF 字宽表超过当前上限。');
    }
  } else {
    const values = await document.resolve(font.Widths);
    const initial = integer(font.FirstChar ?? 0, '字体 FirstChar', 255);
    if (values !== undefined && !Array.isArray(values)) throw new Error('PDF 简单字体字宽表无效。');
    (values || []).forEach((width, index) => widths.set(initial + index, numeric(width, '字形宽度')));
    if (!values && !/Courier/i.test(baseName)) document.warn('部分 PDF 字体未提供字宽，文字位置采用近似字宽。');
  }
  const differences = new Map();
  if (isDictionary(encoding) && encoding.Differences) {
    const values = await document.resolve(encoding.Differences);
    if (!Array.isArray(values)) throw new Error('PDF 字体编码差异表无效。');
    let code = 0;
    for (const value of values) {
      if (typeof value === 'number') code = integer(value, '字体编码', 255);
      else {
        const unicode = glyphName(asName(value));
        if (unicode === undefined) document.warn('部分 PDF 字形名称尚未映射，显示为替代符。');
        differences.set(code++, unicode ?? '�');
      }
    }
  }
  const encodingName = isDictionary(encoding) ? asName(encoding.BaseEncoding) : asName(encoding);
  const decoder = encodingName === 'WinAnsiEncoding' ? new TextDecoder('windows-1252') : encodingName === 'MacRomanEncoding' ? new TextDecoder('macintosh') : null;
  if (encodingName && !['WinAnsiEncoding', 'MacRomanEncoding', 'StandardEncoding', 'Identity-H', 'Identity-V'].includes(encodingName)) document.warn(`PDF 字体编码 ${encodingName} 尚未完整支持。`);

  let prepared, cidToGid;
  const embedded = descriptor?.FontFile2 || descriptor?.FontFile3;
  if (embedded && subtype !== 'Type3') {
    try {
      const record = await document.resolve(embedded);
      if (descriptor.FontFile3 && asName(record?.dictionary?.Subtype) !== 'OpenType') throw new Error('该 FontFile3 的 Type1/CFF 字体程序尚未实现。');
      if (composite && !['Identity-H', 'Identity-V'].includes(encodingName)) throw new Error('该嵌入字体的自定义 CID Encoding 尚未实现。');
      prepared = preparePdfFontProgram(await document.stream(record), 128 * 1024 * 1024 - (document.totalDecoded || 0));
      if (composite && asName(descendant.Subtype) === 'CIDFontType2' && descendant.CIDToGIDMap !== undefined) {
        const map = await document.resolve(descendant.CIDToGIDMap);
        if (asName(map) !== 'Identity') {
          cidToGid = await document.stream(map);
          if (cidToGid.length % 2) throw new Error('PDF CIDToGIDMap 字形映射流被截断。');
        }
      }
      document.totalDecoded = (document.totalDecoded || 0) + prepared.bytes.length;
      if (document.totalDecoded > 128 * 1024 * 1024) throw new Error('PDF 嵌入字体重建超过当前 128MB 总量上限。');
    } catch (error) {
      prepared = undefined;
      document.warn(`PDF 嵌入字体无法用于绘制，使用系统字体替代：${error.message}`);
    }
  }
  if (!prepared && !embedded) document.warn('部分 PDF 字体未嵌入可用的字形程序，使用系统字体替代。');
  const flags = descriptor?.Flags || 0;
  const fontName = `${baseName} ${asName(descriptor?.FontName)}`;
  const namedWeight = /(?:Extra|Ultra)[ -]?Bold/i.test(fontName) ? 800 : /Black|Heavy/i.test(fontName) ? 900 : /(?:Semi|Demi)[ -]?Bold/i.test(fontName) ? 600 : /Bold/i.test(fontName) ? 700 : /(?:Extra|Ultra)[ -]?Light/i.test(fontName) ? 200 : /Thin/i.test(fontName) ? 100 : /Light/i.test(fontName) ? 300 : /Medium/i.test(fontName) ? 500 : /Regular|Normal|Book/i.test(fontName) ? 400 : undefined;
  const descriptorWeight = Number.isFinite(descriptor?.FontWeight) && descriptor.FontWeight >= 1 && descriptor.FontWeight <= 1000 ? descriptor.FontWeight : undefined;
  const weight = prepared?.weight ?? descriptorWeight ?? namedWeight ?? (flags & 262144 || descriptor?.StemV >= 120 ? 700 : 400);
  const style = prepared?.italic || flags & 64 || descriptor?.ItalicAngle || /Italic|Oblique/i.test(fontName) ? 'italic' : 'normal';
  const id = `pdf-font-${++nextFontId}`;

  function mappedGlyph(code, text) {
    if (composite && asName(descendant.Subtype) === 'CIDFontType2') {
      if (!cidToGid) return code;
      return code * 2 + 1 < cidToGid.length ? cidToGid[code * 2] * 256 + cidToGid[code * 2 + 1] : 0;
    }
    // Simple TrueType encodings choose a glyph through the original font cmap.
    // ToUnicode remains the semantic text; Differences defines the glyph name.
    const glyphText = differences.get(code) ?? text;
    const scalars = Array.from(glyphText);
    let glyph = scalars.length === 1 ? prepared.glyphIndex(scalars[0].codePointAt(0)) : 0;
    if (!glyph && !composite) glyph = prepared.glyphIndex(0xf000 + code, true) || prepared.glyphIndex(code, true);
    return glyph;
  }

  return {
    id,
    family: prepared ? id : fallbackFamily,
    fallbackFamily,
    weight,
    style,
    ...(prepared ? { program: { bytes: prepared.bytes, format: prepared.format } } : {}),
    decode(bytes) {
      const result = [];
      let cursor = 0;
      while (cursor < bytes.length) {
        let length = composite ? 2 : 1;
        let text;
        if (toUnicode) {
          for (const candidate of toUnicode.lengths) {
            if (cursor + candidate <= bytes.length) {
              const mapped = toUnicode.mapping.get(HEX(bytes.subarray(cursor, cursor + candidate)));
              if (mapped !== undefined) { length = candidate; text = mapped; break; }
            }
          }
          if (text === undefined) length = toUnicode.lengths.find(candidate => cursor + candidate <= bytes.length) || length;
        }
        if (cursor + length > bytes.length) throw new Error('PDF 文本字符代码被截断。');
        const code = bytesInteger(bytes.subarray(cursor, cursor + length));
        if (text === undefined) {
          if (composite) { text = '�'; document.warn('部分 PDF 字符没有可靠的 Unicode 映射，显示为替代符。'); }
          else if (differences.has(code)) text = differences.get(code);
          else if (decoder) text = decoder.decode(bytes.subarray(cursor, cursor + 1));
          else if (code >= 32 && code <= 126 || code === 9 || code === 10 || code === 13) text = String.fromCharCode(code);
          else { text = '�'; document.warn('部分 PDF 字符没有可靠的 Unicode 映射，显示为替代符。'); }
        }
        const glyph = { text, code, width: widths.get(code) ?? defaultWidth };
        if (prepared) {
          const gid = mappedGlyph(code, text);
          if (gid > 0 && gid < prepared.glyphCount) glyph.renderText = prepared.renderText(gid);
          else document.warn('部分 PDF 字符没有可用的嵌入字形，使用系统字体替代。');
        }
        result.push(glyph);
        cursor += length;
      }
      return result;
    },
  };
}
