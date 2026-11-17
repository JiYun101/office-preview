import { binary, lowUnicode, utf16, warn } from './legacy-binary.js';

const alignments = ['left', 'center', 'right', 'justify', 'justify', 'center', 'justify'];
const signed16 = value => value & 0x8000 ? value - 65536 : value;

function cursorFor(bytes) {
  const source = binary(bytes);
  let position = 0;
  return {
    u16() { const value = source.u16(position); position += 2; return value; },
    u32() { const value = source.u32(position); position += 4; return value; },
    skip(length) { source.need(position, length); position += length; },
    get position() { return position; }, get length() { return bytes.length; }
  };
}

export function createPptTextReader({ source, children, warnings, color, fonts }) {
  const clean = text => text.replace(/\r/g, '\n').replace(/[\x00-\x08\x0e-\x1f]/g, '');
  const textFromRecord = record => {
    if (record.type === 0x0fa0) {
      if (record.length % 2) throw new Error('PPT Unicode 文本记录长度无效。');
      return utf16.decode(source.slice(record.data, record.length));
    }
    return record.type === 0x0fa8 ? lowUnicode(source.slice(record.data, record.length)) : null;
  };

  function paragraphStyle(reader, scheme) {
    const mask = reader.u32(), style = {};
    if (mask & 0x0f) {
      const flags = reader.u16();
      if (mask & 1) style.hasBullet = !!(flags & 1);
    }
    if (mask & 0x80) style.bulletChar = String.fromCharCode(reader.u16());
    if (mask & 0x10) reader.skip(2);
    if (mask & 0x40) reader.skip(2);
    if (mask & 0x20) reader.skip(4);
    if (mask & 0x800) style.align = alignments[reader.u16()] || 'left';
    if (mask & 0x1000) { const value = signed16(reader.u16()); if (value >= 0) style.lineHeight = value / 100; else style.lineHeightPx = -value / 6; }
    if (mask & 0x2000) { const value = signed16(reader.u16()); if (value >= 0) style.spaceBeforeLines = value / 100; else style.spaceBefore = -value / 6; }
    if (mask & 0x4000) { const value = signed16(reader.u16()); if (value >= 0) style.spaceAfterLines = value / 100; else style.spaceAfter = -value / 6; }
    if (mask & 0x100) style.marginLeft = signed16(reader.u16()) / 6;
    if (mask & 0x400) style.indent = signed16(reader.u16()) / 6;
    if (mask & 0x8000) reader.skip(2);
    if (mask & 0x100000) { const count = reader.u16(); if (count > 1000) throw new Error('PPT 文本制表位数量过多。'); reader.skip(count * 4); }
    if (mask & 0x10000) reader.skip(2);
    if (mask & 0xe0000) reader.skip(2);
    if (mask & 0x200000) { if (reader.u16()) warn(warnings, '部分 PPT 竖排或双向文本尚未完整排版。'); }
    if (mask & 0xffc00000) throw new Error('尚未支持的 PPT 段落扩展格式。');
    return style;
  }

  function characterStyle(reader, scheme) {
    const mask = reader.u32(), style = {};
    if (mask & 0xffff) {
      const flags = reader.u16();
      if (mask & 1) style.bold = !!(flags & 1);
      if (mask & 2) style.italic = !!(flags & 2);
      if (mask & 4) style.underline = !!(flags & 4);
    }
    if (mask & 0x10000) { const index = reader.u16(); if (fonts[index]) style.fontFamily = fonts[index]; }
    if (mask & 0x200000) { const index = reader.u16(); if (fonts[index]) style.fontFamily = fonts[index]; }
    if (mask & 0x400000) reader.skip(2);
    if (mask & 0x800000) reader.skip(2);
    if (mask & 0x20000) {
      const size = signed16(reader.u16());
      if (size <= 0 || size > 4000) throw new Error('PPT 字号无效。');
      style.fontSize = size * 96 / 72;
    }
    if (mask & 0x40000) style.color = color(reader.u32(), scheme, true);
    if (mask & 0x80000) reader.skip(2);
    if (mask & 0xff100000) throw new Error('尚未支持的 PPT 字符扩展格式。');
    return style;
  }

  function readStyles(record, text, scheme) {
    const reader = cursorFor(source.slice(record.data, record.length));
    const paragraphs = [], characters = [];
    const length = text.length + 1;
    let position = 0;
    while (position < length) {
      const count = reader.u32(), level = reader.u16();
      if (!count || count > length - position || level > 8) throw new Error('PPT 段落样式范围无效。');
      paragraphs.push({ start: position, end: position + count, level, ...paragraphStyle(reader, scheme) });
      position += count;
    }
    position = 0;
    while (position < length) {
      const count = reader.u32();
      if (!count || count > length - position) throw new Error('PPT 字符样式范围无效。');
      characters.push({ start: position, end: position + count, ...characterStyle(reader, scheme) });
      position += count;
    }
    return { paragraphs, characters };
  }

  function readBlocks(records) {
    const blocks = [];
    let block;
    for (const record of records) {
      if (record.type === 0x0f9f) {
        if (record.length < 4) throw new Error('PPT TextHeaderAtom 长度无效。');
        block = { text: '', textType: source.u32(record.data), records: [], placeholderIndex: record.instance };
        blocks.push(block);
      } else {
        const text = textFromRecord(record);
        if (text !== null && !block) { block = { text: '', textType: 4, records: [] }; blocks.push(block); }
        if (block) { block.records.push(record); if (text !== null) block.text += text; }
      }
    }
    return blocks;
  }

  function masterStyles(container, scheme) {
    const result = new Map();
    for (const record of children(container).filter(record => record.type === 0x0fa3)) {
      try {
        const reader = cursorFor(source.slice(record.data, record.length));
        const count = reader.u16();
        if (count > 9) throw new Error('PPT 母版样式级数无效。');
        const levels = [];
        for (let i = 0; i < count; i++) {
          const level = record.instance >= 5 ? reader.u16() : i;
          if (level >= 9) throw new Error('PPT 母版样式层级无效。');
          levels[level] = { paragraph: paragraphStyle(reader, scheme), character: characterStyle(reader, scheme) };
        }
        result.set(record.instance, levels);
      } catch { warn(warnings, '部分 PPT 母版文字扩展样式尚未还原。'); }
    }
    return result;
  }

  function formatBlock(block, scheme, defaults = new Map()) {
    let styles = { paragraphs: [], characters: [] };
    const record = block.records.find(record => record.type === 0x0fa1);
    if (record) {
      try { styles = readStyles(record, block.text, scheme); }
      catch { warn(warnings, '部分 PPT 文字扩展样式或样式范围尚未完整还原。'); }
    }
    const paragraphs = [];
    let start = 0;
    for (const text of block.text.split('\r')) {
      const end = start + text.length;
      const paragraph = styles.paragraphs.find(style => style.start <= start && style.end > start) || {};
      const inherited = defaults.get(block.textType)?.[paragraph.level || 0] || defaults.get(4)?.[0] || {};
      const runs = [];
      const stops = new Set([start, end]);
      for (const style of styles.characters) { if (style.start > start && style.start < end) stops.add(style.start); if (style.end > start && style.end < end) stops.add(style.end); }
      const positions = [...stops].sort((a, b) => a - b);
      for (let i = 0; i + 1 < positions.length; i++) {
        const from = positions[i], to = positions[i + 1];
        const own = styles.characters.find(style => style.start <= from && style.end > from) || {};
        const { start: unusedStart, end: unusedEnd, ...character } = own;
        runs.push({ ...inherited.character, ...character, text: clean(block.text.slice(from, to)) });
      }
      const merged = { ...inherited.paragraph, ...paragraph };
      if (merged.hasBullet) runs.unshift({ ...inherited.character, text: `${merged.bulletChar || '•'} ` });
      const linePixels = merged.lineHeightPx || Math.max(1, ...runs.map(run => run.fontSize || 24)) * (merged.lineHeight || 1.2);
      const spaceBefore = merged.spaceBefore ?? (merged.spaceBeforeLines || 0) * linePixels;
      const spaceAfter = merged.spaceAfter ?? (merged.spaceAfterLines || 0) * linePixels;
      paragraphs.push({ runs, ...(merged.align ? { align: merged.align } : {}), ...(merged.lineHeight ? { lineHeight: merged.lineHeight } : {}), ...(merged.lineHeightPx ? { lineHeightPx: merged.lineHeightPx } : {}), ...(spaceBefore ? { spaceBefore } : {}), ...(spaceAfter ? { spaceAfter } : {}) });
      start = end + 1;
    }
    const runs = [];
    paragraphs.forEach((paragraph, index) => { if (index) runs.push({ text: '\n' }); runs.push(...paragraph.runs); });
    return { paragraphs, runs, align: paragraphs[0]?.align || 'left' };
  }

  function readBox(record, outline, scheme, defaults) {
    const records = children(record);
    const reference = records.find(child => child.type === 0x0f9e);
    if (reference) {
      if (reference.length < 4) throw new Error('PPT 文字引用记录不完整。');
      const block = outline[source.u32(reference.data)];
      if (!block) { warn(warnings, '部分 PPT 文本框引用了缺少的外部文字块。'); return null; }
      block.used = true;
      return formatBlock(block, scheme, defaults);
    }
    const blocks = readBlocks(records);
    const formatted = blocks.map(block => formatBlock(block, scheme, defaults));
    const paragraphs = formatted.flatMap(item => item.paragraphs);
    const runs = [];
    formatted.forEach((item, index) => { if (index) runs.push({ text: '\n' }); runs.push(...item.runs); });
    return paragraphs.length ? { runs, paragraphs, align: paragraphs[0].align || 'left' } : null;
  }

  return { textFromRecord, readBlocks, readBox, masterStyles, formatBlock, clean };
}
