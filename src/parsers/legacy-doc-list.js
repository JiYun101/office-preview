import { binary, utf16, warn } from './legacy-binary.js';
import { readDocPrls, applyDocPrls } from './legacy-doc-style.js';

function chinese(number) {
  if (number === 0) return '零';
  if (number < 0 || number > 9999) return String(number);
  const digits = '零一二三四五六七八九', units = ['', '十', '百', '千'];
  let text = '', zero = false;
  for (let power = 3; power >= 0; power--) {
    const place = 10 ** power, digit = Math.floor(number / place) % 10;
    if (!digit) { if (text && number % place) zero = true; continue; }
    if (zero) text += digits[0]; zero = false;
    text += (digit === 1 && power === 1 && !text ? '' : digits[digit]) + units[power];
  }
  return text;
}
function formatted(number, kind, warnings) {
  if (kind === 0) return String(number);
  if (kind === 1 || kind === 2) {
    let value = number, text = '';
    for (const [n, symbol] of [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']]) while (value >= n && text.length < 100) { text += symbol; value -= n; }
    return kind === 2 ? text.toLowerCase() : text;
  }
  if (kind === 3 || kind === 4) { let value = number, text = ''; while (value > 0) { value--; text = String.fromCharCode(65 + value % 26) + text; value = Math.floor(value / 26); } return kind === 4 ? text.toLowerCase() : text; }
  if ([11, 14, 19, 39].includes(kind)) return chinese(number);
  if (kind === 22) return String(number).padStart(2, '0');
  warn(warnings, '部分 DOC 编号格式暂按阿拉伯数字显示。'); return String(number);
}

export function createDocLists(tableBytes, pair, pairCount, warnings) {
  const data = binary(tableBytes), definitions = new Map(), overrides = [], states = new Map();
  const part = pair(73), overridePart = pair(74);
  let levelBoundary = tableBytes.length;
  if (part.length) {
    for (let i = 0; i < pairCount; i++) { const other = pair(i); if (other.length && other.offset > part.offset) levelBoundary = Math.min(levelBoundary, other.offset); }
  }
  const readLevel = (position, boundary) => {
    data.need(position, 28);
    if (position + 28 > boundary) throw new Error('DOC 列表层级头不完整。');
    const header = position, startAt = data.u32(position), kind = data.u8(position + 4), flags = data.u8(position + 5), chpx = data.u8(position + 24), papx = data.u8(position + 25);
    if (startAt > 32767) throw new Error('DOC 编号起始值无效。');
    position += 28;
    const paragraph = readDocPrls(data.slice(position, papx)); position += papx;
    const character = readDocPrls(data.slice(position, chpx)); position += chpx;
    const length = data.u16(position); position += 2;
    if (length > 1000 || position + length * 2 > boundary) throw new Error('DOC 列表编号文字超出层级范围。');
    const text = utf16.decode(data.slice(position, length * 2)); position += length * 2;
    return { end: position, startAt, kind, flags, follow: data.u8(header + 15), restart: data.u8(header + 26), paragraph, character, text, placeholders: Array.from(data.slice(header + 6, 9)).filter(value => value > 0).map(value => value - 1) };
  };
  if (part.length) {
    data.need(part.offset, part.length);
    const count = data.u16(part.offset);
    if (count > 4096 || part.length !== 2 + count * 28) throw new Error('DOC 列表目录长度无效。');
    let position = part.offset + part.length;
    for (let i = 0; i < count; i++) {
      const offset = part.offset + 2 + i * 28, id = data.u32(offset), levels = [], simple = !!(data.u8(offset + 26) & 1);
      for (let j = 0; j < (simple ? 1 : 9); j++) { const level = readLevel(position, levelBoundary); position = level.end; levels.push(level); }
      definitions.set(id, { levels, linkedStyles: Array.from({ length: 9 }, (_, j) => data.u16(offset + 8 + j * 2)) });
    }
  }
  if (overridePart.length) {
    data.need(overridePart.offset, overridePart.length);
    const count = data.u32(overridePart.offset), boundary = overridePart.offset + overridePart.length;
    if (count > 4096 || 4 + count * 16 > overridePart.length) throw new Error('DOC 列表覆盖目录长度无效。');
    let position = overridePart.offset + 4 + count * 16;
    for (let i = 0; i < count; i++) {
      const header = overridePart.offset + 4 + i * 16, entry = { id: data.u32(header), levels: new Map() }, levelCount = data.u8(header + 12);
      if (levelCount > 9 || position + 4 > boundary) throw new Error('DOC 列表覆盖层级数量无效。');
      position += 4;
      for (let j = 0; j < levelCount; j++) {
        if (position + 8 > boundary) throw new Error('DOC 列表覆盖层级不完整。');
        const startAt = data.u32(position), flags = data.u32(position + 4), index = flags & 15;
        if (index > 8) throw new Error('DOC 列表覆盖层级无效。');
        position += 8;
        let level;
        if (flags & 32) { level = readLevel(position, boundary); position = level.end; }
        entry.levels.set(index, { level, ...(flags & 16 && !(flags & 32) ? { startAt } : {}) });
      }
      overrides.push(entry);
    }
  }
  return {
    apply(paragraph, character) {
      const id = paragraph.paragraph._listId, levelIndex = paragraph.paragraph._listLevel || 0;
      if (!id) return null;
      const override = overrides[id - 1], definition = definitions.get(override?.id), change = override?.levels.get(levelIndex);
      const level = change?.level || definition?.levels[levelIndex];
      if (!level || levelIndex > 8) { warn(warnings, '部分 DOC 列表缺少对应定义，未生成编号。'); return null; }
      if (definition.linkedStyles[levelIndex] !== 0x0fff) warn(warnings, 'DOC 列表关联样式尚未完整还原。');
      let state = states.get(id);
      if (!state) { state = []; states.set(id, state); }
      state[levelIndex] = state[levelIndex] === undefined ? change?.startAt ?? level.startAt : state[levelIndex] + 1;
      for (let i = levelIndex + 1; i < 9; i++) if (!(definition.levels[i]?.flags & 8) || levelIndex < (definition.levels[i]?.restart || 0)) state[i] = undefined;
      const replacements = new Set(level.placeholders);
      let text = '';
      for (let i = 0; i < level.text.length; i++) {
        const ch = level.text.charCodeAt(i);
        if (level.kind === 23) text += String.fromCharCode(ch & 0x0fff);
        else if (replacements.has(i)) text += formatted(state[ch] ?? definition.levels[ch]?.startAt ?? 1, level.flags & 4 ? 0 : definition.levels[ch]?.kind ?? level.kind, warnings);
        else text += level.text[i];
      }
      return {
        paragraph: applyDocPrls(applyDocPrls(paragraph.baseParagraph, level.paragraph), paragraph.direct),
        character: applyDocPrls(character, level.character),
        text, suffix: ['tab', 'space', 'nothing'][level.follow] || 'nothing', align: ['left', 'center', 'right'][level.flags & 3] || 'left'
      };
    }
  };
}
