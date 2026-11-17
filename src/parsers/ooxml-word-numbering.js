import { attr, children, elements, first, numberAttr, relatedPath, warn, xmlPart } from './ooxml-common.js';

function roman(value) {
  if (value < 1 || value > 3999) return String(value);
  let text = '';
  for (const [n, symbol] of [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]) while (value >= n) { text += symbol; value -= n; }
  return text;
}
function chinese(value) {
  if (value < 1 || value > 9999) return String(value);
  const digits = '零一二三四五六七八九'; let text = '', zero = false;
  for (const [base, unit] of [[1000, '千'], [100, '百'], [10, '十'], [1, '']]) {
    const digit = Math.floor(value / base); value %= base;
    if (digit) { if (zero) text += '零'; if (!(base === 10 && digit === 1 && !text)) text += digits[digit]; text += unit; zero = false; }
    else if (text && value) zero = true;
  }
  return text;
}
function label(value, format, warnings) {
  if (format === 'decimal' || !format) return String(value);
  if (format === 'decimalZero') return String(value).padStart(2, '0');
  if (/^(?:lower|upper)Roman$/.test(format)) { const text = roman(value); return format.startsWith('lower') ? text.toLowerCase() : text; }
  if (/^(?:lower|upper)Letter$/.test(format)) {
    let text = ''; for (let n = value; n > 0 && text.length < 8; n = Math.floor((n - 1) / 26)) text = String.fromCharCode(65 + (n - 1) % 26) + text;
    return format.startsWith('lower') ? text.toLowerCase() : text;
  }
  if (['chineseCounting', 'chineseCountingThousand', 'ideographTraditional', 'japaneseCounting'].includes(format)) return chinese(value);
  if (format === 'none') return '';
  if (format !== 'bullet') warn(warnings, `Word 编号格式 ${format} 当前按阿拉伯数字近似显示。`);
  return String(value);
}

export function createWordNumbering(parts, rels, styles, warnings) {
  const path = relatedPath(rels, 'numbering'), document = path ? xmlPart(parts, path) : null;
  const abstracts = new Map(), numbers = new Map(), counters = new Map();
  const level = node => ({
    start: numberAttr(first(node, 'start'), 'val', 1),
    format: attr(first(node, 'numFmt'), 'val') || 'decimal',
    text: attr(first(node, 'lvlText'), 'val'),
    suffix: attr(first(node, 'suff'), 'val') || 'tab',
    align: attr(first(node, 'lvlJc'), 'val') || 'left',
    paragraph: styles.paragraph(first(node, 'pPr')),
    run: styles.run(first(node, 'rPr')),
  });
  for (const node of children(document?.documentElement, 'abstractNum')) {
    const levels = new Map();
    for (const item of children(node, 'lvl')) { const index = numberAttr(item, 'ilvl'); if (index < 0 || index > 8) throw new Error('Word 编号层级无效。'); levels.set(index, level(item)); }
    abstracts.set(attr(node, 'abstractNumId'), levels);
  }
  for (const node of children(document?.documentElement, 'num')) {
    const levels = new Map(abstracts.get(attr(first(node, 'abstractNumId'), 'val')) || []);
    for (const override of children(node, 'lvlOverride')) {
      const index = numberAttr(override, 'ilvl'), own = first(override, 'lvl'), start = first(override, 'startOverride');
      if (own) levels.set(index, level(own));
      if (start && levels.has(index)) levels.set(index, { ...levels.get(index), start: numberAttr(start, 'val', 1) });
    }
    numbers.set(attr(node, 'numId'), levels);
  }
  return {
    next(properties) {
      if (!properties?.numId || properties.numId === '0') return null;
      const levels = numbers.get(properties.numId), index = properties.level || 0, current = levels?.get(index);
      if (!current) { warn(warnings, '部分 Word 编号引用的定义缺失，未生成前缀。'); return null; }
      if (index < 0 || index > 8) throw new Error('Word 编号层级无效。');
      const values = counters.get(properties.numId) || [];
      values[index] = values[index] === undefined ? current.start : values[index] + 1;
      for (let n = index + 1; n < 9; n++) values[n] = undefined;
      counters.set(properties.numId, values);
      const text = current.text.replace(/%([1-9])/g, (_, n) => { const k = Number(n) - 1, definition = levels.get(k) || current; return label(values[k] ?? definition.start, definition.format, warnings); });
      return { numberingLabel: text, numberingSuffix: current.suffix, numberingLevel: index, numberingStyle: current.run, paragraph: current.paragraph };
    }
  };
}
