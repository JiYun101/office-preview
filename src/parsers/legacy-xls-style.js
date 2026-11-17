import { binary, warn, lowUnicode, utf16 } from './legacy-binary.js';

// MS-XLS FONT, XF/CellXF, FORMAT and Icv tables. Values remain independent of DOM.
const DEFAULT_PALETTE = [
  '000000','ffffff','ff0000','00ff00','0000ff','ffff00','ff00ff','00ffff',
  '800000','008000','000080','808000','800080','008080','c0c0c0','808080',
  '9999ff','993366','ffffcc','ccffff','660066','ff8080','0066cc','ccccff',
  '000080','ff00ff','ffff00','00ffff','800080','800000','008080','0000ff',
  '00ccff','ccffff','ccffcc','ffff99','99ccff','ff99cc','cc99ff','ffcc99',
  '3366ff','33cccc','99cc00','ffcc00','ff9900','ff6600','666699','969696',
  '003366','339966','003300','333300','993300','993366','333399','333333'
];
const BUILTIN_FORMATS = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
  5: '"$"#,##0_);("$"#,##0)', 6: '"$"#,##0_);[Red]("$"#,##0)',
  7: '"$"#,##0.00_);("$"#,##0.00)', 8: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'm/d/yy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM',
  19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'm/d/yy h:mm',
  27: 'yyyy"年"m"月"', 28: 'm"月"d"日"', 29: 'm"月"d"日"', 30: 'm/d/yy',
  31: 'yyyy"年"m"月"d"日"', 32: 'h"时"mm"分"', 33: 'h"时"mm"分"ss"秒"',
  34: 'AM/PM h"时"mm"分"', 35: 'AM/PM h"时"mm"分"ss"秒"', 36: 'yyyy"年"m"月"',
  37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 48: '##0.0E+0', 49: '@',
  50: 'yyyy"年"m"月"', 51: 'm"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"',
  54: 'm"月"d"日"', 55: 'AM/PM h"时"mm"分"', 56: 'AM/PM h"时"mm"分"ss"秒"', 57: 'yyyy"年"m"月"', 58: 'm"月"d"日"'
};
const BORDER_STYLES = {
  1: ['solid',1], 2: ['solid',2], 3: ['dashed',1], 4: ['dotted',1],
  5: ['solid',3], 6: ['double',3], 7: ['solid',0.5], 8: ['dashed',2],
  9: ['dashed',1], 10: ['dashed',2], 11: ['dashed',1], 12: ['dashed',2], 13: ['dashed',2]
};

function unicodeString(data, offset, short = false) {
  const count = short ? data.u8(offset) : data.u16(offset);
  const flagsOffset = offset + (short ? 1 : 2);
  const flags = data.u8(flagsOffset);
  const bytes = data.slice(flagsOffset + 1, count * (flags & 1 ? 2 : 1));
  return flags & 1 ? utf16.decode(bytes) : lowUnicode(bytes);
}

export function readXlsStyles(records, warnings) {
  const fonts = [], xfs = [], palette = [...DEFAULT_PALETTE], formats = { ...BUILTIN_FORMATS };
  let date1904 = false;
  for (const record of records) {
    const data = binary(record.data);
    if (record.type === 0x0031) {
      data.need(0, 16);
      const flags = data.u16(2), script = data.u16(8);
      const name = unicodeString(data, 14, true);
      if (flags & 0xf0 || script) warn(warnings, '部分字体的轮廓、阴影、压缩或上下标效果尚未还原。');
      fonts.push({
        fontFamily: name.replace(/^@/, ''), fontSize: data.u16(0) / 15,
        bold: data.u16(6) >= 600, italic: !!(flags & 2), strike: !!(flags & 8), underline: data.u8(10) !== 0,
        colorIndex: data.u16(4)
      });
      if (name.startsWith('@')) warn(warnings, '部分竖排字体按普通字体显示。');
    } else if (record.type === 0x00e0) {
      data.need(0, 20);
      xfs.push(record.data);
    } else if (record.type === 0x0092) {
      const count = data.u16(0);
      if (count > 56) throw new Error('XLS 调色板颜色数量无效。');
      data.need(2, count * 4);
      for (let i = 0; i < count; i++) palette[i] = [...data.slice(2 + i * 4, 3)].map(value => value.toString(16).padStart(2, '0')).join('');
    } else if (record.type === 0x041e) {
      formats[data.u16(0)] = unicodeString(data, 2);
    } else if (record.type === 0x0022) date1904 = data.u16(0) === 1;
  }
  const color = (index, background = false) => {
    if (index === 65) return '#ffffff';
    if (index === 64 || index === 0x7fff) return '#000000';
    if (index < 8) return '#' + DEFAULT_PALETTE[index];
    if (index >= 8 && index < 64) return '#' + palette[index - 8];
    warn(warnings, '部分扩展颜色尚未还原，已使用默认前景/背景颜色。');
    return background ? '#ffffff' : '#000000';
  };
  const font = index => {
    // FontIndex intentionally has no item 4; the fifth FONT is index 5.
    const item = index === 4 ? undefined : fonts[index > 4 ? index - 1 : index];
    if (!item) {
      if (fonts.length) warn(warnings, '部分单元格引用了缺失字体，已使用工作簿默认字体。');
      const fallback = fonts[0];
      if (!fallback) return {};
      const { colorIndex, ...style } = fallback;
      return { ...style, color: color(colorIndex) };
    }
    const { colorIndex, ...style } = item;
    return { ...style, color: color(colorIndex) };
  };
  const border = (kind, index) => {
    if (!kind) return undefined;
    const definition = BORDER_STYLES[kind];
    if (!definition) { warn(warnings, '部分边框线型尚未还原。'); return undefined; }
    return { style: definition[0], width: definition[1], color: color(index) };
  };
  const raw = xfs.map(bytes => {
    const data = binary(bytes), flags = data.u16(4), borders1 = data.u32(10), borders2 = data.u32(14), fill = data.u16(18);
    const alignCode = data.u8(6) & 7, verticalCode = (data.u8(6) >>> 4) & 7;
    const align = { 1: 'left', 2: 'center', 3: 'right', 4: 'left', 5: 'justify', 6: 'center', 7: 'justify' }[alignCode];
    if ([4,6,7].includes(alignCode)) warn(warnings, '填充、跨列居中及分散对齐当前以近似对齐方式显示。');
    const borders = {};
    for (const [side, kind, index] of [
      ['left', borders1 & 15, (borders1 >>> 16) & 127], ['right', (borders1 >>> 4) & 15, (borders1 >>> 23) & 127],
      ['top', (borders1 >>> 8) & 15, borders2 & 127], ['bottom', (borders1 >>> 12) & 15, (borders2 >>> 7) & 127]
    ]) {
      const value = border(kind, index);
      if (value) borders[side] = value;
    }
    if (borders1 >>> 30) warn(warnings, '单元格对角线边框尚未绘制。');
    if (borders2 & 0x02000000) warn(warnings, '部分 Office 扩展单元格格式尚未解析。');
    const pattern = borders2 >>> 26;
    if (pattern > 1) warn(warnings, '非纯色的单元格填充图案尚未还原。');
    const rotation = data.u8(7);
    if (rotation === 255) warn(warnings, '堆叠文字当前按普通横排显示。');
    return {
      isStyle: !!(flags & 4), parent: flags >>> 4, attributes: data.u8(9),
      font: font(data.u16(0)), number: { format: formats[data.u16(2)] || 'General' },
      alignment: {
        ...(align ? { align } : {}), verticalAlign: { 0: 'top', 1: 'middle', 2: 'bottom', 3: 'middle', 4: 'middle' }[verticalCode] || 'bottom',
        wrapText: !!(data.u8(6) & 8), indent: data.u8(8) & 15,
        ...(rotation > 0 && rotation <= 180 ? { rotation: rotation <= 90 ? rotation : 90 - rotation } : {})
      },
      borders: { borders }, fill: { background: pattern === 1 ? color(fill & 127, true) : 'transparent' }
    };
  });
  const resolvedGroups = new Map(), styles = new Map();
  const groups = (index, seen = new Set()) => {
    if (resolvedGroups.has(index)) return resolvedGroups.get(index);
    if (seen.has(index) || seen.size > 64) throw new Error('XLS 单元格格式继承包含循环或层数过深。');
    const item = raw[index];
    if (!item) return null;
    const result = { font: item.font, number: item.number, alignment: item.alignment, borders: item.borders, fill: item.fill };
    if (!item.isStyle && item.parent !== 0xfff && item.parent !== index && raw[item.parent]?.isStyle) {
      seen.add(index);
      const parent = groups(item.parent, seen);
      for (const [name, mask] of [['number',4],['font',8],['alignment',16],['borders',32],['fill',64]]) {
        if (!(item.attributes & mask)) result[name] = parent[name];
      }
      seen.delete(index);
    }
    resolvedGroups.set(index, result); return result;
  };
  const style = index => {
    if (!xfs.length) return undefined;
    if (styles.has(index)) return styles.get(index);
    const definition = groups(index);
    if (!definition) { warn(warnings, '部分单元格格式索引无效，已使用工作簿默认格式。'); return style(defaultIndex); }
    const result = Object.assign({}, definition.font, definition.number, definition.alignment, definition.borders, definition.fill);
    styles.set(index, result); return result;
  };
  const firstCell = raw.findIndex(item => !item.isStyle);
  const defaultIndex = firstCell >= 0 ? firstCell : 0;
  return { style, defaultIndex, defaultStyle: style(defaultIndex), defaultFont: font(0), date1904 };
}
