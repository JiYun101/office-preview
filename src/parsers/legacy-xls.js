import { binary, warn, lowUnicode, utf16 } from './legacy-binary.js';
import { readXlsStyles } from './legacy-xls-style.js';

const MAX_ROWS = 50000;
const MAX_CELLS = 1000000;

function biffRecords(bytes, start) {
  const source = binary(bytes);
  const records = [];
  let cursor = start;
  while (cursor + 4 <= bytes.length) {
    const type = source.u16(cursor);
    const length = source.u16(cursor + 2);
    const data = source.slice(cursor + 4, length);
    records.push({ type, data, offset: cursor });
    if (records.length > MAX_CELLS + 10000) throw new Error('XLS 的记录数量超过当前预览上限。');
    cursor += 4 + length;
    if (type === 0x000a) return records;
  }
  throw new Error('XLS 的 BIFF 记录没有完整的 EOF 结束标记。');
}

class StringSegments {
  constructor(segments) { this.segments = segments; this.segment = 0; this.offset = 0; }
  next() {
    if (++this.segment >= this.segments.length) throw new Error('XLS 共享字符串或续块不完整。');
    this.offset = 0;
  }
  byte() {
    while (this.offset === this.segments[this.segment].length) this.next();
    return this.segments[this.segment][this.offset++];
  }
  u16() { return this.byte() + this.byte() * 256; }
  u32() { return this.u16() + this.u16() * 65536; }
  skip(length) {
    while (length > 0) {
      if (this.offset === this.segments[this.segment].length) this.next();
      const take = Math.min(length, this.segments[this.segment].length - this.offset);
      this.offset += take;
      length -= take;
    }
  }
  string() {
    const characterCount = this.u16();
    const flags = this.byte();
    let highByte = flags & 1;
    const formatRuns = flags & 8 ? this.u16() : 0;
    const phoneticLength = flags & 4 ? this.u32() : 0;
    let text = '';
    let remaining = characterCount;
    while (remaining > 0) {
      if (this.offset === this.segments[this.segment].length) {
        this.next();
        highByte = this.byte() & 1;
      }
      const current = this.segments[this.segment];
      const width = highByte ? 2 : 1;
      const take = Math.min(remaining, Math.floor((current.length - this.offset) / width));
      if (take === 0) throw new Error('XLS 字符串在双字节字符中间断开。');
      const part = current.subarray(this.offset, this.offset + take * width);
      text += highByte ? utf16.decode(part) : lowUnicode(part);
      this.offset += take * width;
      remaining -= take;
    }
    this.skip(formatRuns * 4 + phoneticLength);
    return text;
  }
}

function continuedSegments(records, index) {
  const segments = [records[index].data];
  while (index + 1 < records.length && records[index + 1].type === 0x003c) segments.push(records[++index].data);
  return segments;
}

function rkNumber(value) {
  let result;
  if (value & 2) result = value >> 2;
  else {
    const bytes = new Uint8Array(8);
    const view = new DataView(bytes.buffer);
    view.setUint32(4, value & 0xfffffffc, true);
    result = view.getFloat64(0, true);
  }
  return value & 1 ? result / 100 : result;
}

const excelErrors = { 0: '#NULL!', 7: '#DIV/0!', 15: '#VALUE!', 23: '#REF!', 29: '#NAME?', 36: '#NUM!', 42: '#N/A', 43: '#GETTING_DATA' };

export function parseXls(compound, name) {
  const stream = compound.get('Workbook') || compound.get('Book');
  if (!stream) throw new Error('文件缺少 Workbook 流，不是受支持的 XLS 工作簿。');
  const globals = biffRecords(stream, 0);
  const first = binary(globals[0].data);
  if (globals[0].type !== 0x0809 || first.u16(0) !== 0x0600 || first.u16(2) !== 5) {
    throw new Error('当前 XLS 解析器仅支持 Excel 97–2003 BIFF8 工作簿。');
  }
  if (globals.some(record => record.type === 0x002f)) throw new Error('XLS 工作簿已加密，当前版本不支持解密。');
  const warnings = ['XLS 已读取基本字体、对齐、边框、填色及行列尺寸；图表、绘图、条件格式和打印分页尚未还原。'];
  const formatting = readXlsStyles(globals, warnings);
  const bounds = [];
  let sharedStrings = [];
  for (let index = 0; index < globals.length; index++) {
    const record = globals[index];
    const data = binary(record.data);
    if (record.type === 0x0085) {
      const count = data.u8(6);
      const highByte = data.u8(7) & 1;
      const bytes = data.slice(8, count * (highByte ? 2 : 1));
      bounds.push({ offset: data.u32(0), type: data.u8(5), state: data.u8(4), name: highByte ? utf16.decode(bytes) : lowUnicode(bytes) });
    } else if (record.type === 0x00fc) {
      const strings = new StringSegments(continuedSegments(globals, index));
      strings.u32();
      const count = strings.u32();
      if (count > MAX_CELLS) throw new Error('XLS 共享字符串数量超过当前预览上限。');
      sharedStrings = Array.from({ length: count }, () => strings.string());
    }
  }
  let cellCount = 0;
  const sheets = [];
  for (const bound of bounds) {
    if (bound.type !== 0) { warn(warnings, '非普通工作表（图表页、宏页等）尚未预览。'); continue; }
    if (bound.state !== 0) { warn(warnings, '隐藏工作表未显示。'); continue; }
    const records = biffRecords(stream, bound.offset);
    const sheetBof = binary(records[0].data);
    if (records[0].type !== 0x0809 || sheetBof.u16(0) !== 0x0600 || sheetBof.u16(2) !== 0x0010) {
      throw new Error('XLS 的 BoundSheet8 没有指向有效的工作表。');
    }
    const rows = [];
    const merges = [];
    const columnWidths = [], columnWidthUnits = [], columnStyles = [], rowHeights = [], rowStyles = [];
    let defaultRowHeight = 20, defaultColumnWidthUnits = 8 * 256, standardColumnWidth;
    for (const record of records) {
      const data = binary(record.data);
      if (record.type === 0x0225) {
        const flags = data.u16(0), height = data.u16(2) & 0x7fff;
        if (height) defaultRowHeight = height / 15;
        if (flags & 2) warn(warnings, '隐藏行和列目前仍可在预览中查看。');
      } else if (record.type === 0x0055) defaultColumnWidthUnits = data.u16(0) * 256;
      else if (record.type === 0x0099) standardColumnWidth = data.u16(0);
      else if (record.type === 0x007d) {
        data.need(0, 10);
        const start = data.u16(0), end = data.u16(2), units = data.u16(4), xf = data.u16(6);
        if (end < start) throw new Error('XLS 列格式范围无效。');
        for (let column = start; column <= Math.min(end, 1023); column++) {
          if (units > 0) columnWidthUnits[column] = units;
          columnStyles[column] = formatting.style(xf);
        }
        if (data.u16(8) & 1) warn(warnings, '隐藏行和列目前仍可在预览中查看。');
      } else if (record.type === 0x0208) {
        data.need(0, 16);
        const row = data.u16(0), height = data.u16(6) & 0x7fff, flags = data.u16(12);
        if (row >= MAX_ROWS) { warn(warnings, '超出当前行数上限的行格式未显示。'); continue; }
        if (height) rowHeights[row] = height / 15;
        if (flags & 128) rowStyles[row] = formatting.style(data.u16(14) & 0x0fff);
        if (flags & 32) warn(warnings, '隐藏行和列目前仍可在预览中查看。');
      }
    }
    if (standardColumnWidth !== undefined) defaultColumnWidthUnits = standardColumnWidth;
    const digitWidth = (formatting.defaultFont.fontSize || 14) * 0.5;
    for (let column = 0; column < columnWidthUnits.length; column++) {
      if (columnWidthUnits[column] !== undefined) columnWidths[column] = columnWidthUnits[column] / 256 * digitWidth + 5;
    }
    const defaultColumnWidth = defaultColumnWidthUnits / 256 * digitWidth + 5;
    let pendingString = null;
    const put = (row, column, value, xf) => {
      if (row >= MAX_ROWS || column >= 1024) throw new Error('工作表超过当前预览上限：50000 行、1024 列。');
      if (typeof value === 'number' && !Number.isFinite(value)) throw new Error('XLS 单元格包含无效数值。');
      rows[row] ||= [];
      if (!rows[row][column] && ++cellCount > MAX_CELLS) throw new Error('工作簿超过当前预览上限：1000000 个有数据单元格。');
      const style = xf === undefined ? rowStyles[row] || columnStyles[column] || formatting.defaultStyle : formatting.style(xf);
      rows[row][column] = style ? { value, style } : { value };
    };
    for (let index = 1; index < records.length; index++) {
      const record = records[index];
      const data = binary(record.data);
      if (record.type === 0x002f) throw new Error('XLS 工作簿已加密，当前版本不支持解密。');
      if (record.type === 0x0203) put(data.u16(0), data.u16(2), data.f64(6), data.u16(4));
      else if (record.type === 0x00fd) {
        const stringIndex = data.u32(6);
        if (stringIndex >= sharedStrings.length) throw new Error('XLS 的共享字符串索引越界。');
        put(data.u16(0), data.u16(2), sharedStrings[stringIndex], data.u16(4));
      } else if (record.type === 0x027e) put(data.u16(0), data.u16(2), rkNumber(data.u32(6)), data.u16(4));
      else if (record.type === 0x00bd) {
        const count = (record.data.length - 6) / 6;
        if (!Number.isInteger(count) || count < 1 || data.u16(record.data.length - 2) !== data.u16(2) + count - 1) {
          throw new Error('XLS 的 MulRK 连续数值记录长度无效。');
        }
        for (let item = 0; item < count; item++) put(data.u16(0), data.u16(2) + item, rkNumber(data.u32(6 + item * 6)), data.u16(4 + item * 6));
      } else if (record.type === 0x0205) {
        put(data.u16(0), data.u16(2), data.u8(7) ? excelErrors[data.u8(6)] || '#ERROR!' : !!data.u8(6), data.u16(4));
      } else if (record.type === 0x0006) {
        const row = data.u16(0);
        const column = data.u16(2);
        const xf = data.u16(4);
        warn(warnings, '公式仅显示文件保存时的缓存结果，不执行公式重算，缓存值可能已过期。');
        if (data.u16(12) !== 0xffff) put(row, column, data.f64(6), xf);
        else {
          const type = data.u8(6);
          if (type === 0) { pendingString = { row, column, xf }; put(row, column, '', xf); }
          else if (type === 1) put(row, column, !!data.u8(8), xf);
          else if (type === 2) put(row, column, excelErrors[data.u8(8)] || '#ERROR!', xf);
          else if (type === 3) put(row, column, '', xf);
          else throw new Error('XLS 公式缓存值类型无效。');
        }
      } else if (record.type === 0x0207 && pendingString) {
        put(pendingString.row, pendingString.column, new StringSegments(continuedSegments(records, index)).string(), pendingString.xf);
        pendingString = null;
      } else if (record.type === 0x0204) {
        // BIFF8 Label (less common than LabelSst) stores a Unicode string.
        put(data.u16(0), data.u16(2), new StringSegments([data.slice(6, record.data.length - 6)]).string(), data.u16(4));
      } else if (record.type === 0x0201) {
        put(data.u16(0), data.u16(2), '', data.u16(4));
      } else if (record.type === 0x00be) {
        const count = (record.data.length - 6) / 2;
        if (!Number.isInteger(count) || count < 1 || data.u16(record.data.length - 2) !== data.u16(2) + count - 1) throw new Error('XLS 的 MulBlank 格式记录长度无效。');
        for (let item = 0; item < count; item++) put(data.u16(0), data.u16(2) + item, '', data.u16(4 + item * 2));
      } else if (record.type === 0x00e5) {
        const count = data.u16(0);
        data.need(2, count * 8);
        for (let item = 0; item < count; item++) {
          const offset = 2 + item * 8;
          const merge = { startRow: data.u16(offset), endRow: data.u16(offset + 2), startCol: data.u16(offset + 4), endCol: data.u16(offset + 6) };
          if (merge.endRow < merge.startRow || merge.endCol < merge.startCol || merge.endRow >= MAX_ROWS || merge.endCol >= 1024) {
            throw new Error('XLS 合并区域超出当前预览范围或坐标无效。');
          }
          merges.push(merge);
        }
      }
    }
    if (pendingString) warn(warnings, '部分公式缺少文本缓存，已显示为空单元格。');
    sheets.push({ name: bound.name, rows, merges, columnWidths, columnWidthUnits, rowHeights,
      defaultRowHeight, defaultColumnWidth, defaultColumnWidthUnits,
      defaultFont: formatting.defaultFont, defaultStyle: formatting.defaultStyle,
      rowStyles, columnStyles, date1904: formatting.date1904 });
  }
  if (!bounds.length) throw new Error('XLS 中没有可解析的 BoundSheet8 工作表目录。');
  if (!sheets.length) throw new Error('该 XLS 没有当前版本可显示的普通可见工作表。');
  return { kind: 'spreadsheet', format: 'xls', name, warnings, sheets };
}

