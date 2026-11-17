import { PdfLexer, asName, byteString, isDictionary, isReference, integer, numeric } from './pdf-syntax.js';
const MAX_DECODED = 128 * 1024 * 1024;

function ascii85(bytes, limit) {
  let group = [];
  const source = byteString(bytes).replace(/[\x00\t\n\f\r ]/g, '').replace(/^<~/, '');
  const end = source.indexOf('~>');
  if (end < 0 || end !== source.length - 2) throw new Error('PDF ASCII85 流结束标记无效。');
  let expected = 0, digits = 0;
  for (let index = 0; index < end; index += 1) {
    const code = source.charCodeAt(index);
    if (code === 122) {
      if (digits) throw new Error('PDF ASCII85 零缩写位置无效。');
      expected += 4;
    } else {
      if (code < 33 || code > 117) throw new Error('PDF ASCII85 字符无效。');
      digits += 1;
      if (digits === 5) { expected += 4; digits = 0; }
    }
    if (expected > limit) throw new Error('PDF ASCII85 解码超过当前 128MB 总量上限。');
  }
  if (digits === 1) throw new Error('PDF ASCII85 尾组无效。');
  expected += digits ? digits - 1 : 0;
  if (expected > limit) throw new Error('PDF ASCII85 解码超过当前 128MB 总量上限。');
  const output = new Uint8Array(expected);
  let cursor = 0;
  function flush(partial = false) {
    const length = group.length;
    while (group.length < 5) group.push(84);
    let value = 0;
    for (const digit of group) value = value * 85 + digit;
    if (value > 0xffffffff) throw new Error('PDF ASCII85 数值无效。');
    const count = partial ? length - 1 : 4;
    for (let index = 0; index < count; index += 1) output[cursor++] = (value >>> ((3 - index) * 8)) & 255;
    group = [];
  }
  for (const character of source.slice(0, end)) {
    if (character === 'z') {
      if (group.length) throw new Error('PDF ASCII85 零缩写位置无效。');
      output.fill(0, cursor, cursor + 4); cursor += 4;
    } else {
      const code = character.charCodeAt(0) - 33;
      if (code < 0 || code > 84) throw new Error('PDF ASCII85 字符无效。');
      group.push(code);
      if (group.length === 5) flush();
    }
  }
  if (group.length === 1) throw new Error('PDF ASCII85 尾组无效。');
  if (group.length) flush(true);
  return output;
}

function unPredict(bytes, parameters) {
  const predictor = numeric(parameters?.Predictor, 'Predictor', 1);
  if (predictor === 1) return bytes;
  const colors = integer(parameters?.Colors ?? 1, '颜色分量', 16);
  const bits = integer(parameters?.BitsPerComponent ?? 8, '颜色位数', 16);
  const columns = integer(parameters?.Columns ?? 1, '预测行宽', 1000000);
  if (bits !== 8 || !colors || !columns) throw new Error('当前 PDF 预测解码仅支持 8 位颜色分量。');
  const rowLength = colors * columns;
  if (predictor === 2) {
    if (bytes.length % rowLength) throw new Error('PDF TIFF 预测流长度无效。');
    const output = bytes.slice();
    for (let row = 0; row < output.length; row += rowLength) for (let index = colors; index < rowLength; index += 1) output[row + index] = (output[row + index] + output[row + index - colors]) & 255;
    return output;
  }
  if (predictor < 10 || predictor > 15 || bytes.length % (rowLength + 1)) throw new Error('PDF PNG 预测参数或流长度无效。');
  const output = new Uint8Array(bytes.length / (rowLength + 1) * rowLength);
  for (let row = 0, source = 0; source < bytes.length; row += rowLength) {
    const filter = bytes[source++];
    if (filter > 4) throw new Error('PDF PNG 预测过滤器无效。');
    for (let col = 0; col < rowLength; col += 1) {
      const left = col >= colors ? output[row + col - colors] : 0;
      const up = row ? output[row + col - rowLength] : 0;
      const upperLeft = row && col >= colors ? output[row + col - rowLength - colors] : 0;
      let correction = 0;
      if (filter === 1) correction = left;
      if (filter === 2) correction = up;
      if (filter === 3) correction = Math.floor((left + up) / 2);
      if (filter === 4) {
        const target = left + up - upperLeft;
        const a = Math.abs(target - left), b = Math.abs(target - up), c = Math.abs(target - upperLeft);
        correction = a <= b && a <= c ? left : b <= c ? up : upperLeft;
      }
      output[row + col] = (bytes[source++] + correction) & 255;
    }
  }
  return output;
}

export class PdfDocument {
  constructor(bytes, warnings) {
    this.bytes = bytes;
    this.warnings = warnings;
    this.xref = new Map();
    this.objects = new Map();
    this.decoded = new Map();
    this.objectStreams = new Map();
    this.xrefOffsets = new Set();
    this.totalDecoded = 0;
    this.budget = { tokens: 0 };
    this.trailer = Object.create(null);
  }

  warn(message) { if (!this.warnings.includes(message)) this.warnings.push(message); }

  async load() {
    if (this.bytes.length > 32 * 1024 * 1024) throw new Error('PDF 文件超过当前 32MB 上限。');
    if (!/^%PDF-1\.[0-7](?:\s|$)/.test(byteString(this.bytes.subarray(0, 16)))) throw new Error('不是支持的 PDF 1.x 文件。');
    const tail = byteString(this.bytes.subarray(Math.max(0, this.bytes.length - 16384)));
    const match = /startxref[\x00\t\n\f\r ]+(\d+)[\x00\t\n\f\r ]+%%EOF[\x00\t\n\f\r ]*$/.exec(tail);
    if (!match) throw new Error('PDF 缺少有效的 startxref / EOF，文件可能损坏。');
    await this.loadXref(Number(match[1]));
    if (this.trailer.Encrypt !== undefined && this.trailer.Encrypt !== null) throw new Error('当前不支持加密 PDF。');
    if (!isReference(this.trailer.Root)) throw new Error('PDF trailer 缺少文档目录引用。');
    const catalog = await this.object(this.trailer.Root);
    if (!isDictionary(catalog) || asName(catalog.Type) !== 'Catalog') throw new Error('PDF 文档目录无效。');
    if (catalog.AcroForm) this.warn('PDF 表单字段和交互控件尚未渲染。');
    if (catalog.OCProperties) this.warn('PDF 可选图层的可见性尚未计算。');
    return catalog;
  }

  async loadXref(offset, depth = 0, hybrid = false, revision = 0) {
    integer(offset, 'xref 偏移', this.bytes.length - 1);
    if (depth > 32 || this.xrefOffsets.has(offset)) throw new Error('PDF xref 链循环或层级过深。');
    this.xrefOffsets.add(offset);
    const lexer = new PdfLexer(this.bytes, offset, this.budget);
    const sectionIds = new Set();
    lexer.skip();
    let trailer;
    const start = lexer.position;
    const first = lexer.word();
    if (first === 'xref') {
      while (true) {
        const token = lexer.word();
        if (token === 'trailer') { trailer = lexer.value(); break; }
        if (!/^\d+$/.test(token)) throw new Error('PDF xref 小节头无效。');
        const initial = integer(Number(token), 'xref 对象编号');
        const count = integer(Number(lexer.word()), 'xref 项数');
        if (initial + count > 100000) throw new Error('PDF 对象数量超过当前 100000 上限。');
        for (let index = 0; index < count; index += 1) {
          if (sectionIds.has(initial + index)) throw new Error('PDF xref 小节包含重复对象。');
          sectionIds.add(initial + index);
          const position = integer(Number(lexer.word()), 'xref 对象偏移', this.bytes.length);
          const generation = integer(Number(lexer.word()), '对象代数', 65535);
          const flag = lexer.word();
          if (!['n', 'f'].includes(flag)) throw new Error('PDF xref 使用状态无效。');
          if (!this.xref.has(initial + index)) this.xref.set(initial + index, { type: flag === 'n' ? 1 : 0, position, generation, revision });
        }
      }
    } else {
      const record = await this.parseAt(start);
      trailer = record?.dictionary;
      if (!record?.stream || asName(trailer.Type) !== 'XRef') throw new Error('PDF startxref 没有指向 xref 表或 xref 流。');
      const widths = trailer.W;
      const size = integer(trailer.Size, 'xref Size');
      const indices = trailer.Index || [0, size];
      if (!Array.isArray(widths) || widths.length !== 3 || !Array.isArray(indices) || indices.length % 2) throw new Error('PDF xref 流参数无效。');
      widths.forEach(value => integer(value, 'xref 字段宽度', 8));
      const stride = widths.reduce((sum, value) => sum + value, 0);
      if (!stride) throw new Error('PDF xref 流字段全为空。');
      const bytes = await this.stream(record);
      let cursor = 0;
      for (let group = 0; group < indices.length; group += 2) {
        const initial = integer(indices[group], 'xref Index');
        const count = integer(indices[group + 1], 'xref Index 数量');
        if (initial + count > 100000) throw new Error('PDF xref 对象数量超限。');
        for (let index = 0; index < count; index += 1) {
          if (sectionIds.has(initial + index)) throw new Error('PDF xref 流索引包含重复对象。');
          sectionIds.add(initial + index);
          if (cursor + stride > bytes.length) throw new Error('PDF xref 流已截断。');
          const fields = widths.map((width, field) => {
            let value = width === 0 && field === 0 ? 1 : 0;
            for (let byte = 0; byte < width; byte += 1) value = value * 256 + bytes[cursor++];
            if (!Number.isSafeInteger(value)) throw new Error('PDF xref 字段超出安全整数范围。');
            return value;
          });
          // Unknown entry types denote null objects, and must not revive an older revision.
          const entry = fields[0] === 2 ? { type: 2, streamId: fields[1], index: fields[2], generation: 0 } : [0, 1].includes(fields[0]) ? { type: fields[0], position: fields[1], generation: fields[2] } : { type: 0, position: 0, generation: 0 };
          entry.revision = revision;
          const previous = this.xref.get(initial + index);
          if (!previous || hybrid && previous.type === 0 && previous.revision === revision) this.xref.set(initial + index, entry);
        }
      }
      if (cursor !== bytes.length) throw new Error('PDF xref 流长度与索引不一致。');
    }
    if (!isDictionary(trailer)) throw new Error('PDF trailer 不是有效字典。');
    for (const [key, value] of Object.entries(trailer)) if (!Object.hasOwn(this.trailer, key)) this.trailer[key] = value;
    if (trailer.Encrypt !== undefined && trailer.Encrypt !== null) throw new Error('当前不支持加密 PDF。');
    if (trailer.XRefStm !== undefined) await this.loadXref(trailer.XRefStm, depth + 1, true, revision);
    if (trailer.Prev !== undefined && !hybrid) await this.loadXref(trailer.Prev, depth + 1, false, revision + 1);
  }

  async parseAt(offset, expected, trail = []) {
    integer(offset, '对象偏移', this.bytes.length - 1);
    const lexer = new PdfLexer(this.bytes, offset, this.budget);
    const id = integer(Number(lexer.word()), '对象编号');
    const generation = integer(Number(lexer.word()), '对象代数', 65535);
    if (lexer.word() !== 'obj' || expected && (id !== expected.ref || generation !== expected.generation)) throw new Error('PDF xref 与对象头不一致。');
    const value = lexer.value();
    const token = lexer.word();
    if (token === 'endobj') return value;
    if (token !== 'stream' || !isDictionary(value)) throw new Error('PDF 对象没有正确结束。');
    while ([0, 9, 12, 32].includes(this.bytes[lexer.position])) lexer.position += 1;
    if (this.bytes[lexer.position] === 13) { lexer.position += 1; if (this.bytes[lexer.position] === 10) lexer.position += 1; }
    else if (this.bytes[lexer.position] === 10) lexer.position += 1;
    else throw new Error('PDF stream 后缺少换行。');
    const length = integer(await this.resolve(value.Length, trail), '流长度', this.bytes.length);
    const start = lexer.position;
    if (start + length > this.bytes.length) throw new Error('PDF 流已截断。');
    lexer.position = start + length;
    if (lexer.word() !== 'endstream' || lexer.word() !== 'endobj') throw new Error('PDF 流长度或结束标记无效。');
    return { dictionary: value, stream: this.bytes.subarray(start, start + length), id, generation };
  }

  async resolve(value, trail = []) { return isReference(value) ? this.object(value, trail) : value; }

  async object(reference, trail = []) {
    if (!isReference(reference)) return reference;
    const key = `${reference.ref}:${reference.generation}`;
    if (trail.includes(key) || trail.length > 64) throw new Error('PDF 对象引用循环或层级过深。');
    if (this.objects.has(key)) return this.objects.get(key);
    const entry = this.xref.get(reference.ref);
    if (!entry || entry.type === 0 || entry.generation !== reference.generation) throw new Error(`PDF 对象引用无效：${key}。`);
    const promise = entry.type === 1 ? this.parseAt(entry.position, reference, [...trail, key]) : this.compressedObject(reference, entry, [...trail, key]);
    this.objects.set(key, promise);
    return promise;
  }

  async compressedObject(reference, entry, trail) {
    if (!this.objectStreams.has(entry.streamId)) {
      const load = (async () => {
        const stream = await this.object({ ref: entry.streamId, generation: 0 }, trail);
        if (!stream?.stream || asName(stream.dictionary.Type) !== 'ObjStm') throw new Error('PDF 压缩对象没有有效的对象流。');
        const count = integer(stream.dictionary.N, '对象流数量');
        const data = await this.stream(stream);
        const first = integer(stream.dictionary.First, '对象流首偏移', data.length);
        const lexer = new PdfLexer(data, 0, this.budget);
        const entries = [];
        for (let index = 0; index < count; index += 1) entries.push({ id: integer(Number(lexer.word()), '对象流对象号'), offset: integer(Number(lexer.word()), '对象流相对偏移', data.length) });
        if (lexer.position > first) throw new Error('PDF 对象流头越界。');
        for (let index = 0; index < entries.length; index += 1) {
          const entry = entries[index];
          if (first + entry.offset >= data.length || index && entry.offset <= entries[index - 1].offset) throw new Error('PDF 对象流的对象偏移无效或未按顺序排列。');
          entry.end = index + 1 < entries.length ? first + entries[index + 1].offset : data.length;
        }
        return { entries, data, first };
      })();
      this.objectStreams.set(entry.streamId, load);
    }
    const stream = await this.objectStreams.get(entry.streamId);
    const selected = stream.entries[entry.index];
    if (!selected || selected.id !== reference.ref || stream.first + selected.offset >= stream.data.length) throw new Error('PDF 压缩对象索引不匹配。');
    const bytes = stream.data.subarray(stream.first + selected.offset, selected.end);
    const lexer = new PdfLexer(bytes, 0, this.budget);
    const value = lexer.value();
    lexer.skip();
    if (lexer.position !== bytes.length) throw new Error('PDF 对象流的单个对象包含多余内容。');
    return value;
  }

  async stream(record) {
    if (!record?.stream) throw new Error('PDF 内容引用不是流对象。');
    if (this.decoded.has(record)) return this.decoded.get(record);
    const decode = (async () => {
      let bytes = record.stream;
      const rawFilters = await this.resolve(record.dictionary.Filter);
      const filters = rawFilters ? Array.isArray(rawFilters) ? rawFilters : [rawFilters] : [];
      const rawParameters = await this.resolve(record.dictionary.DecodeParms);
      const parameters = Array.isArray(rawParameters) ? rawParameters : filters.map(() => rawParameters);
      for (let index = 0; index < filters.length; index += 1) {
        const filter = asName(filters[index]);
        if (['FlateDecode', 'Fl'].includes(filter)) {
          let reader;
          try { reader = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate')).getReader(); }
          catch { throw new Error('当前环境不支持 PDF 所需的原生 zlib 解压。'); }
          const chunks = [];
          let length = 0;
          try {
            while (true) {
              const { value, done } = await reader.read();
              if (done) break;
              length += value.length;
              if (this.totalDecoded + length > MAX_DECODED) throw new Error('PDF 解压内容超过当前 128MB 上限。');
              chunks.push(value);
            }
          } catch (error) { await reader.cancel().catch(() => {}); throw error; }
          bytes = new Uint8Array(length);
          let offset = 0;
          for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
          bytes = unPredict(bytes, await this.resolve(parameters[index]));
        } else if (['ASCIIHexDecode', 'AHx'].includes(filter)) {
          const source = byteString(bytes).replace(/[\x00\t\n\f\r ]/g, '');
          if (!/^[0-9a-f]*>$/i.test(source)) throw new Error('PDF ASCIIHex 流无效。');
          let hex = source.slice(0, -1);
          if (hex.length % 2) hex += '0';
          const size = hex.length / 2;
          if (this.totalDecoded + size > MAX_DECODED) throw new Error('PDF ASCIIHex 解码超过当前 128MB 总量上限。');
          bytes = new Uint8Array(size);
          for (let index = 0; index < size; index += 1) bytes[index] = Number.parseInt(hex.slice(index * 2, index * 2 + 2), 16);
        } else if (['ASCII85Decode', 'A85'].includes(filter)) bytes = ascii85(bytes, MAX_DECODED - this.totalDecoded);
        else throw new Error(`当前尚不支持 PDF 流过滤器：${filter || '未知'}。`);
      }
      if (this.totalDecoded + bytes.length > MAX_DECODED) throw new Error('PDF 解码内容超过当前 128MB 上限。');
      this.totalDecoded += bytes.length;
      return bytes;
    })();
    this.decoded.set(record, decode);
    return decode;
  }
}
