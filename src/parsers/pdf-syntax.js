// Self-authored PDF lexical reader. Structural syntax follows ISO 32000-1, §7.3.
export const isName = (value, expected) => !!value && typeof value.name === 'string' && (expected === undefined || value.name === expected);
export const isReference = value => !!value && Number.isInteger(value.ref) && Number.isInteger(value.generation);
export const isString = value => !!value && value.bytes instanceof Uint8Array;
export const isDictionary = value => !!value && typeof value === 'object' && !Array.isArray(value) && !isName(value) && !isReference(value) && !isString(value) && !value.stream;
export const asName = value => isName(value) ? value.name : '';
const whitespace = value => [0, 9, 10, 12, 13, 32].includes(value);
const delimiter = value => value === undefined || whitespace(value) || [40, 41, 60, 62, 91, 93, 123, 125, 47, 37].includes(value);

export function byteString(bytes) {
  const chunks = [];
  for (let index = 0; index < bytes.length; index += 16384) chunks.push(String.fromCharCode(...bytes.subarray(index, index + 16384)));
  return chunks.join('');
}

export class PdfLexer {
  constructor(bytes, position = 0, budget = { tokens: 0 }) {
    this.bytes = bytes;
    this.position = position;
    this.budget = budget;
  }

  skip() {
    while (this.position < this.bytes.length) {
      if (whitespace(this.bytes[this.position])) this.position += 1;
      else if (this.bytes[this.position] === 37) {
        while (this.position < this.bytes.length && ![10, 13].includes(this.bytes[this.position])) this.position += 1;
      } else break;
    }
  }

  word() {
    this.skip();
    const start = this.position;
    while (!delimiter(this.bytes[this.position])) {
      this.position += 1;
      if (this.position - start > 4096) throw new Error('PDF 关键字长度超过当前 4096 字节上限。');
    }
    if (start === this.position) throw new Error(`PDF 语法无效（字节 ${start}）。`);
    return byteString(this.bytes.subarray(start, this.position));
  }

  value(depth = 0, references = true) {
    this.skip();
    if (this.position >= this.bytes.length) throw new Error('PDF 对象意外结束。');
    if (depth > 64 || ++this.budget.tokens > 2000000) throw new Error('PDF 对象嵌套或令牌数量超过当前解析上限。');
    const code = this.bytes[this.position++];
    if (code === 47) {
      const start = this.position;
      while (!delimiter(this.bytes[this.position])) {
        this.position += 1;
        if (this.position - start > 4096) throw new Error('PDF 名称长度超过当前 4096 字节上限。');
      }
      const raw = byteString(this.bytes.subarray(start, this.position));
      if (/#(?![0-9a-f]{2})/i.test(raw)) throw new Error('PDF 名称转义无效。');
      return { name: raw.replace(/#([0-9a-f]{2})/gi, (_, hex) => String.fromCharCode(Number.parseInt(hex, 16))) };
    }
    if (code === 40) {
      const output = [];
      let nested = 1;
      while (this.position < this.bytes.length) {
        if (output.length > 2 * 1024 * 1024) throw new Error('PDF 单字符串超过当前 2MB 上限。');
        let character = this.bytes[this.position++];
        if (character === 92) {
          if (this.position >= this.bytes.length) throw new Error('PDF 字符串转义未结束。');
          character = this.bytes[this.position++];
          if ([10, 13].includes(character)) {
            if (character === 13 && this.bytes[this.position] === 10) this.position += 1;
            continue;
          }
          if (character >= 48 && character <= 55) {
            let octal = String.fromCharCode(character);
            for (let digit = 0; digit < 2 && this.bytes[this.position] >= 48 && this.bytes[this.position] <= 55; digit += 1) octal += String.fromCharCode(this.bytes[this.position++]);
            output.push(Number.parseInt(octal, 8) & 255);
          } else output.push(({ 110: 10, 114: 13, 116: 9, 98: 8, 102: 12 })[character] ?? character);
          continue;
        }
        if (character === 40 && ++nested > 128) throw new Error('PDF 字符串嵌套过深。');
        if (character === 41 && --nested === 0) return { bytes: new Uint8Array(output) };
        if (character === 13) {
          if (this.bytes[this.position] === 10) this.position += 1;
          character = 10;
        }
        output.push(character);
      }
      throw new Error('PDF 字符串未结束。');
    }
    if (code === 60 && this.bytes[this.position] !== 60) {
      let hex = '';
      while (this.position < this.bytes.length && this.bytes[this.position] !== 62) {
        const value = this.bytes[this.position++];
        if (!whitespace(value)) {
          if (!/[0-9a-f]/i.test(String.fromCharCode(value))) throw new Error('PDF 十六进制字符串无效。');
          hex += String.fromCharCode(value);
          if (hex.length > 4 * 1024 * 1024) throw new Error('PDF 单字符串超过当前 2MB 上限。');
        }
      }
      if (this.bytes[this.position++] !== 62) throw new Error('PDF 十六进制字符串未结束。');
      if (hex.length % 2) hex += '0';
      return { bytes: Uint8Array.from(hex.match(/.{2}/g) || [], pair => Number.parseInt(pair, 16)) };
    }
    if (code === 60 && this.bytes[this.position] === 60) {
      this.position += 1;
      const dictionary = Object.create(null);
      while (true) {
        this.skip();
        if (this.bytes[this.position] === 62 && this.bytes[this.position + 1] === 62) { this.position += 2; return dictionary; }
        const key = this.value(depth + 1, false);
        if (!isName(key) || Object.hasOwn(dictionary, key.name)) throw new Error('PDF 字典键无效或重复。');
        dictionary[key.name] = this.value(depth + 1, references);
      }
    }
    if (code === 91) {
      const items = [];
      while (true) {
        this.skip();
        if (this.bytes[this.position] === 93) { this.position += 1; return items; }
        items.push(this.value(depth + 1, references));
      }
    }
    this.position -= 1;
    const token = this.word();
    if (/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)$/.test(token)) {
      const number = Number(token);
      if (!Number.isFinite(number)) throw new Error('PDF 数值超出范围。');
      if (references && Number.isInteger(number) && number >= 0) {
        const saved = this.position;
        try {
          const generation = this.word();
          if (/^\d+$/.test(generation) && this.word() === 'R') return { ref: number, generation: Number(generation) };
        } catch { /* It was an ordinary numeric object. */ }
        this.position = saved;
      }
      return number;
    }
    if (token === 'true') return true;
    if (token === 'false') return false;
    if (token === 'null') return null;
    return { operator: token };
  }
}

export function numeric(value, label, fallback) {
  if (value === undefined && fallback !== undefined) return fallback;
  if (typeof value !== 'number' || !Number.isFinite(value)) throw new Error(`PDF ${label} 不是有效数值。`);
  return value;
}

export function integer(value, label, max = 100000) {
  const number = numeric(value, label);
  if (!Number.isSafeInteger(number) || number < 0 || number > max) throw new Error(`PDF ${label} 超出当前范围。`);
  return number;
}

export function matrixMultiply(left, right) {
  return [left[0] * right[0] + left[2] * right[1], left[1] * right[0] + left[3] * right[1], left[0] * right[2] + left[2] * right[3], left[1] * right[2] + left[3] * right[3], left[0] * right[4] + left[2] * right[5] + left[4], left[1] * right[4] + left[3] * right[5] + left[5]];
}
