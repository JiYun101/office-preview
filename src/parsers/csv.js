export function decodeText(bytes, encoding, warnings = []) {
  if (encoding && encoding !== 'auto') return new TextDecoder(encoding, { fatal: true }).decode(bytes);
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch {
    warnings.push('未检测到 UTF-8 / BOM，已按 GB18030 尝试解码；乱码时请手动选择编码。');
    return new TextDecoder('gb18030').decode(bytes);
  }
}

function detectDelimiter(text) {
  const counts = new Map([[',', 0], ['\t', 0], [';', 0]]);
  let quoted = false;
  let lines = 0;
  for (let i = 0; i < Math.min(text.length, 32768) && lines < 12; i++) {
    const ch = text[i];
    if (ch === '"') {
      if (quoted && text[i + 1] === '"') i++;
      else quoted = !quoted;
    } else if (!quoted) {
      if (ch === '\n') lines++;
      if (counts.has(ch)) counts.set(ch, counts.get(ch) + 1);
    }
  }
  return [...counts].sort((a, b) => b[1] - a[1])[0][0];
}

export function parseCsv(bytes, format = 'csv', name = '', options = {}) {
  const warnings = [];
  const text = decodeText(bytes, options.encoding, warnings).replace(/^\uFEFF/, '');
  const delimiter = options.delimiter && options.delimiter !== 'auto' ? options.delimiter : format === 'tsv' ? '\t' : detectDelimiter(text);
  if (delimiter.length !== 1 || /["\r\n]/.test(delimiter)) throw new Error('CSV 分隔符必须是单个有效字符。');
  const rows = [];
  let row = [], field = '', quoted = false, closed = false, count = 0;
  const pushField = () => {
    if (++count > 1000000 || row.length >= 1024) throw new Error('CSV 超过首版限制（100万单元格 / 1024列）。');
    row.push({ value: field }); field = ''; closed = false;
  };
  const pushRow = () => {
    pushField();
    if (rows.length >= 50000) throw new Error('CSV 超过首版 50000 行限制。');
    rows.push(row); row = [];
  };
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else { quoted = false; closed = true; }
      } else field += ch;
    } else if (ch === delimiter) pushField();
    else if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      pushRow();
    } else if (ch === '"' && !field.length && !closed) quoted = true;
    else {
      if (closed || ch === '"') throw new Error('CSV 引号格式无效：引号字段需用双引号转义，并在闭合后使用分隔符。');
      field += ch;
    }
  }
  if (quoted) throw new Error('CSV 存在未闭合的引号字段。');
  if (text.length && !/[\r\n]$/.test(text) || field.length || row.length || closed) pushRow();
  return { kind: 'spreadsheet', format, name, warnings, delimiter, sheets: [{ name: '数据', rows }] };
}
