import { readCompound } from '../containers/cfb.js';

// Self-authored readers based on RFC 5322, 2045–2047, 2231 and MS-OXMSG /
// MS-OXCMSG. HTML is converted to text without creating a DOM or fetching URLs.
const MAX_BYTES = 32 * 1024 * 1024;
const MAX_PARTS = 1000;
const MAX_DEPTH = 16;
const HEADER_LIMIT = 256 * 1024;

function warn(warnings, text) { if (!warnings.includes(text)) warnings.push(text); }
function binaryString(bytes) {
  let result = '';
  for (let offset = 0; offset < bytes.length; offset += 8192) result += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
  return result;
}
function rawBytes(text) {
  const bytes = new Uint8Array(text.length);
  for (let index = 0; index < text.length; index++) bytes[index] = text.charCodeAt(index) & 255;
  return bytes;
}
function charge(state, length) {
  state.decoded += length;
  if (state.decoded > MAX_BYTES) throw new Error('邮件解码内容超过当前 32MB 预览上限。');
}
function charsetText(bytes, charset, warnings) {
  const requested = (charset || '').trim().toLowerCase();
  const aliases = { gb2312: 'gb18030', gbk: 'gb18030', cp936: 'gb18030', 'x-gbk': 'gb18030', utf8: 'utf-8', unicode: 'utf-16le' };
  const encoding = aliases[requested] || requested;
  if (encoding) {
    try { return new TextDecoder(encoding, { fatal: true }).decode(bytes); }
    catch { warn(warnings, `邮件字符集 ${requested} 不受支持或内容无效，已尝试 UTF-8 / GB18030 解码。`); }
  }
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
  catch {
    warn(warnings, '邮件部分内容缺少有效字符集声明，已尝试 GB18030；字符可能与原件不同。');
    return new TextDecoder('gb18030').decode(bytes);
  }
}

function base64Bytes(text) {
  const compact = text.replace(/[\t\r\n ]/g, '');
  if (!/^[A-Za-z0-9+/]*={0,2}$/.test(compact) || compact.length % 4 === 1 || compact.includes('=') && compact.length % 4 !== 0) {
    throw new Error('邮件 Base64 编码无效或已损坏。');
  }
  try { return rawBytes(atob(compact)); }
  catch { throw new Error('邮件 Base64 编码无效或已损坏。'); }
}
function quotedPrintable(text, header = false) {
  if (header) text = text.replace(/_/g, ' ');
  const bytes = new Uint8Array(text.length);
  let length = 0;
  for (let index = 0; index < text.length; index++) {
    if (text[index] !== '=') { bytes[length++] = text.charCodeAt(index) & 255; continue; }
    if (!header && text.slice(index + 1, index + 3) === '\r\n') { index += 2; continue; }
    if (!header && text[index + 1] === '\n') { index++; continue; }
    const hex = text.slice(index + 1, index + 3);
    if (!/^[0-9a-f]{2}$/i.test(hex)) throw new Error('邮件 quoted-printable 编码无效或已截断。');
    bytes[length++] = parseInt(hex, 16); index += 2;
  }
  return bytes.subarray(0, length);
}

function headerText(value, warnings, unicode = false) {
  // Adjacent RFC 2047 encoded words do not retain intervening folding spaces.
  const normalized = value.replace(/(\?=)[ \t]+(?==\?)/g, '$1');
  let decoded = normalized;
  if (!unicode && !/[\u0100-\uffff]/.test(normalized)) decoded = charsetText(rawBytes(normalized), '', warnings);
  return decoded.replace(/=\?([^?\s]+)\?([bq])\?([^?]*)\?=/gi, (_, charset, encoding, data) => {
    const bytes = encoding.toLowerCase() === 'b' ? base64Bytes(data) : quotedPrintable(data, true);
    return charsetText(bytes, charset.split('*')[0], warnings);
  });
}

function splitEntity(raw) {
  const separator = /\r?\n\r?\n/.exec(raw);
  if (!separator) throw new Error('EML 缺少邮件头与正文之间的空行。');
  if (separator.index > HEADER_LIMIT) throw new Error('邮件头超过当前 256KB 预览上限。');
  const header = raw.slice(0, separator.index);
  if (header.includes('\0') || /\r(?!\n)/.test(header)) throw new Error('EML 邮件头含无效控制字符。');
  const lines = header.split(/\r?\n/);
  const fields = [];
  for (const line of lines) {
    if (/^[ \t]/.test(line)) {
      if (!fields.length) throw new Error('EML 邮件头折行没有对应字段。');
      fields[fields.length - 1] += ' ' + line.trim();
    } else fields.push(line);
  }
  if (fields.length > 2000) throw new Error('邮件头字段数量超过当前预览上限。');
  const headers = new Map();
  for (const line of fields) {
    if (!line && fields.length === 1) continue;
    const match = /^([!-9;-~]+):[ \t]*(.*)$/.exec(line);
    if (!match) throw new Error('EML 邮件头字段格式无效。');
    const key = match[1].toLowerCase();
    if (!headers.has(key)) headers.set(key, match[2]);
  }
  return { headers, body: raw.slice(separator.index + separator[0].length) };
}

function valueParameters(raw = '') {
  const pieces = [];
  let quoted = false, escaped = false, start = 0;
  for (let index = 0; index < raw.length; index++) {
    const character = raw[index];
    if (escaped) { escaped = false; continue; }
    if (quoted && character === '\\') { escaped = true; continue; }
    if (character === '"') quoted = !quoted;
    if (!quoted && character === ';') { pieces.push(raw.slice(start, index)); start = index + 1; }
  }
  if (quoted) throw new Error('MIME 参数中的引号没有闭合。');
  pieces.push(raw.slice(start));
  const value = pieces.shift().trim().toLowerCase();
  const params = new Map();
  for (const piece of pieces) {
    const equals = piece.indexOf('=');
    if (equals < 0) continue;
    const name = piece.slice(0, equals).trim().toLowerCase();
    let parameter = piece.slice(equals + 1).trim();
    if (parameter.startsWith('"')) parameter = parameter.slice(1, -1).replace(/\\(.)/g, '$1');
    params.set(name, parameter);
  }
  return { value, params };
}
function extendedParameter(value, warnings) {
  const match = /^([^']*)'[^']*'(.*)$/.exec(value);
  const charset = match?.[1] || 'utf-8';
  const encoded = match?.[2] ?? value;
  const bytes = new Uint8Array(encoded.length);
  let length = 0;
  for (let index = 0; index < encoded.length; index++) {
    if (encoded[index] === '%') {
      const hex = encoded.slice(index + 1, index + 3);
      if (!/^[0-9a-f]{2}$/i.test(hex)) throw new Error('MIME RFC2231 文件名百分号编码无效。');
      bytes[length++] = parseInt(hex, 16); index += 2;
    } else bytes[length++] = encoded.charCodeAt(index) & 255;
  }
  return charsetText(bytes.subarray(0, length), charset, warnings);
}
function parameterName(params, key, warnings) {
  if (params.has(key + '*')) return extendedParameter(params.get(key + '*'), warnings);
  if (params.has(key + '*0*') || params.has(key + '*0')) {
    let joined = '';
    const encoded = params.has(key + '*0*');
    for (let index = 0; index < 100; index++) {
      const value = params.get(`${key}*${index}*`) ?? params.get(`${key}*${index}`);
      if (value === undefined) return encoded ? extendedParameter(joined, warnings) : headerText(joined, warnings);
      joined += value;
    }
    throw new Error('MIME 文件名续段数量超过当前预览上限。');
  }
  return params.has(key) ? headerText(params.get(key), warnings) : '';
}

const ENTITIES = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', ensp: ' ', emsp: ' ', hellip: '…', mdash: '—', ndash: '–', copy: '©', reg: '®', euro: '€', middot: '·', bull: '•' };
function htmlEntities(text) {
  return text.replace(/&(#x[0-9a-f]{1,8}|#\d{1,10}|[a-z]{2,16});/gi, (whole, entity) => {
    if (entity[0] !== '#') return ENTITIES[entity.toLowerCase()] ?? whole;
    const code = entity[1].toLowerCase() === 'x' ? parseInt(entity.slice(2), 16) : Number(entity.slice(1));
    return code > 0 && code <= 0x10ffff && !(code >= 0xd800 && code <= 0xdfff) ? String.fromCodePoint(code) : '\ufffd';
  });
}
function htmlToText(html) {
  const parts = [];
  let cursor = 0, tagCount = 0;
  while (cursor < html.length) {
    const start = html.indexOf('<', cursor);
    if (start < 0) { parts.push(html.slice(cursor)); break; }
    parts.push(html.slice(cursor, start));
    if (html.startsWith('<!--', start)) {
      const end = html.indexOf('-->', start + 4); cursor = end < 0 ? html.length : end + 3; continue;
    }
    let end = start + 1, quote = '';
    for (; end < html.length; end++) {
      const character = html[end];
      if (quote) { if (character === quote) quote = ''; }
      else if (character === '"' || character === "'") quote = character;
      else if (character === '>') break;
    }
    if (end === html.length) { parts.push(html.slice(start)); break; }
    if (++tagCount > 200000) throw new Error('邮件 HTML 标签数量超过当前预览上限。');
    const tag = /^\s*(\/?)\s*([a-z][a-z0-9:-]*)\b/i.exec(html.slice(start + 1, end));
    cursor = end + 1;
    if (!tag) continue;
    const closing = !!tag[1], name = tag[2].toLowerCase();
    if (!closing && ['script', 'style', 'head', 'template', 'iframe', 'object', 'svg', 'math'].includes(name)) {
      const close = new RegExp(`</${name}\\s*>`, 'ig'); close.lastIndex = cursor;
      const match = close.exec(html); cursor = match ? match.index + match[0].length : html.length;
      continue;
    }
    if (['p', 'div', 'section', 'article', 'blockquote', 'tr', 'li', 'table', 'ul', 'ol', 'header', 'footer', 'pre', 'hr', 'br', 'h1', 'h2', 'h3', 'h4', 'h5', 'h6'].includes(name)) parts.push('\n');
  }
  return htmlEntities(parts.join('')).replace(/\r/g, '').replace(/[ \t\f\v]+/g, ' ').replace(/ *\n */g, '\n').replace(/\n{3,}/g, '\n\n').trim();
}

function multipartParts(body, boundary) {
  if (!boundary || boundary.length > 70 || !/^[0-9a-z'()+_,\-./:=? ]+$/i.test(boundary) || boundary.endsWith(' ')) throw new Error('MIME multipart boundary 缺失或无效。');
  const escaped = boundary.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const delimiter = new RegExp(`^--${escaped}(--)?[ \\t]*(?:\\r?\\n|$)`, 'gm');
  const parts = [];
  let previous = null, closed = false, match;
  while ((match = delimiter.exec(body))) {
    if (previous !== null) {
      let end = match.index;
      if (body.slice(end - 2, end) === '\r\n') end -= 2;
      else if (body[end - 1] === '\n') end--;
      parts.push(body.slice(previous, end));
      if (parts.length > MAX_PARTS) throw new Error('MIME 分段数量超过当前 1000 项预览上限。');
    }
    if (match[1]) { closed = true; break; }
    previous = match.index + match[0].length;
  }
  if (!closed || !parts.length) throw new Error('MIME multipart 分隔符未闭合或没有有效分段。');
  return parts;
}
function transferBody(body, encoding) {
  if (encoding === 'base64') return base64Bytes(body);
  if (encoding === 'quoted-printable') return quotedPrintable(body);
  if (!encoding || ['7bit', '8bit', 'binary'].includes(encoding)) return rawBytes(body);
  throw new Error(`尚不支持邮件传输编码：${encoding}`);
}

function mimePart(entity, state, depth = 0) {
  if (depth > MAX_DEPTH || ++state.parts > MAX_PARTS) throw new Error('MIME 超过当前预览上限：16 层嵌套、1000 个分段。');
  const { headers, body } = entity;
  const type = valueParameters(headers.get('content-type') || 'text/plain');
  const disposition = valueParameters(headers.get('content-disposition') || '');
  const encoding = (headers.get('content-transfer-encoding') || '').trim().toLowerCase();
  if (type.value.startsWith('multipart/')) {
    if (encoding && !['7bit', '8bit', 'binary'].includes(encoding)) throw new Error('不支持经过传输编码的 multipart 容器。');
    if (['multipart/encrypted', 'multipart/signed'].includes(type.value)) warn(state.warnings, '邮件包含加密或签名封装；当前不解密、不验证签名。');
    const results = multipartParts(body, type.params.get('boundary')).map(raw => mimePart(splitEntity(raw), state, depth + 1));
    if (type.value === 'multipart/alternative') return results.find(result => result.rank === 2) || results.find(result => result.rank === 1) || { text: '', rank: 0 };
    return { text: results.map(result => result.text).filter(Boolean).join('\n\n'), rank: Math.max(0, ...results.map(result => result.rank)) };
  }
  const bytes = transferBody(body, encoding); charge(state, bytes.length);
  const filename = parameterName(disposition.params, 'filename', state.warnings) || parameterName(type.params, 'name', state.warnings);
  const attachment = filename || disposition.value === 'attachment' || !['text/plain', 'text/html'].includes(type.value);
  if (attachment) {
    if (state.attachments.length >= MAX_PARTS) throw new Error('邮件附件数量超过当前 1000 项预览上限。');
    state.attachments.push({ name: filename || `附件${state.attachments.length + 1}${type.value === 'message/rfc822' ? '.eml' : ''}`, size: bytes.length, mime: type.value || 'application/octet-stream', bytes });
    if (/pkcs7|pgp|x-tnef/.test(type.value)) warn(state.warnings, '部分附件为加密、签名或 TNEF 封装，仅列出原始附件，未解密或展开。');
    return { text: '', rank: 0 };
  }
  let text = charsetText(bytes, type.params.get('charset'), state.warnings).replace(/\r\n/g, '\n');
  if (type.value === 'text/html') {
    warn(state.warnings, 'HTML 正文已转换为纯文本，排版和内嵌图片不显示，远程资源不加载。');
    text = htmlToText(text);
  }
  return { text, rank: type.value === 'text/plain' ? 2 : 1 };
}

function parseEml(bytes, name) {
  const warnings = ['EML 当前预览文本正文与附件清单；不验证邮件签名，不解密 S/MIME / PGP，不执行 HTML 或加载远程资源。'];
  const raw = binaryString(bytes).replace(/^\xef\xbb\xbf/, '');
  const entity = splitEntity(raw);
  if (!['from', 'to', 'subject', 'date', 'message-id', 'mime-version'].some(key => entity.headers.has(key))) throw new Error('文件没有可识别的 RFC5322 邮件头。');
  const state = { warnings, attachments: [], parts: 0, decoded: 0 };
  const result = mimePart(entity, state);
  const header = key => headerText(entity.headers.get(key) || '', warnings);
  return { kind: 'email', format: 'eml', name, warnings, subject: header('subject'), from: header('from'), to: header('to'), cc: header('cc'), date: header('date'), text: result.text, attachments: state.attachments };
}

function propertyValues(bytes, headerLength) {
  if (!bytes || bytes.length < headerLength || (bytes.length - headerLength) % 16) throw new Error('MSG 的 MAPI 属性流长度无效。');
  const values = new Map();
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let offset = headerLength; offset < bytes.length; offset += 16) {
    const tag = view.getUint32(offset, true);
    values.set(tag, bytes.subarray(offset + 8, offset + 16));
  }
  return values;
}
function fixedNumber(values, tag) {
  const bytes = values.get(tag);
  return bytes ? new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getUint32(0, true) : undefined;
}
function fixedDate(values, tag) {
  const bytes = values.get(tag);
  if (!bytes) return '';
  const filetime = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength).getBigUint64(0, true);
  const milliseconds = Number(filetime / 10000n - 11644473600000n);
  return Math.abs(milliseconds) <= 8640000000000000 ? new Date(milliseconds).toISOString() : '';
}
function codepageCharset(codepage, warnings) {
  const known = { 65001: 'utf-8', 1200: 'utf-16le', 1201: 'utf-16be', 936: 'gb18030', 54936: 'gb18030', 950: 'big5', 932: 'shift_jis', 949: 'euc-kr', 20127: 'us-ascii', 28591: 'windows-1252' };
  if (known[codepage]) return known[codepage];
  if (codepage >= 1250 && codepage <= 1258) return `windows-${codepage}`;
  if (codepage !== undefined) warn(warnings, `MSG 代码页 ${codepage} 尚未实现，已按 Windows-1252 尝试读取 ANSI 字段。`);
  return 'windows-1252';
}

function parseMsg(bytes, name) {
  const compound = readCompound(bytes);
  const warnings = ['MSG 当前读取常用 MAPI 邮件字段、文本正文及附件清单；不还原 Outlook 专用表单、压缩 RTF、签名或加密内容。'];
  const state = { decoded: 0 };
  const names = new Map(compound.names.map(path => [path.toLowerCase(), path]));
  const seen = new Set();
  const get = path => {
    const actual = names.get(path.toLowerCase());
    if (!actual) return undefined;
    const data = compound.get(actual);
    if (!seen.has(actual)) { charge(state, data.length); seen.add(actual); }
    return data;
  };
  if (!names.has('__properties_version1.0')) throw new Error('文件缺少 MSG 顶层 MAPI 属性流，不是受支持的 Outlook MSG。');
  const properties = propertyValues(get('__properties_version1.0'), 32);
  const messageCodepage = fixedNumber(properties, 0x3ffd0003);
  const internetCodepage = fixedNumber(properties, 0x3fde0003);
  const messageCharset = codepageCharset(messageCodepage, warnings);
  const internetCharset = internetCodepage === undefined ? '' : codepageCharset(internetCodepage, warnings);
  const string = (property, base = '') => {
    const unicode = get(`${base}__substg1.0_${property}001F`);
    if (unicode) {
      if (unicode.length % 2) throw new Error('MSG Unicode 字符串长度无效。');
      return new TextDecoder('utf-16le', { fatal: true }).decode(unicode).replace(/\0+$/, '');
    }
    const ansi = get(`${base}__substg1.0_${property}001E`);
    return ansi ? charsetText(ansi, messageCharset, warnings).replace(/\0+$/, '') : '';
  };
  const messageClass = string('001A');
  if (messageClass && !/^IPM\.Note(?:\.|$)|^REPORT\.IPM\.Note(?:\.|$)/i.test(messageClass)) warn(warnings, '该 MSG 是会议、任务或其他 Outlook 对象；当前仅读取其通用邮件字段。');
  const transport = string('007D');
  let transportHeaders = new Map();
  if (transport) {
    try { transportHeaders = splitEntity(transport.replace(/\r?\n\r?\n[\s\S]*$/, '') + '\r\n\r\n').headers; }
    catch { warn(warnings, 'MSG 原始传输邮件头无效，已忽略这些回退字段。'); }
  }
  const transportHeader = key => headerText(transportHeaders.get(key) || '', warnings, true);
  const subject = string('0037') || transportHeader('subject');
  const senderName = string('0C1A') || string('0042');
  const senderAddress = string('5D01') || string('0C1F') || string('0065');
  const address = (display, email) => display && email && display !== email ? `${display} <${email}>` : email || display;
  const from = address(senderName, senderAddress) || transportHeader('from');
  const storagePaths = pattern => [...new Set(compound.names.map(path => pattern.exec(path)?.[1]).filter(Boolean))].sort();
  const recipients = storagePaths(/^(__recip_version1\.0_#[0-9a-f]{8})\//i);
  if (recipients.length > MAX_PARTS) throw new Error('MSG 收件人数量超过当前 1000 项预览上限。');
  const recipientLists = { 1: [], 2: [] };
  for (const storage of recipients) {
    const base = storage + '/';
    const data = get(base + '__properties_version1.0');
    const values = data ? propertyValues(data, 8) : new Map();
    const type = fixedNumber(values, 0x0c150003) ?? 1;
    const value = address(string('3001', base), string('39FE', base) || string('3003', base));
    if (recipientLists[type] && value) recipientLists[type].push(value);
  }
  const to = recipientLists[1].join('; ') || string('0E04') || transportHeader('to');
  const cc = recipientLists[2].join('; ') || string('0E03') || transportHeader('cc');
  const attachments = [];
  const attachmentStorages = storagePaths(/^(__attach_version1\.0_#[0-9a-f]{8})\//i);
  if (attachmentStorages.length > MAX_PARTS) throw new Error('MSG 附件数量超过当前 1000 项预览上限。');
  for (const storage of attachmentStorages) {
    const base = storage + '/';
    const propertyBytes = get(base + '__properties_version1.0');
    const values = propertyBytes ? propertyValues(propertyBytes, 8) : new Map();
    const method = fixedNumber(values, 0x37050003);
    const data = get(base + '__substg1.0_37010102');
    const embedded = method === 5 || [...names.keys()].some(path => path.startsWith((base + '__substg1.0_3701000D/').toLowerCase()));
    const filename = string('3707', base) || string('3704', base) || string('3001', base) || `附件${attachments.length + 1}${embedded ? '.msg' : ''}`;
    const mime = string('370E', base) || (embedded ? 'application/vnd.ms-outlook' : 'application/octet-stream');
    attachments.push({ name: filename, size: data?.length ?? fixedNumber(values, 0x0e200003) ?? 0, mime, ...(data ? { bytes: data } : {}) });
    if (!data) warn(warnings, '嵌入邮件、OLE 对象或按引用保存的附件当前仅列出元数据，不展开或访问外部文件。');
  }
  let text = string('1000');
  if (!text) {
    const html = get('__substg1.0_10130102');
    // Without InternetCodepage, honor an HTML charset declaration before using
    // the message's ANSI code page; raw UTF-8 is common in Unicode MSG files.
    const htmlMeta = html && /<meta\b[^>]*\bcharset\s*=\s*["']?\s*([a-z0-9_-]+)/i.exec(binaryString(html.subarray(0, 8192)));
    const htmlEncoding = internetCharset || htmlMeta?.[1] || (messageCodepage === undefined ? '' : messageCharset);
    const htmlString = html ? charsetText(html, htmlEncoding, warnings) : string('1013');
    if (htmlString) {
      text = htmlToText(htmlString);
      warn(warnings, 'MSG HTML 正文已转为纯文本，图片和原始排版不显示，远程资源不加载。');
    } else if (names.has('__substg1.0_10090102')) warn(warnings, '该 MSG 仅含压缩 RTF 正文；当前未实现 RTF 解压，正文暂无法显示。');
  }
  if (!messageClass && !subject && !from && !to && !text && !attachments.length) throw new Error('MSG 中没有可识别的邮件属性或正文。');
  const date = fixedDate(properties, 0x00390040) || fixedDate(properties, 0x0e060040) || transportHeader('date');
  return { kind: 'email', format: 'msg', name, warnings, subject, from, to, cc, date, text: text.replace(/\r\n/g, '\n'), attachments };
}

export async function parseEmail(input, format, name = '') {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.length > MAX_BYTES) throw new Error('邮件超过当前 32MB 原文件限制。');
  if (format === 'eml') return parseEml(bytes, name);
  if (format === 'msg') return parseMsg(bytes, name);
  throw new Error('当前邮件预览仅支持 EML / MSG 单封邮件，不支持 PST / OST 邮箱数据库。');
}
