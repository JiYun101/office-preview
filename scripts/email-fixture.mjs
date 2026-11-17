import { mkdir, writeFile } from 'node:fs/promises';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeCompound } from './generate-fixtures.mjs';

// Self-authored development messages. They exercise a specified subset, not a
// real-world Outlook compatibility corpus. No third-party email library is used.
const FREE = 0xffffffff;
const encodedWord = text => `=?UTF-8?B?${Buffer.from(text).toString('base64')}?=`;
function quotedPrintable(text) {
  let result = '', line = 0;
  const bytes = Buffer.from(text);
  for (let index = 0; index < bytes.length; index++) {
    const byte = bytes[index];
    if (byte === 13 && bytes[index + 1] === 10) { result += '\r\n'; line = 0; index++; continue; }
    const token = byte >= 33 && byte <= 126 && byte !== 61 ? String.fromCharCode(byte) : '=' + byte.toString(16).toUpperCase().padStart(2, '0');
    if (line + token.length > 72) { result += '=\r\n'; line = 0; }
    result += token; line += token.length;
  }
  return result;
}
const HTML_BODY = '<html><head><style>body{color:red}</style></head><body><h1>邮件预览</h1><p>HTML 正文 &amp; 段落</p><script>globalThis.__emailExecuted=true</script><img src="https://example.invalid/tracker.png" onerror="globalThis.__emailExecuted=true"><p>仅显示文字。</p></body></html>';
const PLAIN_BODY = '这是一封自主生成的 EML 开发邮件。\r\n正文优先使用纯文本。\r\n附件仅显示清单。';

export function buildEmlFixture({ htmlOnly = false, filenameSegments = false } = {}) {
  const mixed = '_office_mail_mix_2026';
  const alternative = '_office_mail_alt_2026';
  const bodyParts = [];
  if (!htmlOnly) bodyParts.push(`--${alternative}\r\nContent-Type: text/plain; charset=UTF-8\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n${quotedPrintable(PLAIN_BODY)}\r\n`);
  bodyParts.push(`--${alternative}\r\nContent-Type: text/html; charset=UTF-8\r\nContent-Transfer-Encoding: base64\r\n\r\n${Buffer.from(HTML_BODY).toString('base64')}\r\n--${alternative}--\r\n`);
  const filename = filenameSegments ? "filename*0*=UTF-8''%E8%AF%B4;\r\n filename*1*=%E6%98%8E.txt" : "filename*=UTF-8''%E8%AF%B4%E6%98%8E.txt";
  return Buffer.from([
    `From: ${encodedWord('研发团队')} <team@example.invalid>`,
    `To: ${encodedWord('预览用户')} <viewer@example.invalid>`,
    'Cc: review@example.invalid',
    'Date: Tue, 17 Nov 2026 09:30:00 +0800',
    `Subject: ${encodedWord('Office 邮件预览')}\r\n =?UTF-8?Q?_SDK?=`,
    'Message-ID: <self-authored-office-fixture@example.invalid>',
    'MIME-Version: 1.0',
    `Content-Type: multipart/mixed; boundary="${mixed}"`,
    '',
    `--${mixed}`,
    `Content-Type: multipart/alternative; boundary="${alternative}"`,
    '',
    bodyParts.join(''),
    `--${mixed}`,
    'Content-Type: text/plain; charset=UTF-8',
    `Content-Disposition: attachment; ${filename}`,
    'Content-Transfer-Encoding: base64',
    '',
    Buffer.from('附件内容：不自动打开。').toString('base64'),
    `--${mixed}--`,
    '',
  ].join('\r\n'));
}

// Reuse the existing allocation writer, then describe genuine nested CFB
// storages in its directory. Every leaf has its own stream and mini-FAT chain.
export function makeEmailCompound(streams) {
  const items = [];
  const byPath = new Map();
  const add = (path, bytes, storage) => {
    if (byPath.has(path)) return byPath.get(path);
    const parts = path.split('/');
    const name = parts.pop();
    const parent = parts.join('/');
    if (parent) add(parent, Buffer.alloc(0), true);
    const item = { path, name, parent, bytes: Buffer.from(bytes), storage, id: items.length + 1, left: FREE, right: FREE, child: FREE };
    items.push(item); byPath.set(path, item); return item;
  };
  for (const [path, bytes] of Object.entries(streams)) add(path, bytes, false);
  const data = makeCompound(Object.fromEntries(items.map(item => [`fixture_${item.id}`, item.bytes])));
  const directoryStart = (data.readUInt32LE(48) + 1) * 512;
  const groups = new Map();
  for (const item of items) { if (!groups.has(item.parent)) groups.set(item.parent, []); groups.get(item.parent).push(item); }
  const tree = (nodes, depth = 0) => {
    if (!nodes.length) return FREE;
    const middle = Math.floor(nodes.length / 2), node = nodes[middle];
    node.depth = depth; node.left = tree(nodes.slice(0, middle), depth + 1); node.right = tree(nodes.slice(middle + 1), depth + 1); return node.id;
  };
  for (const [parent, children] of groups) {
    children.sort((a, b) => a.name.length - b.name.length || a.name.toUpperCase().localeCompare(b.name.toUpperCase()));
    const child = tree(children);
    const deepest = Math.max(...children.map(item => item.depth));
    for (const item of children) item.color = deepest > 0 && item.depth === deepest ? 0 : 1;
    if (parent) byPath.get(parent).child = child;
    else data.writeUInt32LE(child, directoryStart + 76);
  }
  for (const item of items) {
    const offset = directoryStart + item.id * 128;
    const nameBytes = Buffer.from(item.name + '\0', 'utf16le');
    if (nameBytes.length > 64) throw new Error('MSG fixture directory name exceeds 31 characters.');
    data.fill(0, offset, offset + 64); nameBytes.copy(data, offset);
    data.writeUInt16LE(nameBytes.length, offset + 64); data[offset + 66] = item.storage ? 1 : 2; data[offset + 67] = item.color;
    data.writeUInt32LE(item.left, offset + 68); data.writeUInt32LE(item.right, offset + 72); data.writeUInt32LE(item.child, offset + 76);
    if (item.storage) { data.writeUInt32LE(0, offset + 116); data.writeUInt32LE(0, offset + 120); }
  }
  return data;
}

export function msgProperties(entries, headerLength = 32) {
  const data = Buffer.alloc(headerLength + entries.length * 16);
  entries.forEach(([tag, value], index) => {
    const offset = headerLength + index * 16;
    data.writeUInt32LE(tag, offset); data.writeUInt32LE(6, offset + 4);
    if (typeof value === 'bigint') data.writeBigUInt64LE(value, offset + 8);
    else data.writeUInt32LE(value, offset + 8);
  });
  return data;
}

export function buildMsgFixture({ ansi = false, htmlOnly = false, rtfOnly = false, embedded = false } = {}) {
  const streams = {};
  const rootEntries = [];
  const string = (property, text, base = '', entries = rootEntries) => {
    const type = ansi ? '001E' : '001F';
    const bytes = Buffer.from(text + '\0', ansi ? 'utf8' : 'utf16le');
    streams[`${base}__substg1.0_${property}${type}`] = bytes;
    entries.push([parseInt(property + type, 16), bytes.length]);
  };
  string('001A', 'IPM.Note'); string('0037', 'Office 邮件预览 MSG');
  string('0C1A', '研发团队'); string('0C1F', 'team@example.invalid');
  string('0E04', '预览用户'); string('0E03', '审阅用户');
  if (!htmlOnly && !rtfOnly) string('1000', '这是一封自主生成的 MSG 开发邮件。\r\n正文与附件均在本地读取。');
  if (htmlOnly) {
    const bytes = Buffer.from(HTML_BODY);
    streams['__substg1.0_10130102'] = bytes; rootEntries.push([0x10130102, bytes.length]);
  }
  if (rtfOnly) { streams['__substg1.0_10090102'] = Buffer.from([0, 1, 2, 3]); rootEntries.push([0x10090102, 4]); }
  const recipient = '__recip_version1.0_#00000000/';
  const recipientEntries = [[0x0c150003, 1]];
  string('3001', '预览用户', recipient, recipientEntries); string('39FE', 'viewer@example.invalid', recipient, recipientEntries);
  streams[recipient + '__properties_version1.0'] = msgProperties(recipientEntries, 8);
  const attachment = '__attach_version1.0_#00000000/';
  const attachmentEntries = [[0x37050003, embedded ? 5 : 1]];
  string('3707', embedded ? '附加邮件.msg' : '说明.txt', attachment, attachmentEntries);
  string('370E', embedded ? 'application/vnd.ms-outlook' : 'text/plain', attachment, attachmentEntries);
  if (embedded) {
    const base = attachment + '__substg1.0_3701000D/';
    const entries = [];
    string('0037', '嵌入邮件', base, entries); streams[base + '__properties_version1.0'] = msgProperties(entries, 24);
  } else {
    const bytes = Buffer.from('MSG 附件内容：不自动打开。');
    streams[attachment + '__substg1.0_37010102'] = bytes; attachmentEntries.push([0x37010102, bytes.length]);
  }
  streams[attachment + '__properties_version1.0'] = msgProperties(attachmentEntries, 8);
  rootEntries.push([0x3ffd0003, 65001], [0x3fde0003, 65001], [0x00390040, BigInt(Date.parse('2026-11-17T01:30:00.000Z') + 11644473600000) * 10000n]);
  const rootProperties = msgProperties(rootEntries);
  rootProperties.writeUInt32LE(1, 8); rootProperties.writeUInt32LE(1, 12); rootProperties.writeUInt32LE(1, 16); rootProperties.writeUInt32LE(1, 20);
  streams['__properties_version1.0'] = rootProperties;
  return makeEmailCompound(streams);
}

export async function generateEmailFixtures(destination = resolve(dirname(fileURLToPath(import.meta.url)), '../fixtures')) {
  await mkdir(destination, { recursive: true });
  await writeFile(resolve(destination, 'sample.eml'), buildEmlFixture());
  await writeFile(resolve(destination, 'sample.msg'), buildMsgFixture());
}
if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  await generateEmailFixtures();
  console.log('Generated self-authored email fixtures: fixtures/sample.eml, fixtures/sample.msg');
}
