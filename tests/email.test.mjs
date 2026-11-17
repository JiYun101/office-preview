import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { parseEmail } from '../src/parsers/email.js';
import { readCompound } from '../src/containers/cfb.js';
import { makeCompound } from '../scripts/generate-fixtures.mjs';
import { buildEmlFixture, buildMsgFixture, makeEmailCompound, msgProperties } from '../scripts/email-fixture.mjs';

test('EML reads folded UTF-8 B/Q headers, nested alternatives, quoted-printable and RFC2231 attachment names', async () => {
  const model = await parseEmail(buildEmlFixture(), 'eml', 'sample.eml');
  assert.equal(model.kind, 'email'); assert.equal(model.format, 'eml'); assert.equal(model.name, 'sample.eml');
  assert.equal(model.subject, 'Office 邮件预览 SDK');
  assert.equal(model.from, '研发团队 <team@example.invalid>');
  assert.equal(model.to, '预览用户 <viewer@example.invalid>');
  assert.equal(model.cc, 'review@example.invalid');
  assert.match(model.text, /正文优先使用纯文本/); assert.doesNotMatch(model.text, /HTML 正文/);
  assert.equal(model.attachments.length, 1); assert.equal(model.attachments[0].name, '说明.txt');
  assert.equal(model.attachments[0].mime, 'text/plain');
  assert.equal(new TextDecoder().decode(model.attachments[0].bytes), '附件内容：不自动打开。');
  assert.equal(model.attachments[0].size, model.attachments[0].bytes.length);
});

test('EML decodes GB18030 words/body and RFC2047 Q underscore spacing', async () => {
  const header = Buffer.from('From: =?UTF-8?Q?Hello_World?= <sender@example.invalid>\r\nSubject: =?GB18030?B?1tDOxA==?=\r\nContent-Type: text/plain; charset=GB18030\r\nContent-Transfer-Encoding: 8bit\r\n\r\n');
  const model = await parseEmail(Buffer.concat([header, Buffer.from([0xd6, 0xd0, 0xce, 0xc4])]), 'eml');
  assert.equal(model.subject, '中文'); assert.equal(model.text, '中文');
  assert.equal(model.from, 'Hello World <sender@example.invalid>');
});

test('EML joins RFC2231 filename continuation segments in order', async () => {
  const model = await parseEmail(buildEmlFixture({ filenameSegments: true }), 'eml');
  assert.equal(model.attachments[0].name, '说明.txt');
});

test('HTML-only EML preserves paragraphs/entities and strips scripts, styles and remote image markup', async () => {
  const before = globalThis.__emailExecuted;
  const model = await parseEmail(buildEmlFixture({ htmlOnly: true }), 'eml');
  assert.match(model.text, /邮件预览\n+HTML 正文 & 段落\n+仅显示文字。/);
  assert.doesNotMatch(model.text, /__emailExecuted|tracker\.png|body\{color|<img|<script/);
  assert.equal(globalThis.__emailExecuted, before);
  assert.ok(model.warnings.some(warning => warning.includes('HTML')));
});

test('HTML entities that resemble tags remain inert text rather than being parsed twice', async () => {
  const bytes = Buffer.from('Subject: inert\r\nContent-Type: text/html; charset=UTF-8\r\n\r\n<p>&lt;img src=x onerror=evil()&gt;</p><p>&#x4E2D;&#25991;</p>');
  const model = await parseEmail(bytes, 'eml');
  assert.equal(model.text, '<img src=x onerror=evil()>\n\n中文');
});

test('forwarded message/rfc822 is listed as an attachment without mixing its body into the outer email', async () => {
  const bytes = Buffer.from('Subject: outer\r\nContent-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\nContent-Type: text/plain\r\n\r\nOuter body\r\n--b\r\nContent-Type: message/rfc822\r\n\r\nSubject: inner\r\n\r\nInner body\r\n--b--\r\n');
  const model = await parseEmail(bytes, 'eml');
  assert.equal(model.text, 'Outer body');
  assert.equal(model.attachments[0].mime, 'message/rfc822');
  assert.match(new TextDecoder().decode(model.attachments[0].bytes), /Inner body/);
});

test('EML rejects missing mail headers, truncated multipart and malformed transfer encodings', async () => {
  await assert.rejects(parseEmail(Buffer.from('not an email'), 'eml'), /邮件头/);
  await assert.rejects(parseEmail(Buffer.from('X-Only: test\r\n\r\nrandom data'), 'eml'), /RFC5322/);
  await assert.rejects(parseEmail(Buffer.from('Subject: x\r\nContent-Type: multipart/mixed; boundary=b\r\n\r\n--b\r\n\r\ntext'), 'eml'), /闭合/);
  await assert.rejects(parseEmail(Buffer.from('Subject: x\r\nContent-Transfer-Encoding: base64\r\n\r\n***bad***'), 'eml'), /Base64/);
  await assert.rejects(parseEmail(Buffer.from('Subject: x\r\nContent-Transfer-Encoding: quoted-printable\r\n\r\n=XY'), 'eml'), /quoted-printable/);
});

test('EML enforces MIME depth, part count and original file-size limits', async () => {
  let nested = 'Content-Type: text/plain\r\n\r\nleaf';
  for (let index = 0; index < 17; index++) nested = `Content-Type: multipart/mixed; boundary=b${index}\r\n\r\n--b${index}\r\n${nested}\r\n--b${index}--\r\n`;
  await assert.rejects(parseEmail(Buffer.from('Subject: nested\r\n' + nested), 'eml'), /16 层/);
  const many = 'Subject: parts\r\nContent-Type: multipart/mixed; boundary=b\r\n\r\n' + '--b\r\nContent-Type: text/plain\r\n\r\nx\r\n'.repeat(1001) + '--b--\r\n';
  await assert.rejects(parseEmail(Buffer.from(many), 'eml'), /1000/);
  await assert.rejects(parseEmail(new Uint8Array(32 * 1024 * 1024 + 1), 'eml'), /32MB/);
});

for (const ansi of [false, true]) {
  test(`MSG reads ${ansi ? 'PT_STRING8 with code page' : 'Unicode'} fields, FILETIME, recipient storage and binary attachment`, async () => {
    const bytes = buildMsgFixture({ ansi });
    assert.ok(readCompound(bytes).names.some(path => path.startsWith('__attach_version1.0_#00000000/')), 'fixture must contain genuine CFB substorages');
    const model = await parseEmail(bytes, 'msg', 'sample.msg');
    assert.equal(model.subject, 'Office 邮件预览 MSG');
    assert.equal(model.from, '研发团队 <team@example.invalid>');
    assert.equal(model.to, '预览用户 <viewer@example.invalid>');
    assert.equal(model.cc, '审阅用户'); assert.equal(model.date, '2026-11-17T01:30:00.000Z');
    assert.match(model.text, /正文与附件均在本地读取/);
    assert.equal(model.attachments[0].name, '说明.txt');
    assert.equal(new TextDecoder().decode(model.attachments[0].bytes), 'MSG 附件内容：不自动打开。');
  });
}

test('MSG honors a legacy GB18030 code page for ANSI properties', async () => {
  const data = makeEmailCompound({
    '__properties_version1.0': msgProperties([[0x3ffd0003, 936]]),
    '__substg1.0_001A001E': Buffer.from('IPM.Note\0'),
    '__substg1.0_0037001E': Buffer.from([0xd6, 0xd0, 0xce, 0xc4, 0]),
    '__substg1.0_1000001E': Buffer.from([0xd6, 0xd0, 0xce, 0xc4, 0]),
  });
  const model = await parseEmail(data, 'msg');
  assert.equal(model.subject, '中文'); assert.equal(model.text, '中文');
});

test('MSG HTML fallback is inert plain text and compressed RTF-only mail reports its limitation', async () => {
  const html = await parseEmail(buildMsgFixture({ htmlOnly: true }), 'msg');
  assert.match(html.text, /HTML 正文 & 段落/); assert.doesNotMatch(html.text, /script|tracker|__emailExecuted/);
  const rtf = await parseEmail(buildMsgFixture({ rtfOnly: true }), 'msg');
  assert.equal(rtf.text, ''); assert.ok(rtf.warnings.some(warning => warning.includes('仅含压缩 RTF')));
});

test('MSG HTML charset fallback and Unicode transport headers keep their original characters', async () => {
  const data = makeEmailCompound({
    '__properties_version1.0': msgProperties([]),
    '__substg1.0_001A001F': Buffer.from('IPM.Note\0', 'utf16le'),
    '__substg1.0_007D001F': Buffer.from('From: Café <cafe@example.invalid>\r\nSubject: Café news\r\n\0', 'utf16le'),
    '__substg1.0_10130102': Buffer.from('<html><head><meta charset="UTF-8"></head><body><p>中文正文</p></body></html>'),
  });
  const model = await parseEmail(data, 'msg');
  assert.equal(model.subject, 'Café news'); assert.equal(model.from, 'Café <cafe@example.invalid>');
  assert.equal(model.text, '中文正文');
});

test('MSG embedded-message attachments expose metadata without recursively displaying their contents', async () => {
  const model = await parseEmail(buildMsgFixture({ embedded: true }), 'msg');
  assert.equal(model.attachments[0].name, '附加邮件.msg');
  assert.equal(model.attachments[0].mime, 'application/vnd.ms-outlook');
  assert.equal(model.attachments[0].bytes, undefined);
  assert.ok(model.warnings.some(warning => warning.includes('嵌入邮件')));
  assert.doesNotMatch(model.text, /嵌入邮件/);
});

test('MSG rejects other CFB formats, malformed property streams, broken Unicode and unsupported mailbox formats', async () => {
  await assert.rejects(parseEmail(makeCompound({ WordDocument: Buffer.alloc(4096) }), 'msg'), /不是受支持的 Outlook MSG/);
  await assert.rejects(parseEmail(makeEmailCompound({ '__properties_version1.0': Buffer.alloc(33) }), 'msg'), /属性流长度/);
  await assert.rejects(parseEmail(makeEmailCompound({ '__properties_version1.0': msgProperties([]), '__substg1.0_0037001F': Buffer.from([1]) }), 'msg'), /Unicode/);
  await assert.rejects(parseEmail(Buffer.from('not a compound file'), 'msg'), /OLE/);
  await assert.rejects(parseEmail(new Uint8Array(), 'pst'), /PST \/ OST/);
});

test('saved EML and MSG development samples both parse through their real byte containers', async () => {
  for (const format of ['eml', 'msg']) {
    const bytes = await readFile(new URL(`../fixtures/sample.${format}`, import.meta.url));
    const model = await parseEmail(bytes, format);
    assert.equal(model.kind, 'email'); assert.match(model.subject, /Office 邮件预览/); assert.equal(model.attachments.length, 1);
  }
});
