import { mkdir, writeFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeCompound, makeZip, crc32 } from './generate-fixtures.mjs';

// Original, synthetic format fixtures. No user document content is copied here.
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const xml = text => Buffer.from('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + text);

function png() {
  const chunk = (name, data) => {
    const tag = Buffer.from(name), length = Buffer.alloc(4), checksum = Buffer.alloc(4);
    length.writeUInt32BE(data.length); checksum.writeUInt32BE(crc32(Buffer.concat([tag, data])));
    return Buffer.concat([length, tag, data, checksum]);
  };
  const header = Buffer.alloc(13); header.writeUInt32BE(8, 0); header.writeUInt32BE(4, 4); header[8] = 8; header[9] = 2;
  const rows = [];
  for (let y = 0; y < 4; y++) { const row = Buffer.alloc(25); for (let x = 0; x < 8; x++) { row[1+x*3] = x < 4 ? 35 : 231; row[2+x*3] = 120; row[3+x*3] = y < 2 ? 190 : 70; } rows.push(row); }
  return Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk('IHDR', header), chunk('IDAT', deflateSync(Buffer.concat(rows))), chunk('IEND', Buffer.alloc(0))]);
}

export function buildStyledDocx() {
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
  const text = value => `<w:r><w:t xml:space="preserve">${value}</w:t></w:r>`;
  const toc = (style, label, page, first = false, last = false) => `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${first ? '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>' : ''}<w:hyperlink w:anchor="_TocFixture">${text(label)}<w:r><w:tab/></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGEREF _TocFixture \\h </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${text(page)}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>${last ? '<w:r><w:fldChar w:fldCharType="end"/></w:r>' : ''}</w:p>`;
  const styles = `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="21"/><w:color w:val="000000"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
    <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:pPr><w:keepNext/><w:spacing w:before="120" w:after="160"/></w:pPr><w:rPr><w:b/><w:color w:val="2F75B5"/><w:sz w:val="32"/><w:rFonts w:eastAsia="黑体"/></w:rPr></w:style>
    <w:style w:type="paragraph" w:styleId="TOCHeading"><w:name w:val="TOC Heading"/><w:basedOn w:val="Heading1"/><w:pPr><w:spacing w:before="0" w:after="240"/></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="TOC1"><w:name w:val="toc 1"/><w:basedOn w:val="Normal"/><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs><w:spacing w:line="360" w:lineRule="auto" w:after="120"/><w:ind w:left="360"/></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="TOC2"><w:name w:val="toc 2"/><w:basedOn w:val="TOC1"/><w:pPr><w:ind w:left="720"/></w:pPr></w:style>
    <w:style w:type="paragraph" w:styleId="TOC3"><w:name w:val="toc 3"/><w:basedOn w:val="TOC2"/><w:pPr><w:ind w:left="1080"/></w:pPr></w:style>
    <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/><w:color w:val="C0504D"/></w:rPr></w:style>
  </w:styles>`;
  const drawing = `<w:r><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><wp:docPr id="1" name="Synthetic PNG"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="fixture.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="image1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
  const cell = value => `<w:tc><w:tcPr><w:tcW w:w="4500" w:type="dxa"/></w:tcPr><w:p>${text(value)}</w:p></w:tc>`;
  const document = `<w:document xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
    <w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr>${text('目录')}</w:p>
    ${toc('TOC1', '一、办公文档测试', '5', true)}
    ${toc('TOC2', '1. 二级目录示例', '12')}
    ${toc('TOC3', '1). 三级目录示例', '23', false, true)}
    <w:p><w:pPr><w:pStyle w:val="Heading1"/><w:pageBreakBefore/></w:pPr>${text('正文第一页')}</w:p>
    <w:p>${text('这是自主生成的 Word 样式与分页开发样本。')}<w:r><w:rPr><w:rStyle w:val="Emphasis"/></w:rPr><w:t>字符样式继承</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="7"/></w:numPr></w:pPr>${text('编号列表项目')}</w:p>
    <w:p>${drawing}</w:p>
    <w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4500"/><w:gridCol w:w="4500"/></w:tblGrid><w:tr>${cell('类型')}${cell('内容')}</w:tr><w:tr>${cell('开发样本')}${cell('不是用户资料')}</w:tr></w:tbl>
    <w:p>${text('显式分页前')}<w:r><w:br w:type="page"/></w:r>${text('显式分页后')}</w:p>
    <w:p>${text('缓存分页前')}<w:r><w:lastRenderedPageBreak/></w:r>${text('缓存分页后')}</w:p>
    <w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr>
  </w:body></w:document>`;
  return makeZip({
    '[Content_Types].xml': xml('<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'),
    '_rels/.rels': xml(`<Relationships xmlns="${REL}">${rel('office','officeDocument','word/document.xml')}</Relationships>`),
    'word/document.xml': xml(document), 'word/styles.xml': xml(styles),
    'word/_rels/document.xml.rels': xml(`<Relationships xmlns="${REL}">${rel('styles','styles','styles.xml')}${rel('numbering','numbering','numbering.xml')}${rel('settings','settings','settings.xml')}${rel('image1','image','media/fixture.png')}</Relationships>`),
    'word/numbering.xml': xml(`<w:numbering xmlns:w="${W}"><w:abstractNum w:abstractNumId="3"><w:multiLevelType w:val="singleLevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="720"/></w:tabs><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/></w:num></w:numbering>`),
    'word/settings.xml': xml(`<w:settings xmlns:w="${W}"><w:defaultTabStop w:val="720"/></w:settings>`),
    'word/media/fixture.png': png(),
  });
}

export function buildWordPageAreasDocx() {
  const p = text => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
  const rel = (id, type, target) => `<Relationship Id="${id}" Type="${R}/${type}" Target="${target}"/>`;
  const rels = content => xml(`<Relationships xmlns="${REL}">${content}</Relationships>`);
  const header = (id, type) => `<w:headerReference w:type="${type}" r:id="${id}"/>`;
  const footer = '<w:footerReference w:type="default" r:id="footer"/>';
  const geometry = '<w:pgSz w:w="9600" w:h="7200"/><w:pgMar w:top="720" w:right="600" w:bottom="720" w:left="900" w:header="480" w:footer="360"/>';
  const firstSection = `<w:sectPr>${header('first', 'first')}${header('odd', 'default')}${header('even', 'even')}${footer}${geometry}<w:pgNumType w:start="4"/><w:titlePg/></w:sectPr>`;
  const lastSection = `<w:sectPr>${header('last', 'default')}${geometry}<w:pgNumType w:start="8"/></w:sectPr>`;
  const before = (label, section = '') => `<w:p><w:pPr><w:pageBreakBefore/>${section}</w:pPr><w:r><w:t>${label}</w:t></w:r></w:p>`;
  const wordPart = (root, content) => xml(`<w:${root} xmlns:w="${W}" xmlns:r="${R}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${content}</w:${root}>`);
  const drawing = '<w:r><w:drawing><wp:inline><wp:extent cx="304800" cy="228600"/><a:graphic><a:graphicData><a:blip r:embed="logo"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>';
  const cachedField = '<w:p><w:fldSimple w:instr="PAGE"><w:r><w:t>Cached page</w:t></w:r></w:fldSimple></w:p>';
  const table = `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid><w:tr><w:tc>${cachedField}</w:tc></w:tr></w:tbl>`;
  const headerIds = ['first', 'odd', 'even', 'last'];
  const override = (part, kind) => `<Override PartName="/word/${part}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`;
  return makeZip({
    '[Content_Types].xml': xml(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="png" ContentType="image/png"/>${override('document', 'document.main')}${override('settings', 'settings')}${headerIds.map(id => override(id, 'header')).join('')}${override('footer', 'footer')}</Types>`),
    '_rels/.rels': rels(rel('office', 'officeDocument', 'word/document.xml')),
    'word/document.xml': wordPart('document', `<w:body>${p('First body')}${before('Second body')}${before('Third body', firstSection)}${p('Next section')}${before('Last body')}${lastSection}</w:body>`),
    'word/_rels/document.xml.rels': rels(rel('settings', 'settings', 'settings.xml') + headerIds.map(id => rel(id, 'header', `${id}.xml`)).join('') + rel('footer', 'footer', 'footer.xml')),
    'word/settings.xml': wordPart('settings', '<w:evenAndOddHeaders/>'),
    'word/first.xml': wordPart('hdr', p('First header')),
    'word/odd.xml': wordPart('hdr', p('Odd header')),
    'word/even.xml': wordPart('hdr', `<w:p><w:r><w:t>Even header</w:t></w:r>${drawing}</w:p>`),
    'word/_rels/even.xml.rels': rels(rel('logo', 'image', 'media/logo.png')),
    'word/last.xml': wordPart('hdr', p('Next odd header')),
    'word/footer.xml': wordPart('ftr', table),
    'word/media/logo.png': png(),
  });
}

export function buildWordPageAreasDoc({ invalidIndex = false } = {}) {
  const firstSection = 'First body\r\fSecond body\r\fThird body\r';
  const main = firstSection + 'Next section\r\fLast body\r';
  const footnotes = 'Footnote outside the page areas\r';
  const stories = [
    'Hidden separator\r', '', '', '', '', '',
    'Even header\r\r', 'Odd header\r\r', 'Even footer\r\r', '\x13PAGE\x145\x15\r\r', 'First header\r\r', 'First footer\r\r',
    '', 'Next odd header\r\r', '', '', '', '',
  ];
  const headerText = stories.join('') + '\r';
  const text = main + footnotes + headerText, split = main.length + footnotes.length + 20;
  const word = Buffer.alloc(4096), firstOffset = 2048, secondOffset = 2600;
  word.writeUInt16LE(0xa5ec); word.writeUInt16LE(0x00c1, 2); word.writeUInt16LE(0x0204, 10);
  word.writeUInt32LE(firstOffset, 24); word.writeUInt32LE(secondOffset + (text.length - split) * 2, 28);
  word.writeUInt16LE(14, 32); word.writeUInt16LE(22, 62); word.writeUInt32LE(word.length, 64);
  word.writeUInt32LE(main.length, 76); word.writeUInt32LE(footnotes.length, 80); word.writeUInt32LE(headerText.length, 84); word.writeUInt16LE(93, 152);
  word.write(text.slice(0, split), firstOffset, 'utf16le'); word.write(text.slice(split), secondOffset, 'utf16le');

  const sections = Buffer.alloc(36);
  sections.writeUInt32LE(firstSection.length, 4); sections.writeUInt32LE(main.length, 8);
  for (const [index, pageNumber] of [4, 8].entries()) {
    const offset = 3072 + index * 64;
    const properties = Buffer.from([
      0x0a, 0x30, index === 0 ? 1 : 0, 0x11, 0x30, 1, 0x1c, 0x50, pageNumber, 0,
      0x17, 0xb0, 0xe0, 0x01, 0x18, 0xb0, 0x68, 0x01,
    ]);
    word.writeUInt16LE(properties.length, offset); properties.copy(word, offset + 2);
    sections.writeUInt32LE(offset, 14 + index * 12);
    sections.writeUInt32LE(0xffffffff, 20 + index * 12);
  }
  const headerIndex = Buffer.alloc((stories.length + 2) * 4);
  let cp = 0;
  for (const [index, story] of stories.entries()) { headerIndex.writeUInt32LE(cp, index * 4); cp += story.length; }
  headerIndex.writeUInt32LE(cp + Number(invalidIndex), stories.length * 4);
  headerIndex.writeUInt32LE(cp + 3, (stories.length + 1) * 4);
  const dop = Buffer.alloc(500); dop[0] = 1;
  const clx = Buffer.alloc(33); clx[0] = 2; clx.writeUInt32LE(28, 1);
  clx.writeUInt32LE(split, 9); clx.writeUInt32LE(text.length, 13);
  clx.writeUInt32LE(firstOffset, 19); clx.writeUInt32LE(secondOffset, 27);
  const parts = [[6, sections], [11, headerIndex], [31, dop], [33, clx]];
  let offset = 0;
  for (const [index, part] of parts) {
    word.writeUInt32LE(offset, 154 + index * 8); word.writeUInt32LE(part.length, 158 + index * 8);
    offset += part.length;
  }
  return makeCompound({ WordDocument: word, '1Table': Buffer.concat(parts.map(([, part]) => part)) });
}

export async function generateWordFixtures(destination = new URL('../fixtures/', import.meta.url)) {
  await mkdir(destination, { recursive: true });
  const files = { 'styled.docx': buildStyledDocx(), 'page-areas.docx': buildWordPageAreasDocx(), 'page-areas.doc': buildWordPageAreasDoc() };
  for (const [name, bytes] of Object.entries(files)) await writeFile(new URL(name, destination), bytes);
  return Object.entries(files).map(([name, bytes]) => ({ name, bytes: bytes.length }));
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) console.log(JSON.stringify(await generateWordFixtures(), null, 2));
