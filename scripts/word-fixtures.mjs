import { mkdir, writeFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeZip, crc32 } from './generate-fixtures.mjs';

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

export async function generateWordFixtures(destination = new URL('../fixtures/', import.meta.url)) {
  await mkdir(destination, { recursive: true });
  const files = { 'styled.docx': buildStyledDocx() };
  for (const [name, bytes] of Object.entries(files)) await writeFile(new URL(name, destination), bytes);
  return Object.entries(files).map(([name, bytes]) => ({ name, bytes: bytes.length }));
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) console.log(JSON.stringify(await generateWordFixtures(), null, 2));
