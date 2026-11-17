import { mkdir, writeFile } from 'node:fs/promises';
import { deflateRawSync } from 'node:zlib';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

// Self-authored development samples, not a real-world Office compatibility corpus.
// Container and legacy layouts follow Microsoft's public MS-CFB, MS-DOC,
// MS-XLS and MS-PPT specifications; no third-party document libraries are used.
const FREE = 0xffffffff;
const END = 0xfffffffe;
const FAT = 0xfffffffd;
const xml = (text) => Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${text}`);
const utf16 = (text) => Buffer.from(text, 'utf16le');
const u16 = (value) => { const data = Buffer.alloc(2); data.writeUInt16LE(value); return data; };
const u32 = (value) => { const data = Buffer.alloc(4); data.writeUInt32LE(value >>> 0); return data; };

export function crc32(data) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export function makeZip(entries, { compression = 'deflate' } = {}) {
  const locals = [];
  const central = [];
  let offset = 0;
  for (const [name, content] of Object.entries(entries)) {
    const nameBytes = Buffer.from(name);
    const data = typeof content === 'string' ? Buffer.from(content) : Buffer.from(content);
    const method = compression === 'store' ? 0 : 8;
    const packed = method ? deflateRawSync(data) : data;
    const checksum = crc32(data);
    const local = Buffer.alloc(30);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0x0800, 6);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(checksum, 14);
    local.writeUInt32LE(packed.length, 18);
    local.writeUInt32LE(data.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    locals.push(local, nameBytes, packed);
    const directory = Buffer.alloc(46);
    directory.writeUInt32LE(0x02014b50, 0);
    directory.writeUInt16LE(20, 4);
    directory.writeUInt16LE(20, 6);
    directory.writeUInt16LE(0x0800, 8);
    directory.writeUInt16LE(method, 10);
    directory.writeUInt32LE(checksum, 16);
    directory.writeUInt32LE(packed.length, 20);
    directory.writeUInt32LE(data.length, 24);
    directory.writeUInt16LE(nameBytes.length, 28);
    directory.writeUInt32LE(offset, 42);
    central.push(directory, nameBytes);
    offset += local.length + nameBytes.length + packed.length;
  }
  const directoryBytes = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(central.length / 2, 8);
  end.writeUInt16LE(central.length / 2, 10);
  end.writeUInt32LE(directoryBytes.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...locals, directoryBytes, end]);
}

export function makeCompound(streams) {
  const items = Object.entries(streams).map(([name, bytes], index) => ({
    name, bytes: Buffer.from(bytes), id: index + 1, left: FREE, right: FREE,
  }));
  const sectors = [];
  const links = [];
  function allocate(bytes) {
    if (!bytes.length) return END;
    const first = sectors.length;
    const count = Math.ceil(bytes.length / 512);
    for (let index = 0; index < count; index++) {
      const sector = Buffer.alloc(512);
      bytes.copy(sector, 0, index * 512, (index + 1) * 512);
      sectors.push(sector);
      links.push(index + 1 === count ? END : first + index + 1);
    }
    return first;
  }
  const miniChunks = [];
  const miniLinks = [];
  for (const item of items) {
    if (item.bytes.length >= 4096) item.start = allocate(item.bytes);
    else {
      item.start = item.bytes.length ? miniLinks.length : END;
      const count = Math.ceil(item.bytes.length / 64);
      for (let index = 0; index < count; index++) {
        const chunk = Buffer.alloc(64);
        item.bytes.copy(chunk, 0, index * 64, (index + 1) * 64);
        miniChunks.push(chunk);
        miniLinks.push(index + 1 === count ? END : miniLinks.length + 1);
      }
    }
  }
  const miniStream = Buffer.concat(miniChunks);
  const miniStart = allocate(miniStream);
  const miniFatData = Buffer.alloc(Math.ceil(miniLinks.length / 128) * 512, 0xff);
  miniLinks.forEach((link, index) => miniFatData.writeUInt32LE(link, index * 4));
  const miniFatStart = allocate(miniFatData);

  // MS-CFB directory names are ordered by UTF-16 name length and then case.
  const ordered = [...items].sort((a, b) => a.name.length - b.name.length || a.name.toUpperCase().localeCompare(b.name.toUpperCase()));
  function tree(nodes, depth = 0) {
    if (!nodes.length) return FREE;
    const middle = Math.floor(nodes.length / 2);
    const node = nodes[middle];
    node.depth = depth;
    node.left = tree(nodes.slice(0, middle), depth + 1);
    node.right = tree(nodes.slice(middle + 1), depth + 1);
    return node.id;
  }
  const child = tree(ordered);
  const deepest = Math.max(0, ...items.map((item) => item.depth));
  const directory = Buffer.alloc(Math.ceil((items.length + 1) / 4) * 512);
  for (let offset = 0; offset < directory.length; offset += 128) {
    directory.writeUInt32LE(FREE, offset + 68);
    directory.writeUInt32LE(FREE, offset + 72);
    directory.writeUInt32LE(FREE, offset + 76);
  }
  function entry(id, name, type, start, size, left = FREE, right = FREE, firstChild = FREE, color = 1) {
    const offset = id * 128;
    const nameBytes = utf16(name + '\0');
    if (nameBytes.length > 64) throw new Error('Fixture stream name exceeds 31 UTF-16 characters.');
    nameBytes.copy(directory, offset);
    directory.writeUInt16LE(nameBytes.length, offset + 64);
    directory[offset + 66] = type;
    directory[offset + 67] = color;
    directory.writeUInt32LE(left, offset + 68);
    directory.writeUInt32LE(right, offset + 72);
    directory.writeUInt32LE(firstChild, offset + 76);
    directory.writeUInt32LE(start, offset + 116);
    directory.writeUInt32LE(size, offset + 120);
  }
  entry(0, 'Root Entry', 5, miniStart, miniStream.length, FREE, FREE, child);
  items.forEach((item) => entry(item.id, item.name, 2, item.start, item.bytes.length, item.left, item.right, FREE, deepest > 0 && item.depth === deepest ? 0 : 1));
  const directoryStart = allocate(directory);
  let fatCount = 1;
  while (Math.ceil((sectors.length + fatCount) / 128) > fatCount) fatCount++;
  if (fatCount > 109) throw new Error('Fixture builder supports the header DIFAT only.');
  const firstFat = sectors.length;
  const fatData = Buffer.alloc(fatCount * 512, 0xff);
  links.forEach((link, index) => fatData.writeUInt32LE(link, index * 4));
  for (let index = 0; index < fatCount; index++) {
    fatData.writeUInt32LE(FAT, (firstFat + index) * 4);
    sectors.push(fatData.subarray(index * 512, (index + 1) * 512));
  }
  const header = Buffer.alloc(512);
  Buffer.from('d0cf11e0a1b11ae1', 'hex').copy(header);
  header.writeUInt16LE(0x003e, 24);
  header.writeUInt16LE(3, 26);
  header.writeUInt16LE(0xfffe, 28);
  header.writeUInt16LE(9, 30);
  header.writeUInt16LE(6, 32);
  header.writeUInt32LE(fatCount, 44);
  header.writeUInt32LE(directoryStart, 48);
  header.writeUInt32LE(4096, 56);
  header.writeUInt32LE(miniFatStart, 60);
  header.writeUInt32LE(miniFatData.length / 512, 64);
  header.writeUInt32LE(END, 68);
  for (let index = 0; index < 109; index++) header.writeUInt32LE(index < fatCount ? firstFat + index : FREE, 76 + index * 4);
  return Buffer.concat([header, ...sectors]);
}

export function buildLegacyDoc({ encrypted = false, mixed = false } = {}) {
  const main = '自主研发 Office 预览套件\r这是旧版 DOC 中文正文。\r格式、字体与分页仍需逐项实现。\r';
  const tail = mixed ? 'ASCII tail\r' : '';
  const text = main + tail;
  const word = Buffer.alloc(4096);
  word.writeUInt16LE(0xa5ec, 0);
  word.writeUInt16LE(0x00c1, 2);
  word.writeUInt16LE(0x0804, 6);
  word.writeUInt16LE(0x0204 | (encrypted ? 0x0100 : 0), 10);
  word.writeUInt16LE(0x00bf, 12);
  word.writeUInt32LE(2048, 24);
  word.writeUInt32LE(2048 + text.length * 2, 28);
  word.writeUInt16LE(14, 32);
  word.writeUInt16LE(22, 62);
  word.writeUInt32LE(word.length, 64);
  word.writeUInt32LE(text.length, 76);
  word.writeUInt16LE(93, 152);
  const pieceCount = mixed ? 2 : 1;
  const plcPcd = Buffer.alloc((pieceCount + 1) * 4 + pieceCount * 8);
  plcPcd.writeUInt32LE(0, 0);
  plcPcd.writeUInt32LE(main.length, 4);
  if (mixed) plcPcd.writeUInt32LE(text.length, 8);
  plcPcd.writeUInt32LE(2048, (pieceCount + 1) * 4 + 2);
  if (mixed) plcPcd.writeUInt32LE(0x40000000 | (3000 * 2), (pieceCount + 1) * 4 + 10);
  const table = Buffer.concat([Buffer.from([2]), u32(plcPcd.length), plcPcd]);
  word.writeUInt32LE(0, 154 + 33 * 8);
  word.writeUInt32LE(table.length, 158 + 33 * 8);
  utf16(main).copy(word, 2048);
  if (mixed) Buffer.from(tail, 'latin1').copy(word, 3000);
  return makeCompound({ WordDocument: word, '1Table': table });
}

function biff(type, bytes = Buffer.alloc(0)) { return Buffer.concat([u16(type), u16(bytes.length), bytes]); }
function biffBof(kind) {
  const bytes = Buffer.alloc(16);
  bytes.writeUInt16LE(0x0600, 0);
  bytes.writeUInt16LE(kind, 2);
  bytes.writeUInt16LE(0x0dbb, 4);
  bytes.writeUInt16LE(0x07cc, 6);
  bytes.writeUInt32LE(0x41, 8);
  bytes.writeUInt32LE(0x06, 12);
  return biff(0x0809, bytes);
}
function xlString(text, short = false) {
  return Buffer.concat([short ? Buffer.from([text.length, 1]) : Buffer.concat([u16(text.length), Buffer.from([1])]), utf16(text)]);
}

export function buildLegacyXls({ encrypted = false, continuation = false } = {}) {
  const strings = ['名称', '数量', '金额', '自研解析', '开发样本'];
  const sstBytes = Buffer.concat([u32(strings.length), u32(strings.length), ...strings.map((text) => xlString(text))]);
  // A CONTINUE split inside the first UTF-16 string includes a new encoding flag.
  const sst = continuation
    ? Buffer.concat([biff(0x00fc, sstBytes.subarray(0, 13)), biff(0x003c, Buffer.concat([Buffer.from([1]), sstBytes.subarray(13)]))])
    : biff(0x00fc, sstBytes);
  const sheetName = xlString('数据', true);
  const bound = Buffer.concat([u32(0), Buffer.from([0, 0]), sheetName]);
  const globalParts = [biffBof(0x0005), biff(0x0042, u16(1200)), biff(0x0022, u16(0))];
  if (encrypted) globalParts.push(biff(0x002f, Buffer.from([0, 0, 0, 0, 0, 0])));
  const globals = Buffer.concat([...globalParts, biff(0x0085, bound), sst, biff(0x000a)]);
  bound.writeUInt32LE(globals.length, 0);
  const patchedGlobals = Buffer.concat([...globalParts, biff(0x0085, bound), sst, biff(0x000a)]);
  const dimensions = Buffer.alloc(14);
  dimensions.writeUInt32LE(3, 4);
  dimensions.writeUInt16LE(3, 10);
  function label(row, col, index) { return biff(0x00fd, Buffer.concat([u16(row), u16(col), u16(0), u32(index)])); }
  function number(row, col, value) { const data = Buffer.alloc(14); data.writeUInt16LE(row); data.writeUInt16LE(col, 2); data.writeDoubleLE(value, 6); return biff(0x0203, data); }
  const bool = biff(0x0205, Buffer.concat([u16(2), u16(1), u16(0), Buffer.from([1, 0])]));
  const rk = biff(0x027e, Buffer.concat([u16(2), u16(2), u16(0), u32((42 << 2) | 2)]));
  const sheet = Buffer.concat([biffBof(0x0010), biff(0x0200, dimensions), biff(0x0081, Buffer.from([0xc1, 0x04])), label(0, 0, 0), label(0, 1, 1), label(0, 2, 2), label(1, 0, 3), number(1, 1, 12), number(1, 2, 1280.5), label(2, 0, 4), bool, rk, biff(0x000a)]);
  return makeCompound({ Workbook: Buffer.concat([patchedGlobals, sheet]) });
}

function pptRecord(type, data, version = 0, instance = 0) {
  const header = Buffer.alloc(8);
  header.writeUInt16LE((instance << 4) | version, 0);
  header.writeUInt16LE(type, 2);
  header.writeUInt32LE(data.length, 4);
  return Buffer.concat([header, data]);
}

export function buildLegacyPpt({ encrypted = false, orphan = true, reverseSlides = false } = {}) {
  function slidePersist(id, slideId) {
    const data = Buffer.alloc(20); data.writeUInt32LE(id); data.writeUInt32LE(slideId, 12);
    return pptRecord(0x03f3, data);
  }
  const docAtom = Buffer.alloc(40);
  docAtom.writeInt32LE(9600, 0); docAtom.writeInt32LE(5400, 4);
  docAtom.writeInt32LE(5400, 8); docAtom.writeInt32LE(7200, 12);
  const document = pptRecord(0x03e8, Buffer.concat([
    pptRecord(0x03e9, docAtom, 1),
    pptRecord(0x0ff0, Buffer.concat(reverseSlides ? [slidePersist(3, 257), slidePersist(2, 256)] : [slidePersist(2, 256), slidePersist(3, 257)]), 15),
    pptRecord(0x03ea, Buffer.alloc(0)),
  ]), 15);
  function slide(text, unicode) {
    const slideAtom = Buffer.alloc(24);
    const body = pptRecord(unicode ? 0x0fa0 : 0x0fa8, unicode ? utf16(text) : Buffer.from(text, 'latin1'));
    return pptRecord(0x03ee, Buffer.concat([pptRecord(0x03ef, slideAtom, 2), pptRecord(0xf00d, body, 15)]), 15);
  }
  const first = slide('自主研发 Office 预览\r旧版 PPT 中文文本', true);
  const second = slide('Local parsing\rNo document engine dependency', false);
  const deleted = orphan ? slide('ORPHAN DELETED SLIDE - MUST NOT APPEAR', false) : Buffer.alloc(0);
  const slide1Offset = document.length;
  const slide2Offset = slide1Offset + first.length;
  const persistOffset = slide2Offset + second.length + deleted.length;
  const persist = pptRecord(0x1772, Buffer.concat([u32((3 << 20) | 1), u32(0), u32(slide1Offset), u32(slide2Offset)]));
  const editOffset = persistOffset + persist.length;
  const edit = Buffer.alloc(28);
  edit.writeUInt32LE(257, 0); edit[7] = 3;
  edit.writeUInt32LE(persistOffset, 12); edit.writeUInt32LE(1, 16); edit.writeUInt32LE(4, 20); edit.writeUInt16LE(1, 24);
  const username = Buffer.from('Developer');
  const current = Buffer.alloc(20);
  current.writeUInt32LE(20, 0);
  current.writeUInt32LE(encrypted ? 0xf3d1c4df : 0xe391c05f, 4);
  current.writeUInt32LE(editOffset, 8);
  current.writeUInt16LE(username.length, 12); current.writeUInt16LE(0x03f4, 14); current[16] = 3;
  return makeCompound({
    'PowerPoint Document': Buffer.concat([document, first, second, deleted, persist, pptRecord(0x0ff5, edit)]),
    'Current User': pptRecord(0x0ff6, Buffer.concat([current, username, u32(8), utf16('Developer')])),
  });
}

const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CONTENT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';
const PNG = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64');
function rootRels(target) { return xml(`<Relationships xmlns="${REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="${target}"/></Relationships>`); }
function contentTypes(overrides) { return xml(`<Types xmlns="${CONTENT_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/>${Object.entries(overrides).map(([name, type]) => `<Override PartName="/${name}" ContentType="${type}"/>`).join('')}</Types>`); }

export function buildDocx() {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  return makeZip({
    '[Content_Types].xml': contentTypes({ 'word/document.xml': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml' }),
    '_rels/.rels': rootRels('word/document.xml'),
    'word/document.xml': xml(`<w:document xmlns:w="${W}" xmlns:r="${R_NS}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${A}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="184C4C"/><w:sz w:val="40"/></w:rPr><w:t>自主研发 Office 预览套件</w:t></w:r></w:p>
<w:p><w:r><w:t>这是 DOCX 基础预览开发样本，包含中文、格式、表格与图片。</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:i/><w:color w:val="B44B28"/></w:rPr><w:t>文件仅在本地解析，不依赖第三方文档内核。</w:t></w:r></w:p>
<w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4"/><w:left w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:right w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>格式</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>当前内容</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>DOCX</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>段落、直接格式、表格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="1" name="开发样本图片"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="pixel.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>`),
    'word/_rels/document.xml.rels': xml(`<Relationships xmlns="${REL_NS}"><Relationship Id="rIdImage" Type="${R_NS}/image" Target="media/pixel.png"/></Relationships>`),
    'word/media/pixel.png': PNG,
  });
}

export function buildXlsx() {
  const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  return makeZip({
    '[Content_Types].xml': contentTypes({
      'xl/workbook.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
      'xl/worksheets/sheet1.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      'xl/worksheets/sheet2.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
      'xl/sharedStrings.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
      'xl/styles.xml': 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    }),
    '_rels/.rels': rootRels('xl/workbook.xml'),
    'xl/workbook.xml': xml(`<workbook xmlns="${S}" xmlns:r="${R_NS}"><sheets><sheet name="产品清单" sheetId="1" r:id="sheetA"/><sheet name="说明" sheetId="2" r:id="sheetB"/></sheets></workbook>`),
    'xl/_rels/workbook.xml.rels': xml(`<Relationships xmlns="${REL_NS}"><Relationship Id="sheetA" Type="${R_NS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="sheetB" Type="${R_NS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="strings" Type="${R_NS}/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="styles" Type="${R_NS}/styles" Target="styles.xml"/></Relationships>`),
    'xl/sharedStrings.xml': xml(`<sst xmlns="${S}" count="5" uniqueCount="5"><si><t>项目</t></si><si><t>数量</t></si><si><t>单价</t></si><si><t>金额</t></si><si><r><t>自研</t></r><r><t>内核</t></r></si></sst>`),
    'xl/styles.xml': xml(`<styleSheet xmlns="${S}"><fonts count="2"><font><sz val="11"/><name val="Arial"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF24645B"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0" numFmtId="0"/><xf fontId="1" fillId="2" borderId="0" numFmtId="0" applyFill="1" applyFont="1"/></cellXfs></styleSheet>`),
    'xl/worksheets/sheet1.xml': xml(`<worksheet xmlns="${S}"><dimension ref="A1:D5"/><cols><col min="1" max="1" width="24" customWidth="1"/><col min="2" max="4" width="16" customWidth="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Office 预览开发样本</t></is></c></row><row r="2"><c r="A2" t="s" s="1"><v>0</v></c><c r="B2" t="s" s="1"><v>1</v></c><c r="C2" t="s" s="1"><v>2</v></c><c r="D2" t="s" s="1"><v>3</v></c></row><row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3"><v>12</v></c><c r="C3"><v>100.5</v></c><c r="D3"><f>B3*C3</f><v>1206</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><t>多工作表</t></is></c><c r="B4" t="b"><v>1</v></c><c r="D4"><v>42</v></c></row><row r="5"><c r="A5" t="inlineStr"><is><t>公式仅显示缓存值，不重新计算</t></is></c></row></sheetData><mergeCells count="2"><mergeCell ref="A1:D1"/><mergeCell ref="A5:D5"/></mergeCells></worksheet>`),
    'xl/worksheets/sheet2.xml': xml(`<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>这不是复杂真实文件的兼容验收。</t></is></c></row></sheetData></worksheet>`),
  });
}

export function buildPptx() {
  const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
  function textShape(id, text, x, y, cx, cy, size, color) {
    return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Text ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN" sz="${size}" b="1"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`;
  }
  function slide(title, body) {
    return xml(`<p:sld xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R_NS}"><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="F6F3EC"/></a:solidFill><a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${textShape(2, title, 685800, 685800, 10820400, 1371600, 3200, '184C4C')}${textShape(3, body, 685800, 2514600, 10820400, 1828800, 2000, '46544E')}<p:sp><p:nvSpPr><p:cNvPr id="4" name="Accent"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="5029200"/><a:ext cx="3657600" cy="182880"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="B55A3A"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
  }
  return makeZip({
    '[Content_Types].xml': contentTypes({
      'ppt/presentation.xml': 'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml',
      'ppt/slides/slide1.xml': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
      'ppt/slides/slide2.xml': 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    }),
    '_rels/.rels': rootRels('ppt/presentation.xml'),
    'ppt/presentation.xml': xml(`<p:presentation xmlns:p="${P}" xmlns:r="${R_NS}"><p:sldIdLst><p:sldId id="256" r:id="first"/><p:sldId id="257" r:id="second"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000" type="screen16x9"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>`),
    'ppt/_rels/presentation.xml.rels': xml(`<Relationships xmlns="${REL_NS}"><Relationship Id="first" Type="${R_NS}/slide" Target="slides/slide1.xml"/><Relationship Id="second" Type="${R_NS}/slide" Target="slides/slide2.xml"/></Relationships>`),
    'ppt/slides/slide1.xml': slide('Office，嵌入你的软件', '自研解析 / 本地预览 / 统一 SDK'),
    'ppt/slides/slide2.xml': slide('真实读取基础结构', '开发样本用于链路验证，复杂布局与动画仍需实现。'),
  });
}

export async function generateFixtures(destination = new URL('../fixtures/', import.meta.url)) {
  await mkdir(destination, { recursive: true });
  const csvText = '项目,数量,备注\r\n自研内核,12,"支持逗号,和双引号""示例"""\r\n本地预览,42,"支持\n多行内容"\r\n';
  const files = {
    'sample.csv': Buffer.concat([Buffer.from([0xef, 0xbb, 0xbf]), Buffer.from(csvText)]),
    'sample-utf16.csv': Buffer.concat([Buffer.from([0xff, 0xfe]), utf16('名称\t数量\r\n中文编码\t12\r\n')]),
    'sample.docx': buildDocx(), 'sample.xlsx': buildXlsx(), 'sample.pptx': buildPptx(),
    'sample.doc': buildLegacyDoc(), 'sample.xls': buildLegacyXls(), 'sample.ppt': buildLegacyPpt(),
  };
  await Promise.all(Object.entries(files).map(([name, bytes]) => writeFile(new URL(name, destination), bytes)));
  return Object.entries(files).map(([name, bytes]) => ({ name, bytes: bytes.length }));
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  console.log(JSON.stringify(await generateFixtures(), null, 2));
}
