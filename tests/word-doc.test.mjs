import test from 'node:test';
import assert from 'node:assert/strict';
import { makeCompound } from '../scripts/generate-fixtures.mjs';
import { readCompound } from '../src/containers/cfb.js';
import { parseLegacy } from '../src/parsers/legacy.js';

const u16 = value => { const out = Buffer.alloc(2); out.writeUInt16LE(value & 65535); return out; };
const u32 = value => { const out = Buffer.alloc(4); out.writeUInt32LE(value >>> 0); return out; };
const utf16 = text => Buffer.from(text, 'utf16le');
const prl = (id, value) => Buffer.concat([u16(id), value]);
const png = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=', 'base64');
const record = (type, payload, version = 0, instance = 0) => Buffer.concat([u16(instance * 16 + version), u16(type), u32(payload.length), payload]);

function style(index, name, parent, sti, paragraph, character, kind = 1) {
  const base = Buffer.alloc(10); base.writeUInt16LE(sti); base.writeUInt16LE(parent * 16 + kind, 2); base.writeUInt16LE(index * 16 + (kind === 1 ? 2 : 1), 4);
  const nameData = Buffer.concat([u16(name.length), utf16(name + '\0')]);
  const papx = Buffer.concat([u16(index), paragraph]);
  const chunks = [base, nameData];
  if (kind === 1) { chunks.push(u16(papx.length), papx); if (papx.length % 2) chunks.push(Buffer.from([0])); }
  chunks.push(u16(character.length), character);
  const data = Buffer.concat(chunks); return Buffer.concat([u16(data.length), data]);
}
function fkp(word, page, ranges, paragraph) {
  const target = word.subarray(page * 512, (page + 1) * 512), n = ranges.length;
  target[n * 4] = 0;
  for (let i = 0; i < n; i++) target.writeUInt32LE(ranges[i].start, i * 4);
  target.writeUInt32LE(ranges[n - 1].end, n * 4); target[511] = n;
  let position = 510;
  for (let i = 0; i < n; i++) {
    const body = ranges[i].properties;
    if (!body?.length) continue;
    const data = paragraph ? body.length % 2 ? Buffer.concat([Buffer.from([(body.length + 1) / 2]), body]) : Buffer.concat([Buffer.from([0, body.length / 2]), body]) : Buffer.concat([Buffer.from([body.length]), body]);
    position = Math.floor((position - data.length) / 2) * 2;
    data.copy(target, position); target[(n + 1) * 4 + i * (paragraph ? 13 : 1)] = position / 2;
  }
}

function styledDoc({ numbered = false, brokenImage = false, badFkp = false, hyperlink = false, directColor = false, markFont = false } = {}) {
  const paragraphs = ['目录\r', '\x13TOC \\o "1-3"\x14Office系统\t\x13PAGEREF demo\x145\x15\x15\r', '下一页\f示例\r', '\x01\r'];
  const text = paragraphs.join(''), word = Buffer.alloc(8192), textStart = 2048;
  word.writeUInt16LE(0xa5ec); word.writeUInt16LE(0xc1, 2); word.writeUInt16LE(0x804, 6); word.writeUInt16LE(0x204, 10);
  word.writeUInt32LE(textStart, 24); word.writeUInt32LE(textStart + text.length * 2, 28); word.writeUInt16LE(14, 32); word.writeUInt16LE(22, 62); word.writeUInt32LE(word.length, 64); word.writeUInt32LE(text.length, 76); word.writeUInt16LE(93, 152); utf16(text).copy(word, textStart);
  const chunks = []; let tableLength = 0;
  const part = (index, data) => { word.writeUInt32LE(tableLength, 154 + index * 8); word.writeUInt32LE(data.length, 158 + index * 8); chunks.push(data); tableLength += data.length; };
  const header = Buffer.alloc(18); header.writeUInt16LE(hyperlink ? 4 : 3); header.writeUInt16LE(10, 2); header.writeUInt16LE(1, 14);
  const tabs = Buffer.concat([Buffer.from([0,1]),u16(7200),Buffer.from([10])]);
  const normal = Buffer.concat([prl(0x4a43,u16(22)),prl(0x4a4f,u16(0)),prl(0x4a50,u16(1))]);
  part(1, Buffer.concat([u16(header.length), header,
    style(0,'Normal',4095,0,Buffer.alloc(0),normal),
    style(1,'TOC 1',0,19,Buffer.concat([prl(0x840f,u16(480)),prl(0xc60d,Buffer.concat([Buffer.from([tabs.length]),tabs])),...(numbered?[prl(0x460b,u16(1))]:[])]),Buffer.alloc(0)),
    style(2,'TOC Heading',0,4094,prl(0x2406,Buffer.from([1])),Buffer.concat([prl(0x4a43,u16(28)),prl(0x0835,Buffer.from([1])),prl(0x6870,u32(0x00b5752e))])),
    ...(hyperlink ? [style(3,'Hyperlink',4095,85,Buffer.alloc(0),Buffer.concat([prl(0x6870,u32(0x00c16305)),prl(0x2a3e,Buffer.from([1]))]),2)] : [])
  ]));
  function font(name) { const body=Buffer.alloc(39);body.writeUInt16LE(400,1);const bytes=Buffer.concat([body,utf16(name+'\0')]);return Buffer.concat([Buffer.from([bytes.length]),bytes]); }
  part(15,Buffer.concat([u16(2),u16(0),font('Arial'),font('宋体')]));
  const imageCp = text.indexOf('\x01'), picFc = textStart + imageCp * 2, endFc = textStart + text.length * 2;
  const common = Buffer.concat([...(hyperlink?[prl(0x4a30,u16(3))]:[]),...(directColor?[prl(0x6870,u32(0x00008000))]:[])]);
  const markFc=textStart+(paragraphs[0].length+paragraphs[1].length-1)*2;
  const characterRanges=markFont ? [{start:textStart,end:markFc,properties:common},{start:markFc,end:markFc+2,properties:Buffer.concat([common,prl(0x4a50,u16(0)),prl(0x4a43,u16(30))])},{start:markFc+2,end:picFc,properties:common}] : [{start:textStart,end:picFc,properties:common}];
  characterRanges.push({start:picFc,end:picFc+2,properties:Buffer.concat([prl(0x6a03,u32(0)),prl(0x0855,Buffer.from([1]))])},{start:picFc+2,end:endFc});
  fkp(word,8,characterRanges,false);
  part(12,Buffer.concat([u32(textStart),u32(endFc),u32(8)]));
  let cp=0; const ranges=paragraphs.map((value,index)=>{const range={start:textStart+cp*2,end:textStart+(cp+value.length)*2,properties:u16(index===0?2:index===1?1:0)};cp+=value.length;return range;});
  fkp(word,9,ranges,true);if(badFkp)word[9*512+(ranges.length+1)*4]=1;
  part(13,Buffer.concat([u32(textStart),u32(endFc),u32(9)]));
  const pcd=Buffer.concat([u32(0),u32(text.length),u16(0),u32(textStart),u16(0)]);part(33,Buffer.concat([Buffer.from([2]),u32(pcd.length),pcd]));
  const section=Buffer.concat([prl(0xb01f,u16(11906)),prl(0xb020,u16(16838)),prl(0xb021,u16(1800)),prl(0xb022,u16(1800)),prl(0x9023,u16(1440)),prl(0x9024,u16(1440)),prl(0x9031,u16(312))]);
  Buffer.concat([u16(section.length),section]).copy(word,5120);part(6,Buffer.concat([u32(0),u32(text.length),u16(0),u32(5120),u16(0),u32(0xffffffff)]));
  if(numbered){const lst=Buffer.alloc(28,0xff);lst.writeUInt32LE(123,0);lst.writeUInt32LE(0,4);lst[26]=1;lst[27]=0;const pap=Buffer.concat([prl(0x840f,u16(936)),prl(0x8411,u16(-360))]),lvlf=Buffer.alloc(28);lvlf.writeUInt32LE(1);lvlf[6]=1;lvlf[15]=0;lvlf[25]=pap.length;const first=Buffer.concat([u16(1),lst]);part(73,first);chunks.push(Buffer.concat([lvlf,pap,u16(2),utf16('\0.')]));tableLength+=28+pap.length+2+4;const lfo=Buffer.alloc(16);lfo.writeUInt32LE(123);part(74,Buffer.concat([u32(1),lfo,u32(0)]));}
  const blip=record(0xf01e,Buffer.concat([Buffer.alloc(16),Buffer.from([255]),png]),0,0x6e0),bse=Buffer.alloc(36);bse[0]=6;bse[1]=6;
  const picture=Buffer.concat([record(0xf004,Buffer.alloc(0),15),record(0xf007,Buffer.concat([bse,blip]),2,6)]),picf=Buffer.alloc(68);picf.writeUInt32LE(68+picture.length+(brokenImage?1000:0));picf.writeUInt16LE(68,4);picf.writeUInt16LE(100,6);picf.writeUInt16LE(1440,28);picf.writeUInt16LE(720,30);picf.writeUInt16LE(1000,32);picf.writeUInt16LE(1000,34);
  return makeCompound({WordDocument:word,'1Table':Buffer.concat(chunks),Data:Buffer.concat([picf,picture])});
}

function unpaddedData() {
  const content=Buffer.alloc(4096+111,91),full=makeCompound({Payload:content}),fat=full.readUInt32LE(76),fatOffset=(fat+1)*512,dir=(full.readUInt32LE(48)+1)*512;
  const newSector=full.length/512-1;let current=full.readUInt32LE(dir+128+116),previous;
  while(full.readUInt32LE(fatOffset+current*4)!==0xfffffffe){previous=current;current=full.readUInt32LE(fatOffset+current*4);}
  full.writeUInt32LE(newSector,fatOffset+previous*4);full.writeUInt32LE(0xffffffff,fatOffset+current*4);full.writeUInt32LE(0xfffffffe,fatOffset+newSector*4);
  return {bytes:Buffer.concat([full,content.subarray(4096)]),content};
}

test('CFB accepts omitted final user-stream padding while requiring all declared bytes',()=>{const {bytes,content}=unpaddedData();assert.equal(bytes.length%512,111);assert.deepEqual(Buffer.from(readCompound(bytes).get('Payload')),content);assert.throws(()=>readCompound(bytes.subarray(0,-1)).get('Payload'),/截断/);});
test('CFB still rejects incomplete structural allocation sectors',()=>{const bytes=makeCompound({Payload:Buffer.alloc(4096)});assert.throws(()=>readCompound(bytes.subarray(0,-1)),/结构扇区|完整/);});
test('DOC style inheritance, FKP paragraph styles and real font names survive reading',()=>{const model=parseLegacy(styledDoc(),'doc');const title=model.blocks[0],toc=model.blocks[1];assert.equal(title.style.fontSize,28*2/3);assert.equal(title.style.bold,true);assert.equal(title.style.color,'#2e75b5');assert.equal(title.style.fontFamily,'宋体');assert.equal(toc.isToc,true);assert.equal(toc.indentLeft,32);assert.deepEqual(toc.tabs,[{position:480,align:'right',leader:'dot'}]);assert.equal(toc.runs.find(run=>run.text==='Office').fontFamily,'Arial');});
test('DOC nested cached fields retain a real tab and cached page number without instructions',()=>{const toc=parseLegacy(styledDoc(),'doc').blocks[1];assert.equal(toc.runs.map(run=>run.text).join(''),'Office系统\t5');assert.equal(toc.runs.filter(run=>run.type==='tab').length,1);});
test('DOC explicit page breaks and SEPX page dimensions remain semantic model data',()=>{const model=parseLegacy(styledDoc(),'doc');assert.ok(model.blocks[2].runs.some(run=>run.type==='pageBreak'&&run.text==='\f'));assert.equal(model.page.width,11906/15);assert.equal(model.page.marginLeft,120);assert.equal(model.page.linePitch,20.8);});
test('DOC PICF/BSE/BLIP pictures are restored at their original display size',()=>{const image=parseLegacy(styledDoc(),'doc').blocks.flatMap(block=>block.runs).find(run=>run.image).image;assert.equal(image.width,96);assert.equal(image.height,48);assert.equal(image.mime,'image/png');assert.deepEqual(Buffer.from(image.bytes),png);});
test('DOC automatic numbering restores separate prefix metadata and list indents',()=>{const toc=parseLegacy(styledDoc({numbered:true}),'doc').blocks[1];assert.equal(toc.numbering.text,'1.');assert.equal(toc.numbering.suffix,'tab');assert.equal(toc.runs.filter(run=>run.type==='tab').length,1);assert.equal(toc.indentLeft,62.4);assert.equal(toc.firstLineIndent,-24);});
test('DOC malformed FKP ranges and PICF lengths fail rather than borrowing neighboring bytes',()=>{assert.throws(()=>parseLegacy(styledDoc({badFkp:true}),'doc'),/格式页索引/);assert.throws(()=>parseLegacy(styledDoc({brokenImage:true}),'doc'),/PICF/);});
test('DOC cached TOC links inherit paragraph appearance while body links keep hyperlink style',()=>{const model=parseLegacy(styledDoc({hyperlink:true}),'doc'),toc=model.blocks[1],body=model.blocks[2];assert.ok(toc.runs.every(run=>run.color==='#000000'&&!run.underline));assert.ok(toc.runs.every(run=>run.fontSize===22*2/3));assert.ok(body.runs.some(run=>run.color==='#0563c1'&&run.underline));});
test('DOC direct character color still wins inside a cached TOC field',()=>{const toc=parseLegacy(styledDoc({hyperlink:true,directColor:true}),'doc').blocks[1];assert.ok(toc.runs.every(run=>run.color==='#008000'));});
test('DOC paragraph mark keeps its own character metrics without changing visible runs',()=>{const toc=parseLegacy(styledDoc({markFont:true}),'doc').blocks[1];assert.equal(toc.paragraphMarkStyle.fontFamily,'Arial');assert.equal(toc.paragraphMarkStyle.fontSize,20);assert.ok(toc.runs.every(run=>run.fontSize!==20));});

function tableDoc({ indirect = false, damaged = false } = {}) {
  const withinTable = prl(0x2416, Buffer.from([1]));
  const data = [];
  let dataSize = 0;
  const rowEnd = (edges, flags, height) => {
    const cells = flags.map(flag => Buffer.concat([u16(flag), Buffer.alloc(18)]));
    const definition = Buffer.concat([Buffer.from([flags.length]), ...edges.map(u16), ...cells]);
    const columnProperties = [prl(0x7621, Buffer.concat([Buffer.from([0, flags.length]), u16(180)]))];
    for (let column = 0; column < flags.length; column++) {
      columnProperties.push(prl(0x7623, Buffer.concat([Buffer.from([column, column + 1]), u16(edges[column + 1] - edges[column])])));
      if (flags[column]) columnProperties.push(prl(0xd62b, Buffer.from([2, column, flags[column] >> 5])));
    }
    const properties = Buffer.concat([
      withinTable, prl(0x2417, Buffer.from([1])),
      ...(indirect ? columnProperties : [prl(0xd608, Buffer.concat([u16(definition.length + 1), definition]))]),
      prl(0x9407, u16(height)), prl(0x548a, u16(1)),
      prl(0xd605, Buffer.from([24, ...Array.from({ length: 6 }, () => [4, 1, 0, 0]).flat()])),
    ]);
    const reference = prl(indirect ? 0x646b : 0x6646, u32(dataSize));
    data.push(u16(properties.length), properties);
    dataSize += properties.length + 2;
    return reference;
  };
  const edges = [0, 1500, 3900, 4500];
  const records = [
    ['Before\r'], ['Horizontal\x07', withinTable], ['Right\x07', withinTable],
    ['\x07', rowEnd(damaged ? [0, 5000, 4500] : [0, 3900, 4500], [0, 0], -900)],
    ['Vertical\x07', withinTable], ['First paragraph\r', withinTable], ['Second paragraph\x07', withinTable],
    ['\x07', withinTable], ['\x07', rowEnd(edges, [0x60, 0, 0], 450)],
    ['\x07', withinTable], ['Tail\x07', withinTable], ['\x07', withinTable], ['\x07', rowEnd(edges, [0x20, 0, 0], 0)],
    ['After\r'],
  ];
  const text = records.map(([source]) => source).join(''), start = 2048, end = start + text.length * 2;
  const word = Buffer.alloc(8192);
  word.writeUInt16LE(0xa5ec, 0); word.writeUInt16LE(0xc1, 2); word.writeUInt16LE(0x204, 10);
  word.writeUInt16LE(14, 32); word.writeUInt16LE(22, 62); word.writeUInt32LE(word.length, 64);
  word.writeUInt32LE(text.length, 76); word.writeUInt16LE(93, 152);
  utf16(text).copy(word, start);
  let fc = start;
  const ranges = records.map(([source, properties]) => {
    const range = { start: fc, end: fc + source.length * 2, ...(properties ? { properties: Buffer.concat([u16(0), properties]) } : {}) };
    fc = range.end;
    return range;
  });
  fkp(word, 9, ranges, true);
  const papx = Buffer.concat([u32(start), u32(end), u32(9)]);
  word.writeUInt32LE(papx.length, 158 + 13 * 8);
  const pcd = Buffer.concat([u32(0), u32(text.length), u16(0), u32(start), u16(0)]);
  const clx = Buffer.concat([Buffer.from([2]), u32(pcd.length), pcd]);
  word.writeUInt32LE(papx.length, 154 + 33 * 8); word.writeUInt32LE(clx.length, 158 + 33 * 8);
  return makeCompound({ WordDocument: word, '1Table': Buffer.concat([papx, clx]), Data: Buffer.concat(data) });
}

test('DOC table keeps binary row heights, unequal columns, merges, and in-cell paragraphs', () => {
  for (const indirect of [false, true]) {
    const { blocks, warnings } = parseLegacy(tableDoc({ indirect }), 'doc');
    const [before, table, after] = blocks;
    assert.equal(before.runs[0].text, 'Before');
    assert.equal(after.runs[0].text, 'After');
    assert.equal(table.type, 'table');
    assert.deepEqual(table.columnWidths, [100, 160, 40]);
    assert.deepEqual(table.rowProperties, [{ height: 60, heightRule: 'exact' }, { height: 30, heightRule: 'atLeast' }, {}]);
    assert.equal(table.align, 'center');
    assert.equal(table.borders.insideH.width, 2 / 3);
    assert.deepEqual(table.rows.map(row => row.length), [2, 3, 2]);
    assert.equal(table.rows[0][0].colSpan, 2);
    assert.equal(table.rows[1][0].rowSpan, 2);
    assert.equal(table.rows[2][0].column, 1);
    assert.deepEqual(table.rows[1][1].blocks.map(block => block.runs[0].text), ['First paragraph', 'Second paragraph']);
    assert.ok(!warnings.some(message => message.includes('列边界')));
  }
});

test('DOC keeps cell text and a basic table when the binary grid is damaged', () => {
  const { blocks, warnings } = parseLegacy(tableDoc({ damaged: true }), 'doc');
  const table = blocks.find(block => block.type === 'table');
  assert.deepEqual(table.columnWidths, []);
  assert.equal(table.rows[0][0].blocks[0].runs[0].text, 'Horizontal');
  assert.equal(table.rows[2][1].blocks[0].runs[0].text, 'Tail');
  assert.ok(warnings.some(message => message.includes('列边界')));
});
