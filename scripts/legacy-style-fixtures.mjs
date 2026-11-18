import { mkdir, writeFile } from 'node:fs/promises';
import { deflateSync } from 'node:zlib';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { crc32, makeCompound } from './generate-fixtures.mjs';

// Self-authored structural samples for our BIFF8 / OfficeArt readers. These are
// development records, not files exported by Office, and contain no user data.
const u16 = value => { const bytes = Buffer.alloc(2); bytes.writeUInt16LE(value); return bytes; };
const u32 = value => { const bytes = Buffer.alloc(4); bytes.writeUInt32LE(value >>> 0); return bytes; };
const utf16 = value => Buffer.from(value, 'utf16le');
const biff = (type, bytes = Buffer.alloc(0)) => Buffer.concat([u16(type), u16(bytes.length), bytes]);
const xlString = (text, short = false) => Buffer.concat([short ? Buffer.from([text.length, 1]) : Buffer.concat([u16(text.length), Buffer.from([1])]), utf16(text)]);

function bof(kind) {
  const bytes = Buffer.alloc(16);
  bytes.writeUInt16LE(0x0600); bytes.writeUInt16LE(kind, 2);
  bytes.writeUInt16LE(0x0dbb, 4); bytes.writeUInt16LE(0x07cc, 6);
  bytes.writeUInt32LE(0x41, 8); bytes.writeUInt32LE(6, 12);
  return biff(0x0809, bytes);
}

function font(name, points, { bold = false, italic = false, color = 0x7fff } = {}) {
  const bytes = Buffer.alloc(14);
  bytes.writeUInt16LE(points * 20); bytes.writeUInt16LE(italic ? 2 : 0, 2);
  bytes.writeUInt16LE(color, 4); bytes.writeUInt16LE(bold ? 700 : 400, 6);
  bytes[12] = 134;
  return biff(0x0031, Buffer.concat([bytes, xlString(name, true)]));
}

function xf(fontIndex = 0, format = 0, { style = false, parent = 0, align = 2, vertical = 1, fill = false, border = true, attrs = 0xfc, wrap = false, indent = 0 } = {}) {
  const bytes = Buffer.alloc(20);
  bytes.writeUInt16LE(fontIndex); bytes.writeUInt16LE(format, 2);
  bytes.writeUInt16LE(style ? 0xfff5 : (parent << 4) | 1, 4);
  bytes[6] = align | (vertical << 4) | (wrap ? 8 : 0);
  bytes[8] = indent; bytes[9] = attrs;
  if (border) {
    bytes.writeUInt16LE(0x1111, 10); // Thin left/right/top/bottom, automatic black.
    bytes.writeUInt16LE(64 | (64 << 7), 12);
  }
  bytes.writeUInt32LE((border ? 64 | (64 << 7) : 0) | (fill ? 1 << 26 : 0), 14);
  bytes.writeUInt16LE((fill ? 9 : 64) | (65 << 7), 18);
  return biff(0x00e0, bytes);
}

export function buildStyledXls({ hidden = false } = {}) {
  const strings = ['旧版 XLS 格式还原验证', '序号', '代码', '日期', '完成率', '继承字体', '跨过字体索引4'];
  const sst = biff(0x00fc, Buffer.concat([u32(strings.length), u32(strings.length), ...strings.map(value => xlString(value))]));
  const bound = Buffer.concat([u32(0), Buffer.from([0, 0]), xlString('样式验证', true)]);
  const colors = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCCC','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
  colors.splice(0,3,'184C4C','E8F1FF','C43B44');
  const palette = biff(0x0092, Buffer.concat([u16(colors.length), ...colors.map(value=>Buffer.concat([Buffer.from(value,'hex'),Buffer.from([0])]))]));
  const fonts = [
    font('宋体', 12), font('宋体', 24, { color: 8 }), font('宋体', 12, { bold: true }),
    font('Courier New', 10), font('Arial', 14, { italic: true, color: 10 }),
  ]; // The fifth FONT is index 5 because BIFF reserves FontIndex 4.
  const formats = [[164, '000000'], [165, 'yyyy-mm-dd'], [166, '0.00%'], [167, '$0.000']]
    .map(([id, value]) => biff(0x041e, Buffer.concat([u16(id), xlString(value)])));
  const styles = [
    xf(0,0,{style:true,border:false,align:0,vertical:2}),
    xf(0,0,{border:false,align:0,vertical:2}),
    xf(1,0,{border:false}), xf(2,0,{fill:true,wrap:true}), xf(0,0),
    xf(0,164), xf(0,165), xf(0,166),
    xf(3,167,{fill:true,attrs:0x70}), xf(5,0,{align:1,indent:1}),
  ];
  const globalsBeforeBound = [bof(5), biff(0x0042,u16(1200)), biff(0x0022,u16(0)), ...fonts, ...formats, palette, ...styles];
  const globals = () => Buffer.concat([...globalsBeforeBound,biff(0x0085,bound),sst,biff(0x000a)]);
  bound.writeUInt32LE(globals().length);
  const dimensions = Buffer.alloc(14); dimensions.writeUInt32LE(5,4); dimensions.writeUInt16LE(4,10);
  const records = [bof(0x10),biff(0x0200,dimensions),biff(0x0225,Buffer.concat([u16(1),u16(360)])),biff(0x0055,u16(9))];
  [10,18,22,16].forEach((width,col) => records.push(biff(0x007d,Buffer.concat([u16(col),u16(col),u16(width*256),u16(1),u16(hidden && col === 1 ? 1 : 0),u16(0)]))));
  [720,600,480,450,450].forEach((height,row) => {
    const bytes = Buffer.alloc(16); bytes.writeUInt16LE(row); bytes.writeUInt16LE(4,4); bytes.writeUInt16LE(height,6);
    bytes[12] = 0x40; bytes[13] = 1; bytes.writeUInt16LE(1,14);
    if (hidden && row === 1) bytes[12] |= 0x20;
    records.push(biff(0x0208,bytes));
  });
  const label = (row,col,style,index) => biff(0x00fd,Buffer.concat([u16(row),u16(col),u16(style),u32(index)]));
  const number = (row,col,style,value) => { const bytes=Buffer.concat([u16(row),u16(col),u16(style),Buffer.alloc(8)]); bytes.writeDoubleLE(value,6); return biff(0x0203,bytes); };
  records.push(label(0,0,2,0));
  for (let col=0;col<4;col++) records.push(label(1,col,3,col+1));
  records.push(number(2,0,4,1),number(2,1,5,42),number(2,2,6,46000),number(2,3,7,0.125),label(3,0,8,5),label(3,1,9,6));
  records.push(biff(0x00be,Buffer.concat([u16(3),u16(2),u16(3),u16(3),u16(3)])));
  records.push(biff(0x0201,Buffer.concat([u16(4),u16(0),u16(9)])));
  records.push(biff(0x00e5,Buffer.concat([u16(1),u16(0),u16(0),u16(0),u16(3)])),biff(0x000a));
  return makeCompound({ Workbook:Buffer.concat([globals(),...records]) });
}

function pptRecord(type, data, version=0, instance=0) {
  const header=Buffer.alloc(8); header.writeUInt16LE((instance<<4)|version); header.writeUInt16LE(type,2); header.writeUInt32LE(data.length,4);
  return Buffer.concat([header,data]);
}
function colorRef(color) { const bytes=Buffer.from(color.replace('#',''),'hex'); return bytes[0] | bytes[1]<<8 | bytes[2]<<16; }
function properties(values,type=0xf00b) { const entries=Object.entries(values).map(([id,value])=>Buffer.concat([u16(Number(id)),u32(value)])); return pptRecord(type,Buffer.concat(entries),3,entries.length); }
function clientAnchor(x,y,width,height) { const bytes=Buffer.alloc(16); [y,x,x+width,y+height].forEach((value,index)=>bytes.writeInt32LE(value*6,index*4)); return pptRecord(0xf010,bytes); }
function shapeBounds(x,y,width,height) { return Buffer.concat([x,y,x+width,y+height].map(value=>u32(value*6))); }
function textRecords(text,{size=26,color='#ffffff',bold=true,align=1,fontIndex=0}={}) {
  const paragraphRuns=text.split('\r').map(paragraph=>Buffer.concat([u32(paragraph.length+1),u16(0),u32(0x800),u16(align)]));
  const characterRun=Buffer.concat([u32(text.length+1),u32(0x70001),u16(bold?1:0),u16(fontIndex),u16(size),Buffer.from([...Buffer.from(color.slice(1),'hex'),0xfe])]);
  return pptRecord(0xf00d,Buffer.concat([
    pptRecord(0x0f9f,u32(4)),pptRecord(0x0fa0,utf16(text)),pptRecord(0x0fa1,Buffer.concat([...paragraphRuns,characterRun]))
  ]));
}
function shape(id,{x,y,width,height,fill='#184c4c',line='#184c4c',text,size=26,textColor='#ffffff',align=1,pib,shapeType=1,fontIndex=0,inGroup=false}={}) {
  const fsp=Buffer.concat([u32(id),u32(inGroup?0x0a02:0x0a00)]);
  const options={0x0181:colorRef(fill),0x01cb:9525,0x01bf:0x00100010,0x01ff:line?0x00080008:0x00080000};
  if(line) options[0x01c0]=colorRef(line);
  if(pib) options[0x4104]=pib;
  const anchor=inGroup?pptRecord(0xf00f,shapeBounds(x,y,width,height)):clientAnchor(x,y,width,height);
  const records=[pptRecord(0xf00a,fsp,2,shapeType),properties(options),anchor];
  if(text) records.push(textRecords(text,{size,color:textColor,align,fontIndex}));
  return pptRecord(0xf004,Buffer.concat(records),15);
}
function tableGroup(id) {
  const header=pptRecord(0xf004,Buffer.concat([
    pptRecord(0xf009,shapeBounds(0,0,360,136),1),pptRecord(0xf00a,Buffer.concat([u32(id),u32(0x0201)]),2),
    properties({0x039f:1},0xf122),clientAnchor(520,340,360,136),
  ]),15);
  const cells=[
    [0,0,210,40,'合并表头','#245aa4'], [210,0,150,40,'说明','#245aa4'],
    [0,40,80,96,'纵向合并','#4472c4'], [80,40,130,52,'绿色单元格','#22aa44'], [210,40,150,52,'','#437bb7'],
    [80,92,130,44,'数据','#4472c4'], [210,92,150,44,'末格','#4472c4'],
  ].map(([x,y,width,height,text,fill],index)=>shape(id+index+1,{x,y,width,height,text,fill,line:null,size:12,inGroup:true}));
  const border=shape(id+8,{x:80,y:40,width:130,height:0,shapeType:20,line:'#331155',inGroup:true});
  return pptRecord(0xf003,Buffer.concat([header,...cells,border]),15);
}
function drawing(drawingId,shapes,nestedShapes=0) {
  const groupBounds=Buffer.alloc(16); groupBounds.writeInt32LE(5760,8); groupBounds.writeInt32LE(3240,12);
  const rootShape=pptRecord(0xf004,Buffer.concat([pptRecord(0xf009,groupBounds,1),pptRecord(0xf00a,Buffer.concat([u32(drawingId*1024),u32(5)]),2)]),15);
  const count=shapes.length+1+nestedShapes;
  const dg=pptRecord(0xf008,Buffer.concat([u32(count),u32(drawingId*1024+count-1)]),0,drawingId);
  return pptRecord(0x040c,pptRecord(0xf002,Buffer.concat([dg,pptRecord(0xf003,Buffer.concat([rootShape,...shapes]),15)]),15),15);
}
function scheme(background) { return pptRecord(0x07f0,Buffer.concat([background,'#184c4c','#777777','#184c4c','#dce9f5','#d56d35','#437bb7','#2f8668'].map(color=>u32(colorRef(color)))),0,1); }
function slideAtom(masterId,flags=0) { const bytes=Buffer.alloc(24); bytes.writeUInt32LE(16); bytes.writeUInt32LE(masterId>>>0,12); bytes.writeUInt16LE(flags,20); return pptRecord(0x03ef,bytes,2); }
function persistAtom(id,slideId) { const bytes=Buffer.alloc(20); bytes.writeUInt32LE(id); if(slideId<0x80000000)bytes.writeUInt32LE(4,4); bytes.writeUInt32LE(slideId>>>0,12); return pptRecord(0x03f3,bytes); }
function pngChunk(type,data) { const name=Buffer.from(type); const size=Buffer.alloc(4); size.writeUInt32BE(data.length); const crc=Buffer.alloc(4); crc.writeUInt32BE(crc32(Buffer.concat([name,data]))); return Buffer.concat([size,name,data,crc]); }
function imageUid(input) {
  // MS-ODRAW BLIP identifiers use an MD4 digest. Implement the small fixture
  // digest here rather than enabling obsolete crypto providers on the host.
  const padded=Buffer.alloc(Math.ceil((input.length+9)/64)*64); input.copy(padded); padded[input.length]=0x80;
  padded.writeBigUInt64LE(BigInt(input.length)*8n,padded.length-8);
  const state=[0x67452301,0xefcdab89,0x98badcfe,0x10325476];
  const orders=[Array.from({length:16},(_,i)=>i),Array.from({length:16},(_,i)=>(i%4)*4+Math.floor(i/4)),[0,8,4,12,2,10,6,14,1,9,5,13,3,11,7,15]];
  const shifts=[[3,7,11,19],[3,5,9,13],[3,9,11,15]],constants=[0,0x5a827999,0x6ed9eba1];
  for(let offset=0;offset<padded.length;offset+=64){
    const words=Array.from({length:16},(_,i)=>padded.readUInt32LE(offset+i*4)),values=[...state];
    for(let round=0;round<3;round++)for(let i=0;i<16;i++){
      const target=(4-i%4)%4,b=values[(target+1)%4],c=values[(target+2)%4],d=values[(target+3)%4];
      const mixed=round===0?(b&c)|(~b&d):round===1?(b&c)|(b&d)|(c&d):b^c^d;
      const value=(values[target]+mixed+words[orders[round][i]]+constants[round])>>>0,shift=shifts[round][i%4];
      values[target]=((value<<shift)|(value>>>(32-shift)))>>>0;
    }
    for(let i=0;i<4;i++)state[i]=(state[i]+values[i])>>>0;
  }
  return Buffer.concat(state.map(u32));
}
function samplePng() {
  const header=Buffer.alloc(13); header.writeUInt32BE(4); header.writeUInt32BE(4,4); header[8]=8; header[9]=2;
  const rows=[];
  for(let y=0;y<4;y++) { const row=Buffer.alloc(13); for(let x=0;x<4;x++) Buffer.from((x+y)%2 ? 'd56d35':'184c4c','hex').copy(row,1+x*3); rows.push(row); }
  return Buffer.concat([Buffer.from('89504e470d0a1a0a','hex'),pngChunk('IHDR',header),pngChunk('IDAT',deflateSync(Buffer.concat(rows))),pngChunk('IEND',Buffer.alloc(0))]);
}

export function buildStyledPpt({embeddedBlip=false}={}) {
  const png=samplePng();
  const uid=imageUid(png);
  const blip=pptRecord(0xf01e,Buffer.concat([uid,Buffer.from([0xff]),png]),0,0x06e0);
  const bse=Buffer.alloc(36); bse[0]=bse[1]=6; uid.copy(bse,2); bse.writeUInt16LE(0xff,18);
  bse.writeUInt32LE(blip.length,20); bse.writeUInt32LE(1,24); bse.writeUInt32LE(embeddedBlip?0xffffffff:0,28);
  const dgg=Buffer.alloc(40); dgg.writeUInt32LE(3074); dgg.writeUInt32LE(4,4); dgg.writeUInt32LE(18,8); dgg.writeUInt32LE(3,12);
  dgg.writeUInt32LE(1,16); dgg.writeUInt32LE(13,20); dgg.writeUInt32LE(2,24); dgg.writeUInt32LE(3,28);
  dgg.writeUInt32LE(3,32); dgg.writeUInt32LE(2,36);
  const drawingGroup=pptRecord(0x040b,pptRecord(0xf000,Buffer.concat([pptRecord(0xf006,dgg),pptRecord(0xf001,pptRecord(0xf007,embeddedBlip?Buffer.concat([bse,blip]):bse,2,6),15,1)]),15),15);
  const fontCollection=pptRecord(0x07d5,Buffer.concat(['微软雅黑','宋体'].map((name,index)=>{
    const data=Buffer.alloc(68); utf16(name).copy(data); data[64]=134;
    return pptRecord(0x0fb7,data,0,index);
  })),15);
  const docAtom=Buffer.alloc(40); docAtom.writeInt32LE(5760); docAtom.writeInt32LE(3240,4); docAtom.writeInt32LE(4320,8); docAtom.writeInt32LE(5760,12);
  docAtom.writeInt32LE(1,16); docAtom.writeInt32LE(1,20); docAtom.writeUInt16LE(1,32); docAtom.writeUInt16LE(7,34);
  const document=pptRecord(0x03e8,Buffer.concat([
    pptRecord(0x03e9,docAtom,1),pptRecord(0x03f2,fontCollection,15),drawingGroup,
    pptRecord(0x0ff0,persistAtom(4,0x80000000),15,1),
    pptRecord(0x0ff0,Buffer.concat([persistAtom(2,256),persistAtom(3,257)]),15),pptRecord(0x03ea,Buffer.alloc(0))
  ]),15);
  const first=pptRecord(0x03ee,Buffer.concat([slideAtom(0x80000000,7),drawing(1,[
    shape(1025,{x:60,y:52,width:840,height:100,text:'旧版 PPT 图片与颜色'}),
    shape(1026,{x:70,y:200,width:400,height:240,pib:1,shapeType:75,fill:'#ffffff',line:'#dce9f5'}),
    shape(1027,{x:520,y:200,width:360,height:130,shapeType:202,fill:'#f3f7fc',line:'#f3f7fc',text:'图片、文字和形状\r按原始位置显示',size:18,textColor:'#184c4c',align:0}),
    tableGroup(1028)
  ],8)]),15);
  const second=pptRecord(0x03ee,Buffer.concat([slideAtom(0x80000000),scheme('#fff3e9'),drawing(2,[
    shape(2049,{x:100,y:100,width:760,height:220,fill:'#d56d35',line:'#d56d35',text:'第二页仍保留颜色和布局',size:28}),
    shape(2050,{x:150,y:360,width:660,height:70,shapeType:202,fill:'#fff3e9',line:'#fff3e9',text:'翻页不会丢失原始绘图数据',size:18,textColor:'#184c4c',align:0,fontIndex:1})
  ])]),15);
  const master=pptRecord(0x03f8,Buffer.concat([slideAtom(0),scheme('#f3f7fc'),drawing(3,[shape(3073,{x:0,y:500,width:960,height:40,fill:'#dce9f5',line:'#dce9f5'})])]),15);
  const offsets=[0,document.length,document.length+first.length,document.length+first.length+second.length];
  const directoryOffset=offsets[3]+master.length;
  const directory=pptRecord(0x1772,Buffer.concat([u32((4<<20)|1),...offsets.map(u32)]));
  const editOffset=directoryOffset+directory.length,edit=Buffer.alloc(28);
  edit.writeUInt32LE(257); edit[7]=3; edit.writeUInt32LE(directoryOffset,12); edit.writeUInt32LE(1,16); edit.writeUInt32LE(5,20); edit.writeUInt16LE(1,24);
  const username=Buffer.from('Developer'),current=Buffer.alloc(20);
  current.writeUInt32LE(20); current.writeUInt32LE(0xe391c05f,4); current.writeUInt32LE(editOffset,8);
  current.writeUInt16LE(username.length,12); current.writeUInt16LE(0x03f4,14); current[16]=3;
  return makeCompound({
    'PowerPoint Document':Buffer.concat([document,first,second,master,directory,pptRecord(0x0ff5,edit)]),
    'Current User':pptRecord(0x0ff6,Buffer.concat([current,username,u32(8),utf16('Developer')])),
    ...(embeddedBlip?{}:{Pictures:blip}),
  });
}

export async function generateStyledFixtures(destination=new URL('../fixtures/',import.meta.url)) {
  await mkdir(destination,{recursive:true});
  const files={'styled.xls':buildStyledXls(),'hidden.xls':buildStyledXls({hidden:true}),'styled.ppt':buildStyledPpt()};
  await Promise.all(Object.entries(files).map(([name,bytes])=>writeFile(new URL(name,destination),bytes)));
  return Object.entries(files).map(([name,bytes])=>({name,bytes:bytes.length}));
}
if(process.argv[1] && path.resolve(process.argv[1])===fileURLToPath(import.meta.url)) console.log(JSON.stringify(await generateStyledFixtures(),null,2));
