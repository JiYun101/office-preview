import test from 'node:test';
import assert from 'node:assert/strict';
import { readZip, crc32 } from '../src/containers/zip.js';
import { readCompound } from '../src/containers/cfb.js';
import { makeZip, makeCompound } from '../scripts/generate-fixtures.mjs';

test('ZIP CRC32 matches the published 123456789 check vector', () => {
  assert.equal(crc32(Buffer.from('123456789')), 0xcbf43926);
});

for (const compression of ['store', 'deflate']) {
  test(`ZIP ${compression} reads Unicode names, empty entries and exact bytes`, async () => {
    const binary = Buffer.from([0, 1, 127, 128, 255]);
    const result = await readZip(makeZip({ '中文/正文.xml': '<正文>中文内容</正文>', 'empty.bin': Buffer.alloc(0), 'bytes.bin': binary }, { compression }));
    assert.equal(new TextDecoder().decode(result.get('中文/正文.xml')), '<正文>中文内容</正文>');
    assert.equal(result.get('empty.bin').length, 0);
    assert.deepEqual(Buffer.from(result.get('bytes.bin')), binary);
  });
}

test('ZIP rejects damaged payload instead of rendering corrupted document data', async () => {
  const bytes = makeZip({ 'document.xml': '<text>original text</text>' }, { compression: 'store' });
  bytes[30 + Buffer.byteLength('document.xml') + 7] ^= 1;
  await assert.rejects(readZip(bytes), /校验失败/);
});

test('ZIP rejects a truncated directory and invalid internal paths', async () => {
  const bytes = makeZip({ 'document.xml': 'text' });
  await assert.rejects(readZip(bytes.subarray(0, -4)), /ZIP|损坏/);
  await assert.rejects(readZip(makeZip({ '../escape.xml': 'text' })), /路径/);
});

test('ZIP rejects a later entry whose local header and data are hidden inside a central-directory comment', async () => {
  const directoryName = Buffer.from('folder/');
  const documentName = Buffer.from('word/document.xml');
  const documentData = Buffer.from('<document>inside-central-comment</document>');
  const empty = Buffer.alloc(0);
  const localEntry = (name, data) => {
    const header = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50);
    header.writeUInt32LE(crc32(data), 14);
    header.writeUInt32LE(data.length, 18);
    header.writeUInt32LE(data.length, 22);
    header.writeUInt16LE(name.length, 26);
    return Buffer.concat([header, name, data]);
  };
  const centralEntry = (name, data, offset, comment = empty) => {
    const header = Buffer.alloc(46);
    header.writeUInt32LE(0x02014b50);
    header.writeUInt32LE(crc32(data), 16);
    header.writeUInt32LE(data.length, 20);
    header.writeUInt32LE(data.length, 24);
    header.writeUInt16LE(name.length, 28);
    header.writeUInt16LE(comment.length, 32);
    header.writeUInt32LE(offset, 42);
    return Buffer.concat([header, name, comment]);
  };
  const directoryLocal = localEntry(directoryName, empty);
  const forgedLocal = localEntry(documentName, documentData);
  const centralStart = directoryLocal.length;
  const forgedOffset = centralStart + 46 + directoryName.length;
  const central = Buffer.concat([
    centralEntry(directoryName, empty, 0, forgedLocal),
    centralEntry(documentName, documentData, forgedOffset),
  ]);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50);
  end.writeUInt16LE(2, 8);
  end.writeUInt16LE(2, 10);
  end.writeUInt32LE(central.length, 12);
  end.writeUInt32LE(centralStart, 16);
  // The payload and CRC are valid; it is the directory overlap that must fail.
  await assert.rejects(readZip(Buffer.concat([directoryLocal, central, end])), /条目头|重叠/);
});

test('CFB reads regular FAT, multi-sector mini FAT, empty streams and names', () => {
  const large = Buffer.alloc(7001);
  for (let index = 0; index < large.length; index++) large[index] = index % 251;
  const small = Buffer.from('MiniFAT 中文流：' + '多扇区'.repeat(70));
  const compound = readCompound(makeCompound({ RegularData: large, SmallData: small, EmptyData: Buffer.alloc(0) }));
  assert.deepEqual(new Set(compound.names), new Set(['RegularData', 'SmallData', 'EmptyData']));
  assert.deepEqual(Buffer.from(compound.get('regulardata')), large);
  assert.deepEqual(Buffer.from(compound.get('SmallData')), small);
  assert.equal(compound.get('EmptyData').length, 0);
  assert.equal(compound.get('missing'), undefined);
});

test('CFB rejects a truncated sector and a cyclic regular-stream chain', () => {
  const bytes = makeCompound({ Payload: Buffer.alloc(4096, 42) });
  assert.throws(() => readCompound(bytes.subarray(0, -1)), /完整/);
  const corrupt = Buffer.from(bytes);
  const fatSector = corrupt.readUInt32LE(76);
  corrupt.writeUInt32LE(0, (fatSector + 1) * 512); // stream sector 0 -> 0
  assert.throws(() => readCompound(corrupt).get('Payload'), /循环|损坏/);
});

test('CFB rejects a cyclic mini-stream chain', () => {
  const bytes = makeCompound({ Payload: Buffer.alloc(200, 7) });
  const miniFatSector = bytes.readUInt32LE(60);
  bytes.writeUInt32LE(0, (miniFatSector + 1) * 512);
  assert.throws(() => readCompound(bytes).get('Payload'), /循环|损坏/);
});
