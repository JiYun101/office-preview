import test from 'node:test';
import assert from 'node:assert/strict';
import { parsePdf } from '../src/parsers/pdf.js';
import { buildPdfFixture, assemblePdf, pdfStream } from '../scripts/pdf-fixture.mjs';

const text = page => page.operations.filter(operation => operation.type === 'text').map(operation => operation.text).join('');

test('self-authored PDF parses actual page trees, Flate content, Chinese ToUnicode, paths and pixels', async () => {
  const model = await parsePdf(buildPdfFixture(), 'sample.pdf');
  assert.equal(model.kind, 'pdf');
  assert.equal(model.name, 'sample.pdf');
  assert.equal(model.pages.length, 2);
  assert.equal(model.pages[0].width, 612);
  assert.equal(model.pages[0].height, 792);
  assert.match(text(model.pages[0]), /Self-developed PDF preview/);
  assert.match(text(model.pages[0]), /自主预览/);
  assert.match(text(model.pages[1]), /中文/);
  assert.ok(model.pages[0].operations.some(operation => operation.type === 'path' && operation.fill === '#144d4a'));
  assert.ok(model.pages[0].operations.some(operation => operation.type === 'path' && operation.commands.some(command => command[0] === 'C')));
  const image = model.pages[0].operations.find(operation => operation.type === 'image');
  assert.equal(image.pixelWidth, 2);
  assert.deepEqual(Array.from(image.rgba.slice(0, 4)), [210, 55, 30, 255]);
  const jpeg = model.pages[0].operations.find(operation => operation.type === 'image' && operation.mime === 'image/jpeg');
  assert.equal(jpeg.bytes[0], 0xff);
  assert.equal(jpeg.bytes[1], 0xd8);
  assert.ok(model.warnings.some(value => value.includes('系统字体')));
});

test('PDF uncompressed streams, literal escapes, TJ positioning, T* and clipping retain real instructions', async () => {
  const content = `q 10 10 200 100 re W n BT /F1 20 Tf 24 TL 1 0 0 1 30 300 Tm (A\\(B\\)\\101) Tj [(C) 100 (D)] TJ T* (line 2) Tj ET Q`;
  const model = await parsePdf(buildPdfFixture({ compressed: false, firstContent: content }));
  const page = model.pages[0];
  assert.equal(text(page), 'A(B)ACDline 2');
  assert.ok(page.operations.some(operation => operation.type === 'clip'));
  const letters = page.operations.filter(operation => operation.type === 'text');
  assert.equal(letters[0].matrix[4], 30);
  assert.equal(letters.at(-1).matrix[5], 276);
  const c = letters.find(operation => operation.text === 'C');
  const d = letters.find(operation => operation.text === 'D');
  assert.equal(d.matrix[4] - c.matrix[4], 8);
});

test('PDF CropBox, inherited MediaBox, page rotation and UserUnit become explicit transforms', async () => {
  const model = await parsePdf(buildPdfFixture({ crop: true, rotate: 90, userUnit: 2 }));
  const page = model.pages[1];
  assert.equal(page.width, 1400);
  assert.equal(page.height, 1000);
  assert.deepEqual(page.operations[1].matrix, [0, -2, 2, 0, -100, 1100]);
});

test('PDF xref streams and compressed ObjStm objects follow indices rather than binary scans', async () => {
  const plain = await parsePdf(buildPdfFixture({ xrefStream: true }));
  const compressed = await parsePdf(buildPdfFixture({ xrefStream: true, objectStream: true, xrefPredictor: true }));
  assert.equal(plain.pages.length, 2);
  assert.equal(compressed.pages.length, 2);
  assert.match(text(compressed.pages[0]), /自主预览/);
  assert.equal(text(plain.pages[1]), text(compressed.pages[1]));
});

test('PDF compressed object parsing cannot borrow the next indexed object to close a truncated dictionary', async () => {
  const malformed = assemblePdf({
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [5 0 R] /Count 1 /MediaBox [0 0 200 200] >>',
    3: '<< /Type /Font',
    4: '/Subtype /Type1 /BaseFont /Helvetica >>',
    5: '<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 3 0 R >> >> /Contents 6 0 R >>',
    6: pdfStream('', Buffer.from('BT /F1 12 Tf 1 0 0 1 10 50 Tm (Cross-object leak) Tj ET')),
  }, { xrefStream: true, objectStream: true });
  await assert.rejects(parsePdf(malformed), /PDF.*结束|PDF.*无效|PDF.*对象流/);
});

test('PDF incremental xref uses the latest object revision and follows Prev for older objects', async () => {
  const original = Buffer.from(buildPdfFixture());
  const previous = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(original.toString('latin1'))[1]);
  const object = Buffer.concat([Buffer.from('8 0 obj\n'), pdfStream('', Buffer.from('BT /F1 12 Tf (Latest revision) Tj ET')), Buffer.from('\nendobj\n')]);
  const xrefOffset = original.length + object.length;
  const suffix = Buffer.from(`xref\n8 1\n${String(original.length).padStart(10, '0')} 00000 n \ntrailer\n<< /Size 15 /Prev ${previous} >>\nstartxref\n${xrefOffset}\n%%EOF\n`);
  const model = await parsePdf(Buffer.concat([original, object, suffix]));
  assert.equal(text(model.pages[0]), 'Latest revision');
  assert.match(text(model.pages[1]), /Second page/);
  const freed = Buffer.from(`xref\n8 1\n0000000000 00001 f \ntrailer\n<< /Size 15 /Prev ${previous} >>\nstartxref\n${original.length}\n%%EOF\n`);
  await assert.rejects(parsePdf(Buffer.concat([original, freed])), /对象引用无效/);
});

test('PDF Form XObjects use their resources, transforms and bounds', async () => {
  const model = await parsePdf(assemblePdf({
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 200 200] >>',
    3: '<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Form1 5 0 R >> >> /Contents 4 0 R >>',
    4: pdfStream('', Buffer.from('/Form1 Do')),
    5: pdfStream('/Type /XObject /Subtype /Form /BBox [0 0 50 50] /Matrix [2 0 0 2 10 20] /Resources << >>', Buffer.from('1 0 0 rg 0 0 50 50 re f')),
  }));
  assert.ok(model.pages[0].operations.some(operation => operation.type === 'transform' && operation.matrix[4] === 10));
  assert.ok(model.pages[0].operations.some(operation => operation.type === 'clip'));
  assert.ok(model.pages[0].operations.some(operation => operation.type === 'path' && operation.fill === '#ff0000'));
});

test('PDF Flate PNG predictors reconstruct image bytes rather than exposing compressed pixel data', async () => {
  const model = await parsePdf(assemblePdf({
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>',
    3: '<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>',
    4: pdfStream('', Buffer.from('/Im1 Do /Im1 Do')),
    5: pdfStream('/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /DecodeParms << /Predictor 15 /Colors 3 /BitsPerComponent 8 /Columns 2 >>', Buffer.from([1, 10, 20, 30, 30, 30, 30, 2, 10, 20, 30, 40, 50, 60])),
  }));
  const images = model.pages[0].operations.filter(operation => operation.type === 'image');
  assert.deepEqual(Array.from(images[0].rgba), [10, 20, 30, 255, 40, 50, 60, 255, 20, 40, 60, 255, 80, 100, 120, 255]);
  assert.equal(images[0].rgba, images[1].rgba, 'Repeated image resources share one decoded pixel buffer.');
});

test('PDF ASCIIHex and ASCII85 stream filters reconstruct the drawing program', async () => {
  const program = Buffer.concat([Buffer.alloc(4), Buffer.from('0 0 1 rg 0 0 10 10 re f')]);
  function encode85(input) {
    let result = '';
    for (let index = 0; index < input.length; index += 4) {
      const count = Math.min(4, input.length - index);
      let value = 0;
      for (let byte = 0; byte < 4; byte += 1) value = value * 256 + (input[index + byte] || 0);
      if (count === 4 && value === 0) { result += 'z'; continue; }
      const chars = [];
      for (let digit = 0; digit < 5; digit += 1) { chars.unshift(String.fromCharCode(value % 85 + 33)); value = Math.floor(value / 85); }
      result += chars.slice(0, count + 1).join('');
    }
    return result + '~>';
  }
  for (const [filter, encoded] of [['ASCIIHexDecode', program.toString('hex') + '>'], ['ASCII85Decode', encode85(program)]]) {
    const model = await parsePdf(assemblePdf({
      1: '<< /Type /Catalog /Pages 2 0 R >>',
      2: '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>',
      3: '<< /Type /Page /Parent 2 0 R /Resources << >> /Contents 4 0 R >>',
      4: pdfStream(`/Filter /${filter}`, Buffer.from(encoded), false),
    }));
    assert.ok(model.pages[0].operations.some(operation => operation.type === 'path' && operation.fill === '#0000ff'));
  }
});

test('PDF unknown newest xref entry types cannot fall back to an older object version', async () => {
  const original = Buffer.from(buildPdfFixture());
  const previous = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(original.toString('latin1'))[1]);
  const entries = Buffer.from([3, 0, 0, 0, 0, 0, 0]);
  const body = pdfStream(`/Type /XRef /W [1 4 2] /Index [8 1] /Size 16 /Prev ${previous}`, entries);
  const update = Buffer.concat([Buffer.from('15 0 obj\n'), body, Buffer.from(`\nendobj\nstartxref\n${original.length}\n%%EOF\n`)]);
  await assert.rejects(parsePdf(Buffer.concat([original, update])), /对象引用无效/);
});

test('PDF an older hybrid xref supplement cannot revive an object freed by the newest revision', async () => {
  const base = Buffer.from(buildPdfFixture({ compressed: false }));
  const original = base.toString('latin1');
  const baseXref = Number(/startxref\n(\d+)\n%%EOF\n$/.exec(original)[1]);
  function entry(type, position, generation = 0) {
    const bytes = Buffer.alloc(7); bytes[0] = type; bytes.writeUInt32BE(position, 1); bytes.writeUInt16BE(generation, 5); return bytes;
  }
  const supplement = Buffer.concat([
    Buffer.from('15 0 obj\n'),
    pdfStream('/Type /XRef /Size 16 /Index [8 1 15 1] /W [1 4 2] /Root 1 0 R', Buffer.concat([entry(1, original.indexOf('8 0 obj\n')), entry(1, baseXref)]), false),
    Buffer.from('\nendobj\n'),
  ]);
  const hybridXref = baseXref + supplement.length;
  const traditional = original.slice(baseXref, original.lastIndexOf('startxref'))
    .replace('/Size 15', '/Size 16')
    .replace('/Root 1 0 R', `/Root 1 0 R /XRefStm ${baseXref}`)
    .replace('trailer\n', `15 1\n${String(baseXref).padStart(10, '0')} 00000 n \ntrailer\n`);
  const older = Buffer.concat([base.subarray(0, baseXref), supplement, Buffer.from(`${traditional}startxref\n${hybridXref}\n%%EOF\n`)]);
  const newest = Buffer.from(`xref\n8 1\n0000000000 00001 f \ntrailer\n<< /Size 16 /Root 1 0 R /Prev ${hybridXref} >>\nstartxref\n${older.length}\n%%EOF\n`);
  const validHybrid = await parsePdf(older);
  assert.match(text(validHybrid.pages[0]), /Self-developed PDF preview/);
  await assert.rejects(parsePdf(Buffer.concat([older, newest])), /对象引用无效/);
});

test('PDF JPEG dimensions are checked from the SOF header before handing bytes to a browser decoder', async () => {
  const source = await parsePdf(buildPdfFixture());
  const jpeg = source.pages[0].operations.find(operation => operation.type === 'image' && operation.mime === 'image/jpeg').bytes;
  const contradictory = assemblePdf({
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 100 100] >>',
    3: '<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>',
    4: pdfStream('', Buffer.from('/Im1 Do')),
    5: pdfStream('/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode', jpeg, false),
  });
  await assert.rejects(parsePdf(contradictory), /实际尺寸/);
});

test('PDF encrypted, malformed, fake binary text and unsupported inline image files fail explicitly', async () => {
  await assert.rejects(parsePdf(buildPdfFixture({ encrypted: true })), /加密/);
  const broken = buildPdfFixture().slice(0, -30);
  await assert.rejects(parsePdf(broken), /startxref|EOF/);
  await assert.rejects(parsePdf(new TextEncoder().encode('%PDF-1.7\nBT (fake text) Tj ET\n%%EOF')), /startxref/);
  await assert.rejects(parsePdf(buildPdfFixture({ firstContent: 'BI /W 1 /H 1 ID abc EI' })), /内联图像/);
  const badOffset = Buffer.from(buildPdfFixture());
  const start = badOffset.toString('latin1').lastIndexOf('startxref\n');
  badOffset.write('99999999', start + 10, 'ascii');
  await assert.rejects(parsePdf(badOffset), /xref|偏移|EOF/);
});
