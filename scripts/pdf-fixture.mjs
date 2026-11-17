import { deflateSync } from 'node:zlib';
import { writeFile, mkdir } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

// Original development PDFs; no PDF writer or Office engine is used.
// They test this implementation's declared subset, not arbitrary real-world PDFs.
export function pdfStream(dictionary, input, compressed = true) {
  const bytes = Buffer.from(input);
  const stream = compressed ? deflateSync(bytes) : bytes;
  return Buffer.concat([Buffer.from(`<< ${dictionary} /Length ${stream.length}${compressed ? ' /Filter /FlateDecode' : ''} >>\nstream\n`), stream, Buffer.from('\nendstream')]);
}

export function assemblePdf(sourceObjects, options = {}) {
  const objects = new Map(Object.entries(sourceObjects).map(([id, value]) => [Number(id), Buffer.isBuffer(value) ? value : Buffer.from(value)]));
  const compressed = new Map();
  let maximum = Math.max(...objects.keys());
  if (options.objectStream) {
    const id = ++maximum;
    const selected = [1, 2, 3, 4].filter(value => objects.has(value));
    const content = [];
    let offset = 0;
    let header = '';
    selected.forEach((objectId, index) => {
      header += `${objectId} ${offset} `;
      const value = objects.get(objectId);
      content.push(value, Buffer.from('\n'));
      offset += value.length + 1;
      compressed.set(objectId, { streamId: id, index });
      objects.delete(objectId);
    });
    objects.set(id, pdfStream(`/Type /ObjStm /N ${selected.length} /First ${Buffer.byteLength(header)}`, Buffer.concat([Buffer.from(header), ...content])));
  }
  if (options.objectStream && !options.xrefStream) throw new Error('Object stream fixtures require xrefStream.');
  const chunks = [Buffer.from('%PDF-1.7\n%\xe2\xe3\xcf\xd3\n', 'latin1')];
  const offsets = new Map();
  let offset = chunks[0].length;
  for (const [id, body] of objects) {
    offsets.set(id, offset);
    const object = Buffer.concat([Buffer.from(`${id} 0 obj\n`), body, Buffer.from('\nendobj\n')]);
    chunks.push(object); offset += object.length;
  }
  const startxref = offset;
  const extra = options.encrypted ? ' /Encrypt 99 0 R' : '';
  if (options.xrefStream) {
    const xrefId = maximum + 1;
    const size = xrefId + 1;
    const entries = Buffer.alloc(size * 7);
    for (let id = 0; id < size; id += 1) {
      const position = id * 7;
      const item = compressed.get(id);
      entries[position] = item ? 2 : id === xrefId || offsets.has(id) ? 1 : 0;
      entries.writeUInt32BE(item ? item.streamId : id === xrefId ? startxref : offsets.get(id) || 0, position + 1);
      entries.writeUInt16BE(item ? item.index : id === 0 ? 65535 : 0, position + 5);
    }
    let encodedEntries = entries;
    if (options.xrefPredictor) {
      encodedEntries = Buffer.alloc(size * 8);
      for (let row = 0; row < size; row += 1) {
        encodedEntries[row * 8] = 2;
        for (let col = 0; col < 7; col += 1) encodedEntries[row * 8 + col + 1] = (entries[row * 7 + col] - (row ? entries[(row - 1) * 7 + col] : 0)) & 255;
      }
    }
    const body = pdfStream(`/Type /XRef /W [1 4 2] /Size ${size} /Root 1 0 R${extra}${options.xrefPredictor ? ' /DecodeParms << /Predictor 12 /Columns 7 >>' : ''}`, encodedEntries);
    chunks.push(Buffer.from(`${xrefId} 0 obj\n`), body, Buffer.from(`\nendobj\nstartxref\n${startxref}\n%%EOF\n`));
  } else {
    let xref = `xref\n0 ${maximum + 1}\n0000000000 65535 f \n`;
    for (let id = 1; id <= maximum; id += 1) xref += `${String(offsets.get(id) || 0).padStart(10, '0')} 00000 ${offsets.has(id) ? 'n' : 'f'} \n`;
    chunks.push(Buffer.from(`${xref}trailer\n<< /Size ${maximum + 1} /Root 1 0 R${extra} >>\nstartxref\n${startxref}\n%%EOF\n`));
  }
  return new Uint8Array(Buffer.concat(chunks));
}

export function buildPdfFixture(options = {}) {
  const compress = options.compressed !== false;
  const cmap = `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /SelfAuthored-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <ffff>\nendcodespacerange\n4 beginbfchar\n<0001> <81ea>\n<0002> <4e3b>\n<0003> <9884>\n<0004> <89c8>\nendbfchar\n1 beginbfrange\n<0005> <0006> [<4e2d> <6587>]\nendbfrange\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`;
  const firstContent = options.firstContent ?? `q\n0.965 0.953 0.925 rg\n0 0 612 792 re f\n0.08 0.30 0.29 rg\n48 660 516 60 re f\nBT /F1 24 Tf 1 0 0 1 64 682 Tm 1 1 1 rg (Self-developed PDF preview) Tj ET\nBT /F1 16 Tf 0.15 0.20 0.19 rg 1 0 0 1 48 620 Tm (Real page tree, text and vector paths.) Tj ET\nBT /F2 32 Tf 0.08 0.30 0.29 rg 1 0 0 1 48 554 Tm <0001000200030004> Tj ET\nBT /F1 13 Tf 1 0 0 1 48 510 Tm [(Local) -80 ( file parsing / no external PDF engine)] TJ ET\n0.70 0.30 0.16 RG 3 w 48 476 m 564 476 l S\nq 40 0 0 40 48 410 cm /Im1 Do Q\nq 40 0 0 40 100 410 cm /Im2 Do Q\n0.08 0.30 0.29 RG 2 w 160 412 m 240 458 320 388 400 430 c S\nQ\n`;
  const secondContent = `q\n0.98 0.98 0.97 rg 0 0 612 792 re f\nBT /F1 25 Tf 0.08 0.30 0.29 rg 1 0 0 1 48 680 Tm (Second page / native drawing) Tj ET\nBT /F2 28 Tf 1 0 0 1 48 610 Tm <00050006> Tj ET\n0.76 0.36 0.20 rg 48 460 220 70 re f\n0.08 0.30 0.29 RG 2 w 300 460 220 70 re S\nBT /F1 13 Tf 1 0 0 1 48 400 Tm (Fonts are substituted; advanced PDF features remain limited.) Tj ET\nQ\n`;
  const pageOptions = `${options.crop ? ' /CropBox [50 50 550 750]' : ''}${options.rotate ? ` /Rotate ${options.rotate}` : ''}${options.userUnit ? ` /UserUnit ${options.userUnit}` : ''}`;
  const objects = {
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R 4 0 R] /Count 2 /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R /F2 6 0 R >> /XObject << /Im1 10 0 R /Im2 14 0 R >> >> >>',
    3: '<< /Type /Page /Parent 2 0 R /Contents 8 0 R >>',
    4: `<< /Type /Page /Parent 2 0 R /Contents 9 0 R${pageOptions} >>`,
    5: `<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding /FirstChar 32 /LastChar 126 /Widths [${Array.from({ length: 95 }, (_, index) => index === 0 ? 250 : 500).join(' ')}] >>`,
    6: '<< /Type /Font /Subtype /Type0 /BaseFont /HeitiSC /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 11 0 R >>',
    7: '<< /Type /Font /Subtype /CIDFontType0 /BaseFont /HeitiSC /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor 13 0 R /DW 1000 /W [1 [1000 1000 1000 1000 1000 1000]] >>',
    8: pdfStream('', Buffer.from(firstContent), compress),
    9: pdfStream('', Buffer.from(secondContent), compress),
    10: pdfStream('/Type /XObject /Subtype /Image /Width 2 /Height 2 /ColorSpace /DeviceRGB /BitsPerComponent 8', Buffer.from([210, 55, 30, 30, 150, 90, 30, 70, 190, 220, 185, 45]), compress),
    11: pdfStream('', Buffer.from(cmap), compress),
    13: '<< /Type /FontDescriptor /FontName /HeitiSC /Flags 4 /FontBBox [-50 -200 1000 1000] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 880 /StemV 80 >>',
  };
  // One green pixel, created locally with the OS JPEG encoder. Stored to keep fixture generation portable.
  const jpeg = Buffer.from('/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAABAAEDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDlaKKKg+VP/9k=', 'base64');
  objects[14] = pdfStream('/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode', jpeg, false);
  return assemblePdf(objects, options);
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const destination = new URL('../fixtures/sample.pdf', import.meta.url);
  await mkdir(new URL('../fixtures/', import.meta.url), { recursive: true });
  const bytes = buildPdfFixture();
  await writeFile(destination, bytes);
  console.log(`Generated ${fileURLToPath(destination)} (${bytes.length} bytes, 2 pages).`);
}
