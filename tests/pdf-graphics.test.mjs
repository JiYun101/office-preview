import test from 'node:test';
import assert from 'node:assert/strict';
import { parsePdf } from '../src/parsers/pdf.js';
import { PdfContent } from '../src/parsers/pdf-content.js';
import { assemblePdf, pdfStream } from '../scripts/pdf-fixture.mjs';

function graphicPdf(content, states = '') {
  return assemblePdf({
    1: '<< /Type /Catalog /Pages 2 0 R >>',
    2: '<< /Type /Pages /Kids [3 0 R] /Count 1 /MediaBox [0 0 200 200] >>',
    3: `<< /Type /Page /Parent 2 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /Im1 6 0 R >> /ExtGState << ${states} >> >> /Contents 4 0 R >>`,
    4: pdfStream('', Buffer.from(content)),
    5: '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-BoldOblique /FirstChar 65 /LastChar 68 /Widths [500 500 500 500] >>',
    6: pdfStream('/Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceRGB /BitsPerComponent 8', Buffer.from([12, 34, 56])),
  });
}

test('PDF constant alpha applies to paths, text and images and restores across q/Q', async () => {
  const model = await parsePdf(graphicPdf(`
    0 0.6 0.7 rg 0.2 0.3 0.4 RG
    /A gs 10 10 30 20 re B
    q /B gs 10 40 30 20 re B /Im1 Do Q
    BT /F1 16 Tf 1 0 0 1 20 90 Tm (A) Tj ET
    10 80 30 20 re B /Im1 Do`, '/A << /ca 0.9 /CA 0.4 >> /B << /ca 0.25 /CA 0.1 >>'));
  const operations = model.pages[0].operations;
  const paths = operations.filter(operation => operation.type === 'path');
  assert.deepEqual(paths.map(operation => [operation.fillOpacity, operation.strokeOpacity]), [[0.9, 0.4], [0.25, 0.1], [0.9, 0.4]]);
  assert.equal(paths[0].fill, '#0099b3', 'Alpha remains separate from the source color for correct compositing against any backdrop.');
  const text = operations.find(operation => operation.type === 'text');
  assert.equal(text.fillOpacity, 0.9);
  assert.equal(text.strokeOpacity, 0.4);
  assert.deepEqual(operations.filter(operation => operation.type === 'image').map(operation => operation.fillOpacity), [0.25, 0.9]);
  assert.ok(!model.warnings.some(warning => /软蒙版|特殊混合模式/.test(warning)), 'Supported constant alpha does not issue an unsupported feature warning.');
});

test('PDF line cap, join, miter and dash preserve operator and ExtGState values without aliasing saved state', async () => {
  const model = await parsePdf(graphicPdf(`
    2 w 1 J 2 j 7 M [3 2] 1 d
    10 10 m 50 10 l S
    q /Stroke gs 10 20 m 50 20 l S Q
    10 30 m 50 30 l S`, '/Stroke << /LW 4 /LC 2 /LJ 1 /ML 12 /D [[6 3 2] 4] >>'));
  const paths = model.pages[0].operations.filter(operation => operation.type === 'path');
  const styles = paths.map(({ lineWidth, lineCap, lineJoin, miterLimit, dash, dashPhase }) => ({ lineWidth, lineCap, lineJoin, miterLimit, dash, dashPhase }));
  assert.deepEqual(styles[0], { lineWidth: 2, lineCap: 1, lineJoin: 2, miterLimit: 7, dash: [3, 2], dashPhase: 1 });
  assert.deepEqual(styles[1], { lineWidth: 4, lineCap: 2, lineJoin: 1, miterLimit: 12, dash: [6, 3, 2], dashPhase: 4 });
  assert.deepEqual(styles[2], styles[0]);
});

test('PDF text fill, stroke and invisible modes retain colors, font style and advances', async () => {
  const model = await parsePdf(graphicPdf('0.8 0.2 0.1 rg 0.1 0.2 0.8 RG 0.6 w BT /F1 20 Tf 1 0 0 1 20 80 Tm (A) Tj 1 Tr (B) Tj 2 Tr (C) Tj 3 Tr (D) Tj ET'));
  const text = model.pages[0].operations.filter(operation => operation.type === 'text');
  assert.deepEqual(text.map(operation => operation.textMode), [0, 1, 2, 3]);
  assert.deepEqual(text.map(operation => operation.matrix[4]), [20, 30, 40, 50]);
  for (const operation of text) {
    assert.equal(operation.color, '#cc331a');
    assert.equal(operation.stroke, '#1a33cc');
    assert.equal(operation.lineWidth, 0.6);
    assert.equal(operation.fontWeight, 700);
    assert.equal(operation.fontStyle, 'italic');
    assert.equal(operation.width, 10, 'PDF widths describe advances, independently of measured glyph ink bounds.');
  }
  assert.ok(!model.warnings.some(warning => /描边文字/.test(warning)));
  assert.equal(model.fonts.length, 1);
  assert.equal(model.fonts[0].weight, 700);
  assert.equal(model.fonts[0].style, 'italic');
  assert.deepEqual(structuredClone(model), model, 'The worker model contains no font decoder functions or Maps.');
});

test('PDF glyph operations use embedded font IDs only when the glyph has a valid render mapping', async () => {
  const operations = [];
  const warnings = [];
  const document = { budget: { tokens: 0 }, warn: warning => warnings.push(warning) };
  const parser = new PdfContent(document, operations);
  parser.state.font = {
    id: 'embedded-synthetic', family: 'embedded-synthetic', fallbackFamily: 'serif', weight: 600, style: 'normal',
    program: { bytes: new Uint8Array([1, 2]), format: 'truetype' },
    decode: () => [
      { text: 'A', code: 65, width: 500, renderText: '\ue001' },
      { text: 'B', code: 66, width: 600 },
    ],
  };
  await parser.run(Buffer.from('BT (AB) Tj ET'), {});
  parser.finish();
  assert.equal(operations[0].fontId, 'embedded-synthetic');
  assert.equal(operations[0].renderText, '\ue001');
  assert.equal(operations[0].text, 'A');
  assert.equal(operations[1].fontFamily, 'serif');
  assert.equal(operations[1].fontId, undefined);
  assert.equal(operations[1].renderText, undefined);
  assert.equal(operations[1].matrix[4], 6);
});

test('PDF ExtGState font entries share the regular font cache and restore with graphics state', async () => {
  const model = await parsePdf(graphicPdf('BT /F1 10 Tf q /F gs (A) Tj Q (B) Tj ET', '/F << /Font [5 0 R 18] >>'));
  assert.deepEqual(model.pages[0].operations.filter(operation => operation.type === 'text').map(operation => operation.fontSize), [18, 10]);
  assert.equal(model.fonts.length, 1);
});

test('PDF unsupported glyph clipping remains explicit without discarding text rendering mode', async () => {
  const model = await parsePdf(graphicPdf('BT /F1 12 Tf 4 Tr (A) Tj 7 Tr (B) Tj ET'));
  assert.deepEqual(model.pages[0].operations.filter(operation => operation.type === 'text').map(operation => operation.textMode), [4, 7]);
  assert.ok(model.warnings.some(warning => warning.includes('文字裁剪尚未应用')));
  const masked = await parsePdf(graphicPdf('/A gs 10 10 20 20 re f', '/A << /ca 0.5 /BM /Multiply /SMask << >> >>'));
  assert.ok(masked.warnings.some(warning => warning.includes('特殊混合模式')));
});

test('PDF rejects invalid graphics state values instead of silently losing style', async () => {
  await assert.rejects(parsePdf(graphicPdf('/A gs', '/A << /ca 1.1 >>')), /透明度/);
  await assert.rejects(parsePdf(graphicPdf('/A gs', '/A << /LW -1 >>')), /线宽/);
  await assert.rejects(parsePdf(graphicPdf('3 J')), /端点/);
  await assert.rejects(parsePdf(graphicPdf('[0 0] 0 d')), /虚线/);
  await assert.rejects(parsePdf(graphicPdf('/A gs', '/A << /D [[-2 1] 0] >>')), /虚线/);
});
