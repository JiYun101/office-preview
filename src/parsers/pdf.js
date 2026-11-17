// Original basic PDF renderer parser. Reference: Adobe's ISO 32000-1:2008,
// https://opensource.adobe.com/dc-acrobat-sdk-docs/standards/pdfstandards/pdf/PDF32000_2008.pdf
import { PdfDocument } from './pdf-document.js';
import { PdfContent } from './pdf-content.js';
import { asName, isDictionary, isReference, numeric } from './pdf-syntax.js';

export async function parsePdf(input, name = '') {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const warnings = [];
  const document = new PdfDocument(bytes, warnings);
  const catalog = await document.load();
  const pages = [];
  const visited = new Set();

  async function box(value, label) {
    const array = await document.resolve(value);
    if (!Array.isArray(array) || array.length !== 4) throw new Error(`PDF ${label} 无效。`);
    const values = await Promise.all(array.map(item => document.resolve(item)));
    values.forEach(item => numeric(item, label));
    if (values[2] <= values[0] || values[3] <= values[1]) throw new Error(`PDF ${label} 尺寸无效。`);
    return values;
  }

  async function visit(reference, inherited = {}, depth = 0) {
    if (depth > 64) throw new Error('PDF 页树嵌套超过当前上限。');
    const key = isReference(reference) ? `${reference.ref}:${reference.generation}` : reference;
    if (visited.has(key)) throw new Error('PDF 页树包含循环或重复页面引用。');
    visited.add(key);
    const node = await document.resolve(reference);
    if (!isDictionary(node)) throw new Error('PDF 页树节点不是字典。');
    const settings = { ...inherited };
    for (const property of ['Resources', 'MediaBox', 'CropBox', 'Rotate']) if (node[property] !== undefined) settings[property] = node[property];
    if (asName(node.Type) === 'Pages') {
      const children = await document.resolve(node.Kids);
      if (!Array.isArray(children) || children.length > 10000) throw new Error('PDF 页树 Kids 无效或过大。');
      for (const child of children) await visit(child, settings, depth + 1);
      return;
    }
    if (asName(node.Type) !== 'Page') throw new Error('PDF 页树包含未知类型。');
    if (pages.length >= 500) throw new Error('PDF 页数超过当前 500 页上限。');
    const media = await box(settings.MediaBox, 'MediaBox');
    const crop = settings.CropBox ? await box(settings.CropBox, 'CropBox') : media;
    const bounds = [Math.max(media[0], crop[0]), Math.max(media[1], crop[1]), Math.min(media[2], crop[2]), Math.min(media[3], crop[3])];
    if (bounds[2] <= bounds[0] || bounds[3] <= bounds[1]) throw new Error('PDF CropBox 与页面没有交集。');
    const rotationValue = numeric(await document.resolve(settings.Rotate), 'Rotate', 0);
    if (!Number.isInteger(rotationValue) || rotationValue % 90) throw new Error('PDF 页面旋转角度必须为 90 的倍数。');
    const rotation = (rotationValue % 360 + 360) % 360;
    const scale = numeric(await document.resolve(node.UserUnit), 'UserUnit', 1);
    if (scale <= 0 || scale > 75000) throw new Error('PDF 页面 UserUnit 无效。');
    const unrotatedWidth = (bounds[2] - bounds[0]) * scale;
    const unrotatedHeight = (bounds[3] - bounds[1]) * scale;
    const width = rotation % 180 ? unrotatedHeight : unrotatedWidth;
    const height = rotation % 180 ? unrotatedWidth : unrotatedHeight;
    if (!Number.isFinite(width + height)) throw new Error('PDF 页面尺寸计算溢出。');
    const matrices = {
      0: [scale, 0, 0, scale, -bounds[0] * scale, -bounds[1] * scale],
      90: [0, -scale, scale, 0, -bounds[1] * scale, bounds[2] * scale],
      180: [-scale, 0, 0, -scale, bounds[2] * scale, bounds[3] * scale],
      270: [0, scale, -scale, 0, bounds[3] * scale, -bounds[0] * scale],
    };
    const operations = [{ type: 'save' }, { type: 'transform', matrix: matrices[rotation] }];
    const parser = new PdfContent(document, operations);
    const resources = await document.resolve(settings.Resources) || Object.create(null);
    const rawContents = await document.resolve(node.Contents);
    const contents = rawContents === undefined || rawContents === null ? [] : Array.isArray(rawContents) ? rawContents : [rawContents];
    if (contents.length > 10000) throw new Error('PDF 页面内容流数量过多。');
    for (const reference of contents) await parser.run(await document.stream(await document.resolve(reference)), resources);
    parser.finish(); operations.push({ type: 'restore' });
    if (node.Annots) document.warn('PDF 批注、链接热点、表单控件和交互动作尚未绘制或执行。');
    pages.push({ width, height, operations });
  }

  await visit(catalog.Pages);
  if (!pages.length) throw new Error('PDF 中没有可预览的页面。');
  const fonts = await Promise.all(Array.from(document.fontCache?.values() || [], async value => {
    const { id, family, fallbackFamily, weight, style, program } = await value;
    return { id, family, fallbackFamily, weight, style, ...(program ? { program } : {}) };
  }));
  document.warn('PDF 当前为自主实现的基础静态预览；复杂字体、透明度、图案和交互能力尚未完整实现。');
  return { kind: 'pdf', format: 'pdf', name, warnings, pages, fonts };
}
