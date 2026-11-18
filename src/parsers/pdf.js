// Original basic PDF renderer parser. Reference: Adobe's ISO 32000-1:2008,
// https://opensource.adobe.com/dc-acrobat-sdk-docs/standards/pdfstandards/pdf/PDF32000_2008.pdf
import { PdfDocument } from './pdf-document.js';
import { PdfContent, rgb } from './pdf-content.js';
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

  async function annotationNumbers(value, max) {
    const array = await document.resolve(value);
    if (!Array.isArray(array) || array.length > max) return null;
    const values = await Promise.all(array.map(item => document.resolve(item)));
    return values.every(Number.isFinite) ? values : null;
  }

  async function annotations(value, parser, matrix) {
    const references = await document.resolve(value).catch(() => null);
    if (!Array.isArray(references) || references.length > 10000) {
      document.warn('PDF 页面批注列表无效或过大，未显示。');
      return;
    }
    let started = false;
    for (const reference of references) {
      const annotation = await document.resolve(reference).catch(() => null);
      if (!isDictionary(annotation)) {
        document.warn('部分 PDF 批注字典无效，未显示。');
        continue;
      }
      const flags = await document.resolve(annotation.F).catch(() => 0);
      if (Number.isInteger(flags) && flags & (2 | 32)) continue; // Hidden or NoView.
      const subtype = asName(annotation.Subtype);
      const highlight = subtype === 'Highlight';
      if (!highlight && subtype !== 'Underline') {
        document.warn('部分 PDF 批注类型、链接热点、表单控件或交互动作尚未显示。');
        continue;
      }
      let points = await annotationNumbers(annotation.QuadPoints, 8000).catch(() => null);
      if (!points?.length || points.length % 8) {
        const rect = await annotationNumbers(annotation.Rect, 4).catch(() => null);
        if (rect?.length !== 4 || rect[0] === rect[2] || rect[1] === rect[3]) {
          document.warn('部分 PDF 文本标记坐标无效，未显示。');
          continue;
        }
        const [x, right] = [Math.min(rect[0], rect[2]), Math.max(rect[0], rect[2])];
        const [y, top] = [Math.min(rect[1], rect[3]), Math.max(rect[1], rect[3])];
        points = [x, y, right, y, right, top, x, top];
      }
      const components = await annotationNumbers(annotation.C, 4).catch(() => null);
      const opacity = await document.resolve(annotation.CA).catch(() => undefined);
      const alpha = Number.isFinite(opacity) ? Math.max(0, Math.min(1, opacity)) : 1;
      if (!alpha || components?.length === 0) continue;
      let color = highlight ? '#ffff00' : '#000000';
      if (components?.length === 1) color = rgb(Array(3).fill(components[0]));
      if (components?.length === 3) color = rgb(components);
      if (components?.length === 4) color = rgb(components.slice(0, 3).map(value => 1 - Math.min(1, value + components[3])));
      if (annotation.AP) document.warn('部分 PDF 文本标记按坐标近似绘制，自定义外观尚未应用。');
      for (let i = 0; i < points.length; i += 8) {
        let [x0, y0, x1, y1, x2, y2, x3, y3] = points.slice(i, i + 8);
        const cross = (x1 - x0) * (y2 - y0) - (y1 - y0) * (x2 - x0);
        if (!Number.isFinite(cross) || cross === 0) continue;
        // Acrobat also writes top-left, top-right, bottom-left, bottom-right.
        if (cross < 0) [x0, y0, x1, y1, x2, y2, x3, y3] = [x2, y2, x3, y3, x1, y1, x0, y0];
        if (!started) {
          parser.add({ type: 'save' });
          parser.add({ type: 'transform', matrix });
          started = true;
        }
        const commands = [['M', x0, y0], ['L', x1, y1]];
        if (highlight) {
          commands.push(['L', x2, y2], ['L', x3, y3], ['Z']);
          parser.add({ type: 'path', commands, fill: color, fillOpacity: alpha, blend: 'multiply' });
        } else {
          const lineWidth = Math.min(3, Math.max(0.5, Math.abs(cross) / Math.hypot(x1 - x0, y1 - y0) * 0.08));
          parser.add({ type: 'path', commands, stroke: color, strokeOpacity: alpha, lineWidth });
        }
      }
    }
    if (started) parser.add({ type: 'restore' });
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
    if (node.Annots) await annotations(node.Annots, parser, matrices[rotation]);
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
