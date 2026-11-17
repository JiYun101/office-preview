import { EMU_PER_PIXEL, elements, children, first, attr, relationshipId, numberAttr, warn, xmlPart, relationships, relatedPath, imageFromRelationship, themeColors } from './ooxml-common.js';
import { readDrawingColor, cssColor, readLinearFill } from './ooxml-paint.js';

export function parsePresentation(parts, path, warnings) {
  const presentation = xmlPart(parts, path);
  const rels = relationships(parts, path, warnings);
  const size = first(presentation.documentElement, 'sldSz');
  const width = numberAttr(size, 'cx', 9144000) / EMU_PER_PIXEL;
  const height = numberAttr(size, 'cy', 6858000) / EMU_PER_PIXEL;
  if (width <= 0 || height <= 0) throw new Error('PPTX 幻灯片尺寸无效。');
  const baseTheme = themeColors(parts, rels);
  warn(warnings, 'PPTX 当前还原基本文本、图片、几何形状；母版装饰、完整样式继承和精确文字换行尚未实现。');

  function placeholderShape(document, shape) {
    const placeholder = elements(shape, 'ph')[0];
    if (!placeholder || !document) return null;
    const index = attr(placeholder, 'idx') || '0';
    const type = attr(placeholder, 'type') || 'obj';
    const candidates = elements(document, 'sp');
    return candidates.find(candidate => {
      const value = elements(candidate, 'ph')[0];
      return value && (attr(value, 'idx') || '0') === index;
    }) || candidates.find(candidate => attr(elements(candidate, 'ph')[0], 'type') === type) || null;
  }

  const slides = children(first(presentation.documentElement, 'sldIdLst'), 'sldId').map((entry, slideIndex) => {
    const slidePath = rels.get(relationshipId(entry))?.target;
    if (!slidePath) throw new Error('PPTX 幻灯片关系缺失。');
    const doc = xmlPart(parts, slidePath);
    if (doc.documentElement.localName !== 'sld') throw new Error('PPTX 幻灯片 XML 类型无效。');
    const slideRels = relationships(parts, slidePath, warnings);
    const layoutPath = relatedPath(slideRels, 'slideLayout');
    const layout = layoutPath ? xmlPart(parts, layoutPath) : null;
    const layoutRels = layoutPath ? relationships(parts, layoutPath, warnings) : new Map();
    const masterPath = relatedPath(layoutRels, 'slideMaster');
    const master = masterPath ? xmlPart(parts, masterPath) : null;
    const theme = { ...baseTheme, ...(masterPath ? themeColors(parts, relationships(parts, masterPath, warnings)) : {}) };
    const result = [];

    function rectangle(transform, mapping) {
      const off = first(transform, 'off');
      const ext = first(transform, 'ext');
      return {
        x: (numberAttr(off, 'x') * mapping.sx + mapping.x) / EMU_PER_PIXEL,
        y: (numberAttr(off, 'y') * mapping.sy + mapping.y) / EMU_PER_PIXEL,
        width: numberAttr(ext, 'cx') * mapping.sx / EMU_PER_PIXEL,
        height: numberAttr(ext, 'cy') * mapping.sy / EMU_PER_PIXEL,
      };
    }

    function textParagraphs(body) {
      const result = [];
      const paragraphs = children(body, 'p');
      for (let paragraphIndex = 0; paragraphIndex < paragraphs.length; paragraphIndex += 1) {
        const paragraph = paragraphs[paragraphIndex];
        const paragraphProperties = first(paragraph, 'pPr');
        const defaults = first(paragraphProperties, 'defRPr');
        const runs = [];
        for (const item of children(paragraph)) {
          if (item.localName === 'br') { runs.push({ text: '\n' }); continue; }
          if (!['r', 'fld'].includes(item.localName)) continue;
          const properties = first(item, 'rPr');
          const style = {};
          const size = numberAttr(properties, 'sz', numberAttr(defaults, 'sz'));
          if (size > 0) style.fontSize = size / 75;
          const spacing = attr(properties, 'spc') !== '' ? attr(properties, 'spc') : attr(defaults, 'spc');
          if (spacing !== '' && Number.isFinite(Number(spacing))) style.letterSpacing = Number(spacing) / 75;
          const baseline = attr(properties, 'baseline') !== '' ? attr(properties, 'baseline') : attr(defaults, 'baseline');
          if (baseline !== '' && Number.isFinite(Number(baseline)) && style.fontSize) style.baselineShift = Number(baseline) / 100000 * style.fontSize;
          for (const [xml, key] of [['b', 'bold'], ['i', 'italic']]) {
            const value = attr(properties, xml) || attr(defaults, xml);
            if (value !== '') style[key] = value === '1' || value === 'true';
          }
          const color = cssColor(readDrawingColor(first(properties, 'solidFill'), theme, warnings)) || cssColor(readDrawingColor(first(defaults, 'solidFill'), theme, warnings));
          if (color) style.color = color;
          const family = attr(first(properties, 'ea'), 'typeface') || attr(first(properties, 'latin'), 'typeface') || attr(first(defaults, 'ea'), 'typeface') || attr(first(defaults, 'latin'), 'typeface');
          if (family && !family.startsWith('+')) style.fontFamily = family;
          runs.push({ text: first(item, 't')?.textContent || '', ...style });
        }
        const align = attr(paragraphProperties, 'algn');
        const lineSpacing = first(paragraphProperties, 'lnSpc');
        const percent = first(lineSpacing, 'spcPct'), points = first(lineSpacing, 'spcPts');
        const spacing = {};
        for (const [tag, field] of [['spcBef', 'spaceBefore'], ['spcAft', 'spaceAfter']]) {
          const node = first(paragraphProperties, tag), pt = first(node, 'spcPts'), pct = first(node, 'spcPct');
          if (pt) spacing[field] = numberAttr(pt, 'val') / 75;
          else if (pct) spacing[`${field}Ratio`] = numberAttr(pct, 'val') / 100000;
        }
        result.push({ runs, align: { ctr: 'center', r: 'right', just: 'justify' }[align] || 'left', ...(percent ? { lineHeight: numberAttr(percent, 'val') / 100000 } : points ? { lineHeightPx: numberAttr(points, 'val') / 75 } : {}), ...spacing });
      }
      return result;
    }

    function visit(tree, mapping = { x: 0, y: 0, sx: 1, sy: 1 }) {
      for (const shape of children(tree)) {
        if (shape.localName === 'grpSp') {
          const transform = first(first(shape, 'grpSpPr'), 'xfrm');
          const off = first(transform, 'off');
          const ext = first(transform, 'ext');
          const childOff = first(transform, 'chOff');
          const childExt = first(transform, 'chExt');
          const sx = numberAttr(childExt, 'cx') ? numberAttr(ext, 'cx') / numberAttr(childExt, 'cx') : 1;
          const sy = numberAttr(childExt, 'cy') ? numberAttr(ext, 'cy') / numberAttr(childExt, 'cy') : 1;
          if (numberAttr(transform, 'rot') || attr(transform, 'flipH') === '1' || attr(transform, 'flipV') === '1') warn(warnings, '旋转和翻转对象尚未还原。');
          visit(shape, {
            x: mapping.x + (numberAttr(off, 'x') - numberAttr(childOff, 'x') * sx) * mapping.sx,
            y: mapping.y + (numberAttr(off, 'y') - numberAttr(childOff, 'y') * sy) * mapping.sy,
            sx: mapping.sx * sx, sy: mapping.sy * sy,
          });
          continue;
        }
        if (!['sp', 'pic', 'cxnSp', 'graphicFrame'].includes(shape.localName)) continue;
        if (shape.localName === 'graphicFrame') {
          warn(warnings, '幻灯片中的表格、图表、SmartArt 和嵌入对象尚未渲染。');
          continue;
        }
        const properties = first(shape, 'spPr');
        let transform = first(properties, 'xfrm');
        if (!transform && shape.localName === 'sp') {
          const inherited = placeholderShape(layout, shape) || placeholderShape(master, shape);
          transform = first(first(inherited, 'spPr'), 'xfrm');
          if (!transform) {
            const masterShape = placeholderShape(master, inherited || shape);
            transform = first(first(masterShape, 'spPr'), 'xfrm');
          }
        }
        if (!transform) {
          warn(warnings, '部分对象没有可解析的位置，未显示。');
          continue;
        }
        if (numberAttr(transform, 'rot') || attr(transform, 'flipH') === '1' || attr(transform, 'flipV') === '1') warn(warnings, '旋转和翻转对象尚未还原。');
        const rect = rectangle(transform, mapping);
        if (shape.localName === 'pic') {
          const blip = elements(shape, 'blip')[0];
          const image = imageFromRelationship(parts, slideRels, attr(blip, 'embed'), warnings);
          if (image) result.push({ type: 'image', ...rect, ...image });
          if (elements(shape, 'srcRect').length) warn(warnings, '图片裁剪和图片特效尚未还原。');
          continue;
        }
        const fill = cssColor(readDrawingColor(first(properties, 'solidFill'), theme, warnings));
        const fillPaint = readLinearFill(properties, theme, warnings);
        const line = first(properties, 'ln');
        const stroke = cssColor(readDrawingColor(first(line, 'solidFill'), theme, warnings));
        const geometryNode = first(properties, 'prstGeom');
        const geometry = attr(geometryNode, 'prst');
        let cornerRadius;
        if (geometry === 'roundRect') {
          const guide = children(first(geometryNode, 'avLst'), 'gd').find(node => attr(node, 'name') === 'adj');
          const formula = attr(guide, 'fmla');
          const adjustment = guide ? /^val\s+(-?\d+(?:\.\d+)?)$/.exec(formula) : null;
          if (!guide || adjustment) cornerRadius = Math.min(rect.width, rect.height) * Math.max(0, Math.min(50000, adjustment ? Number(adjustment[1]) : 16667)) / 100000;
          else warn(warnings, 'PPTX 部分圆角调整公式尚未计算。');
        }
        if (fill || stroke || fillPaint) result.push({ type: 'shape', ...rect, fill, ...(fillPaint ? { fillPaint } : {}), stroke, ...(stroke && attr(line, 'w') !== '' ? { strokeWidth: numberAttr(line, 'w') / EMU_PER_PIXEL } : {}), ...(cornerRadius !== undefined ? { cornerRadius } : {}), shape: geometry || (shape.localName === 'cxnSp' ? 'line' : 'rect') });
        if (geometry && !['rect', 'roundRect', 'ellipse', 'line'].includes(geometry) || first(properties, 'custGeom')) warn(warnings, '复杂几何形状当前用外接矩形近似显示。');
        if (first(properties, 'pattFill') || first(properties, 'blipFill')) warn(warnings, '纹理和图片填充尚未还原。');
        const body = first(shape, 'txBody');
        if (body) {
          const paragraphs = textParagraphs(body);
          const runs = paragraphs.flatMap((paragraph, index) => index ? [{ text: '\n' }, ...paragraph.runs] : paragraph.runs);
          const firstParagraph = first(body, 'p');
          const align = attr(first(firstParagraph, 'pPr'), 'algn');
          const bodyProperties = first(body, 'bodyPr');
          const left = numberAttr(bodyProperties, 'lIns', 91440) / EMU_PER_PIXEL * mapping.sx;
          const right = numberAttr(bodyProperties, 'rIns', 91440) / EMU_PER_PIXEL * mapping.sx;
          const top = numberAttr(bodyProperties, 'tIns', 45720) / EMU_PER_PIXEL * mapping.sy;
          const bottom = numberAttr(bodyProperties, 'bIns', 45720) / EMU_PER_PIXEL * mapping.sy;
          result.push({ type: 'text', x: rect.x + left, y: rect.y + top, width: Math.max(1, rect.width - left - right), height: Math.max(1, rect.height - top - bottom), paragraphs, runs, verticalOverflow: attr(bodyProperties, 'vertOverflow') || 'overflow', verticalAlign: { ctr: 'middle', b: 'bottom' }[attr(bodyProperties, 'anchor')] || 'top', align: { ctr: 'center', r: 'right', just: 'justify' }[align] || 'left' });
        }
      }
    }

    const canvas = first(doc.documentElement, 'cSld');
    const tree = first(canvas, 'spTree');
    if (!canvas || !tree) throw new Error('PPTX 缺少幻灯片形状树。');
    visit(tree);
    const backgroundNode = first(canvas, 'bg') || first(first(layout?.documentElement, 'cSld'), 'bg') || first(first(master?.documentElement, 'cSld'), 'bg');
    const backgroundProperties = first(backgroundNode, 'bgPr');
    const background = cssColor(readDrawingColor(first(backgroundProperties, 'solidFill'), theme, warnings));
    const backgroundPaint = readLinearFill(backgroundProperties, theme, warnings);
    if (backgroundNode && !background && !backgroundPaint) warn(warnings, '部分幻灯片背景使用了尚未支持的主题或填充。');
    if (elements(doc, 'timing').length || elements(doc, 'transition').length) warn(warnings, '动画和切换效果尚未播放，仅显示静态幻灯片。');
    if (elements(doc, 'videoFile').length || elements(doc, 'audioFile').length || elements(doc, 'media').length) warn(warnings, '嵌入音频和视频尚未播放。');
    if (attr(doc.documentElement, 'show') === '0') warn(warnings, '隐藏幻灯片目前仍在预览列表中显示。');
    return { name: attr(canvas, 'name') || `幻灯片 ${slideIndex + 1}`, ...(background ? { background } : {}), ...(backgroundPaint ? { backgroundPaint } : {}), elements: result };
  });
  return { kind: 'presentation', width, height, slides };
}
