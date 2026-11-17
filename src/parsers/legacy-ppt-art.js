import { binary, utf16, warn } from './legacy-binary.js';
import { createPptTextReader } from './legacy-ppt-text.js';
import { readLegacyFillPaint, legacyCornerRadius } from './legacy-ppt-paint.js';
import { readZip } from '../containers/zip.js';
import { parseXml } from '../xml.js';
import { readDrawingColor, cssColor, readLinearFill } from './ooxml-paint.js';

// Object identity keeps compatibility ZIPs private; they never enter the model
// sent from the worker to the host. The async entry consumes and deletes them.
const roundtripData = new WeakMap();

const UNIT = 6; // MS-PPT master units are 1/576 inch; CSS pixels are 1/96 inch.
const hex = value => `#${[value & 255, value >>> 8 & 255, value >>> 16 & 255].map(component => component.toString(16).padStart(2, '0')).join('')}`;
const signed = value => value | 0;

export function createPptArtReader(context) {
  const { source, header, children, document, compound, persists, warnings } = context;
  const documentChildren = children(document);
  const atom = documentChildren.find(record => record.type === 0x03e9);
  let width = 960, height = 540;
  if (atom?.length >= 8) {
    width = signed(source.u32(atom.data)) / UNIT; height = signed(source.u32(atom.data + 4)) / UNIT;
    if (!(width > 0 && height > 0 && width <= 20000 && height <= 20000)) throw new Error('PPT 页面尺寸无效或超过当前预览范围。');
  }
  const walk = (record, visit, depth = 0) => {
    if (depth > 64) throw new Error('PPT 绘图嵌套层数超过当前预览上限。');
    visit(record);
    if (record.version === 15) for (const child of children(record)) walk(child, visit, depth + 1);
  };
  const fonts = [];
  walk(document, record => {
    if (record.type === 0x0fb7 && record.length >= 64) fonts[record.instance] = utf16.decode(source.slice(record.data, 64)).replace(/\0.*$/s, '');
  });
  function colors(container) {
    const record = children(container).find(record => record.type === 0x07f0 && record.instance === 1) || children(container).find(record => record.type === 0x07f0);
    if (!record || record.length < 32) return null;
    return Array.from({ length: 8 }, (_, index) => hex(source.u32(record.data + index * 4)));
  }
  function color(value, scheme, text = false) {
    if (text) {
      const index = value >>> 24;
      if (index < 8) return scheme?.[index] || '#000000';
      if (index === 255) return undefined;
      return hex(value);
    }
    if (value & 0x08000000) return scheme?.[value & 255] || '#000000';
    if (value & 0x10000000) { warn(warnings, '部分 PPT 系统颜色使用当前配色近似显示。'); return scheme?.[1] || '#000000'; }
    return hex(value);
  }
  const textReader = createPptTextReader({ source, children, warnings, color, fonts });
  function properties(records) {
    const result = new Map();
    for (const record of records.filter(record => [0xf00b, 0xf121, 0xf122].includes(record.type))) {
      if (record.instance * 6 > record.length) throw new Error('PPT 绘图属性表长度无效。');
      let complexOffset = record.data + record.instance * 6;
      for (let index = 0; index < record.instance; index++) {
        const offset = record.data + index * 6;
        const flags = source.u16(offset), id = flags & 0x3fff, value = source.u32(offset + 2);
        if (flags & 0x8000) {
          if (complexOffset + value > record.end) throw new Error('PPT 绘图复杂属性长度无效。');
          result.set(id, { value, bytes: source.slice(complexOffset, value), complex: true });
          complexOffset += value;
        } else result.set(id, { value, complex: false });
      }
    }
    return result;
  }
  const value = (properties, id, fallback) => properties.get(id)?.value ?? fallback;
  const flag = (properties, id, bit, fallback) => {
    const flags = value(properties, id, 0);
    return flags & (bit << 16) ? !!(flags & bit) : fallback;
  };
  const defaultScheme = ['#ffffff', '#000000', '#808080', '#000000', '#ffffff', '#4472c4', '#ed7d31', '#70ad47'];
  const stores = [];
  walk(document, record => {
    if (record.type === 0xf001) stores.push(...children(record));
  });
  const imageCache = new Map();
  let imageBytes = 0;
  function blip(bytes, offset, boundary = bytes.length) {
    const input = binary(bytes);
    input.need(offset, 8);
    const flags = input.u16(offset), type = input.u16(offset + 2), size = input.u32(offset + 4);
    const start = offset + 8, end = start + size, instance = flags >>> 4;
    if (end > boundary || end > bytes.length) throw new Error('PPT 图片 BLIP 记录长度无效。');
    let mime, uids;
    if (type === 0xf01e && [0x6e0, 0x6e1].includes(instance)) { mime = 'image/png'; uids = instance === 0x6e1 ? 2 : 1; }
    else if ([0xf01d, 0xf02a].includes(type) && [0x46a, 0x46b, 0x6e2, 0x6e3].includes(instance)) { mime = 'image/jpeg'; uids = instance & 1 ? 2 : 1; }
    else if (type === 0xf01f && [0x7a8, 0x7a9].includes(instance)) { mime = 'image/bmp'; uids = instance === 0x7a9 ? 2 : 1; }
    else { warn(warnings, 'PPT 中的 EMF / WMF / PICT / TIFF 或其他图片编码尚未显示。'); return null; }
    const dataOffset = start + uids * 16 + 1;
    if (dataOffset >= end) throw new Error('PPT 图片没有有效的图像数据。');
    let data = bytes.subarray(dataOffset, end);
    if (mime === 'image/png') {
      if (data.length < 24 || ![137, 80, 78, 71, 13, 10, 26, 10].every((byte, i) => data[i] === byte)) throw new Error('PPT PNG 图片标识无效。');
      const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
      if (!dv.getUint32(16) || !dv.getUint32(20) || dv.getUint32(16) * dv.getUint32(20) > 16000000) throw new Error('PPT 图片像素数量超过当前上限。');
    } else if (mime === 'image/jpeg') {
      if (data[0] !== 255 || data[1] !== 216) throw new Error('PPT JPEG 图片标识无效。');
      let cursor = 2, found = false;
      while (cursor + 3 < data.length) {
        if (data[cursor++] !== 255) throw new Error('PPT JPEG 标记结构无效。');
        while (cursor < data.length && data[cursor] === 255) cursor++;
        const marker = data[cursor++];
        if (marker === 0xd9 || marker === 0xda) break;
        if (marker === 0xd8 || marker >= 0xd0 && marker <= 0xd7 || marker === 1) continue;
        if (cursor + 2 > data.length) break;
        const length = data[cursor] * 256 + data[cursor + 1];
        if (length < 2 || cursor + length > data.length) throw new Error('PPT JPEG 段长度无效。');
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4, 0xc8, 0xcc].includes(marker)) {
          if (length < 8) throw new Error('PPT JPEG 尺寸记录不完整。');
          const height = data[cursor + 3] * 256 + data[cursor + 4], width = data[cursor + 5] * 256 + data[cursor + 6];
          if (!width || !height || width * height > 16000000) throw new Error('PPT JPEG 图片像素数量超过当前上限。');
          found = true; break;
        }
        cursor += length;
      }
      if (!found) throw new Error('PPT JPEG 图片缺少可读取的尺寸记录。');
    } else {
      const dib = binary(data), headerSize = dib.u32(0);
      if (headerSize < 40 || headerSize > data.length) { warn(warnings, 'PPT 中的旧式 DIB 图片头尚未支持。'); return null; }
      const w = Math.abs(signed(dib.u32(4))), h = Math.abs(signed(dib.u32(8))), bits = dib.u16(14), compression = dib.u32(16);
      if (!w || !h || w * h > 16000000) throw new Error('PPT DIB 图片像素数量超过当前上限。');
      if (![0, 3].includes(compression)) { warn(warnings, 'PPT 中经过 RLE 等压缩的 DIB 图片尚未支持。'); return null; }
      const palette = dib.u32(32) || (bits <= 8 ? 2 ** bits : 0);
      const pixelOffset = 14 + headerSize + palette * 4 + (compression === 3 && headerSize === 40 ? 12 : 0);
      if (pixelOffset > data.length + 14) throw new Error('PPT DIB 调色板长度无效。');
      const bmp = new Uint8Array(data.length + 14), dv = new DataView(bmp.buffer);
      bmp[0] = 66; bmp[1] = 77; dv.setUint32(2, bmp.length, true); dv.setUint32(10, pixelOffset, true); bmp.set(data, 14); data = bmp;
    }
    imageBytes += data.length;
    if (data.length > 32 * 1024 * 1024 || imageBytes > 128 * 1024 * 1024) throw new Error('PPT 图片数据超过当前容量限制。');
    return { bytes: data, mime };
  }
  function image(index) {
    if (imageCache.has(index)) return imageCache.get(index);
    const record = stores[index - 1];
    if (!record) { warn(warnings, '部分 PPT 图片引用在当前文档图片目录中不存在。'); return null; }
    let result;
    if (record.type === 0xf007) {
      if (record.length < 36) throw new Error('PPT 图片 BSE 目录记录不完整。');
      const references = source.u32(record.data + 24), delay = source.u32(record.data + 28), nameLength = source.u8(record.data + 33);
      if (36 + nameLength > record.length) throw new Error('PPT 图片目录名称超出所属记录。');
      if (!references) { warn(warnings, '部分 PPT 图片引用了空的图片目录槽。'); result = null; }
      else if (record.length > 36 + nameLength) result = blip(source.bytes, record.data + 36 + nameLength, record.end);
      else {
        const pictures = compound.get('Pictures');
        if (pictures && delay !== 0xffffffff) result = blip(pictures, delay);
        else { warn(warnings, '部分 PPT 图片缺少 Pictures 流中的有效数据。'); result = null; }
      }
    } else result = blip(source.bytes, record.offset, record.end);
    imageCache.set(index, result); return result;
  }
  function rect(record) {
    if (!record) return null;
    const count = record.length === 8 ? 2 : record.length === 16 ? 4 : 0;
    if (!count) throw new Error('PPT 形状锚点长度无效。');
    const values = Array.from({ length: 4 }, (_, index) => count === 2 ? (source.u16(record.data + index * 2) << 16 >> 16) : signed(source.u32(record.data + index * 4)));
    return record.type === 0xf010 ? { left: values[1], top: values[0], right: values[2], bottom: values[3] } : { left: values[0], top: values[1], right: values[2], bottom: values[3] };
  }
  const masters = new Map();
  for (const list of documentChildren.filter(record => record.type === 0x0ff0 && record.instance === 1)) {
    for (const entry of children(list).filter(record => record.type === 0x03f3)) {
      if (entry.length >= 16 && persists.has(source.u32(entry.data))) masters.set(source.u32(entry.data + 12), header(persists.get(source.u32(entry.data))));
    }
  }
  const rootMatrix = { sx: 1 / UNIT, sy: 1 / UNIT, tx: 0, ty: 0 };
  const mapRect = (bounds, matrix) => ({ x: bounds.left * matrix.sx + matrix.tx, y: bounds.top * matrix.sy + matrix.ty, width: (bounds.right - bounds.left) * matrix.sx, height: (bounds.bottom - bounds.top) * matrix.sy });

  function renderSlide(container, entry, depth = 0, inheritedScheme) {
    if (depth > 16) throw new Error('PPT 母版引用嵌套层数过多。');
    const records = children(container), slideAtom = records.find(record => record.type === 0x03ef);
    const masterId = slideAtom?.length >= 24 ? source.u32(slideAtom.data + 12) : 0;
    const flags = slideAtom?.length >= 24 ? source.u16(slideAtom.data + 20) : 0;
    const master = masters.get(masterId);
    const scheme = inheritedScheme || ((flags & 2) && master ? colors(master) : colors(container)) || (master ? colors(master) : null) || defaultScheme;
    let background = scheme[0], elements = [], masterProperties = new Map();
    const defaults = new Map();
    if (master) {
      const inherited = renderSlide(master, { blocks: [] }, depth + 1, scheme);
      if (flags & 4) background = inherited.background;
      if (flags & 1) elements.push(...inherited.elements.filter(element => !element.placeholder));
      masterProperties = inherited.shapeProperties;
      for (const [type, levels] of inherited.defaults) defaults.set(type, levels);
    }
    for (const [type, levels] of textReader.masterStyles(container, scheme)) defaults.set(type, levels);
    const shapeProperties = new Map(masterProperties);
    const outline = entry.blocks || [];
    let positioned = 0;
    const renderShape = (record, matrix, group = false) => {
      const records = children(record), sp = records.find(record => record.type === 0xf00a);
      if (!sp || sp.length < 8) return;
      const shapeFlags = source.u32(sp.data + 4);
      if (shapeFlags & 8) return;
      const ownProperties = properties(records), masterShape = masterProperties.get(value(ownProperties, 0x0301, -1));
      const props = new Map([...(masterShape?.properties || []), ...ownProperties]);
      const anchorRecord = records.find(record => record.type === 0xf00f) || records.find(record => record.type === 0xf010);
      const bounds = rect(anchorRecord) || masterShape?.bounds;
      shapeProperties.set(source.u32(sp.data), { properties: props, bounds });
      const placeholder = records.some(record => record.type === 0xf011 && children(record).some(child => child.type === 0x0bc3));
      const isBackground = !!(shapeFlags & 0x400), shapeType = sp.instance;
      if (group || shapeFlags & 1) return;
      const fillColor = color(value(props, 0x0181, 0x08000004), scheme);
      const fill = flag(props, 0x01bf, 0x10, shapeType !== 202) ? fillColor : 'transparent';
      const fillPaint = fill !== 'transparent' ? readLegacyFillPaint(props, value => color(value, scheme), warnings) : undefined;
      if (isBackground && (!(flags & 4) || !master)) { background = fill === 'transparent' ? background : fill; }
      if (!bounds && !isBackground) { if (records.some(record => record.type === 0xf00d)) warn(warnings, '部分 PPT 文本框没有可读取的定位信息。'); return; }
      const location = bounds ? mapRect(bounds, anchorRecord?.type === 0xf010 ? rootMatrix : matrix) : { x: 0, y: 0, width, height };
      const common = { ...location, ...(placeholder ? { placeholder: true } : {}) };
      if (common.width < 0) { common.x += common.width; common.width *= -1; common.flipH = true; }
      if (common.height < 0) { common.y += common.height; common.height *= -1; common.flipV = true; }
      if (shapeFlags & 0x40) common.flipH = !common.flipH;
      if (shapeFlags & 0x80) common.flipV = !common.flipV;
      const rotation = signed(value(props, 0x0004, 0)) / 65536;
      if (rotation) common.rotation = rotation;
      const stroke = flag(props, 0x01ff, 0x08, shapeType !== 202) ? color(value(props, 0x01c0, 0), scheme) : undefined;
      const imageProperty = props.get(0x0104) || props.get(0x0186);
      const bitmap = imageProperty?.complex ? blip(imageProperty.bytes, 0) : imageProperty?.value ? image(imageProperty.value) : null;
      if (value(props, 0x0180, 0) && ![2, 3, 4, 7].includes(value(props, 0x0180, 0))) warn(warnings, 'PPT 路径渐变、图案等填充当前使用基础颜色近似显示。');
      if (!isBackground && (!bitmap || ownProperties.has(0x0181) || ownProperties.has(0x01c0)) && (fill !== 'transparent' || stroke)) {
        const shape = { 1: 'rect', 2: 'roundRect', 3: 'ellipse', 20: 'line', 32: 'line', 202: 'rect' }[shapeType] || 'rect';
        if (![0, 1, 2, 3, 20, 32, 75, 202].includes(shapeType)) warn(warnings, '部分 PPT 自定义或预设复杂形状暂以外接矩形近似显示。');
        const element = { type: 'shape', ...common, shape, fill, stroke, strokeWidth: Math.max(0, value(props, 0x01cb, 9525) / 9525), ...(fillPaint ? { fillPaint } : {}) };
        if (shapeType === 2) element.cornerRadius = legacyCornerRadius(props, common.width, common.height);
        const fillOpacity = value(props, 0x0182, 65536) / 65536, strokeOpacity = value(props, 0x01c1, 65536) / 65536;
        if (fillOpacity !== 1 && !fillPaint) element.fillOpacity = Math.max(0, Math.min(1, fillOpacity));
        if (strokeOpacity !== 1) element.strokeOpacity = Math.max(0, Math.min(1, strokeOpacity));
        elements.push(element);
        const metro = props.get(0x03a9);
        if (metro?.complex) roundtripData.set(element, { bytes: metro.bytes, scheme });
      }
      if (bitmap) {
        const crop = { top: signed(value(props, 0x0100, 0)) / 65536, bottom: signed(value(props, 0x0101, 0)) / 65536, left: signed(value(props, 0x0102, 0)) / 65536, right: signed(value(props, 0x0103, 0)) / 65536 };
        elements.push({ type: 'image', ...common, ...bitmap, ...(Object.values(crop).some(Boolean) ? { crop } : {}) });
      }
      for (const textbox of records.filter(record => record.type === 0xf00d)) {
        const text = textReader.readBox(textbox, outline, scheme, defaults);
        if (!text?.runs.some(run => run.text)) continue;
        const left = value(props, 0x0081, 91440) / 9525, top = value(props, 0x0082, 45720) / 9525;
        const right = value(props, 0x0083, 91440) / 9525, bottom = value(props, 0x0084, 45720) / 9525;
        const anchor = value(props, 0x0087, 0), verticalAlign = ['top', 'middle', 'bottom'][anchor % 3] || 'top';
        const { flipH, flipV, ...textPosition } = common;
        const element = { type: 'text', ...textPosition, x: common.x + left, y: common.y + top, width: Math.max(0, common.width - left - right), height: Math.max(0, common.height - top - bottom), fontSize: 24, color: scheme[1], fontFamily: fonts[0], verticalAlign, ...text };
        elements.push(element);
        const metro = props.get(0x03a9);
        if (metro?.complex) roundtripData.set(element, { bytes: metro.bytes, scheme, outer: common });
        positioned++;
      }
    };
    const drawing = (record, matrix = rootMatrix, nesting = 0) => {
      if (nesting > 64) throw new Error('PPT 绘图组嵌套过多。');
      if (record.type === 0xf004) { renderShape(record, matrix); return; }
      if (record.type === 0xf003) {
        const records = children(record), head = records.find(record => record.type === 0xf004);
        let childMatrix = matrix;
        if (head) {
          const headRecords = children(head), groupBounds = headRecords.find(record => record.type === 0xf009);
          const anchor = headRecords.find(record => record.type === 0xf00f) || headRecords.find(record => record.type === 0xf010);
          if (groupBounds && anchor) {
            const from = rect(groupBounds), to = mapRect(rect(anchor), anchor.type === 0xf010 ? rootMatrix : matrix);
            if (from.right === from.left || from.bottom === from.top) throw new Error('PPT 绘图组坐标范围为空。');
            childMatrix = { sx: to.width / (from.right - from.left), sy: to.height / (from.bottom - from.top), tx: 0, ty: 0 };
            childMatrix.tx = to.x - from.left * childMatrix.sx; childMatrix.ty = to.y - from.top * childMatrix.sy;
            if (value(properties(headRecords), 4, 0)) warn(warnings, 'PPT 旋转组合图形中的子对象尚未完整还原。');
          }
          renderShape(head, matrix, true);
        }
        for (const child of records) if (child !== head) drawing(child, childMatrix, nesting + 1);
      } else if (record.version === 15) for (const child of children(record)) drawing(child, matrix, nesting + 1);
    };
    for (const record of records.filter(record => record.type === 0x040c)) drawing(record);
    if (!positioned && depth === 0) {
      const fallback = outline.filter(block => block.text && !block.used);
      if (!fallback.length) {
        const raw = [];
        const collect = record => {
          for (const child of children(record)) {
            const text = textReader.textFromRecord(child);
            if (text !== null) raw.push({ text, textType: 4, records: [] });
            else if (child.version === 15 || child.type === 0xf00d) collect(child);
          }
        };
        collect(container); fallback.push(...raw);
      }
      if (fallback.length) {
        warn(warnings, '部分 PPT 页面没有可定位的文本框，正文按阅读模式显示。');
        elements.push({ type: 'text', x: 48, y: 44, width: Math.max(1, width - 96), height: Math.max(1, height - 88), fontSize: 24, runs: [{ text: fallback.map(block => textReader.clean(block.text)).join('\n\n') }] });
      }
    }
    return { background, elements, shapeProperties, defaults };
  }
  return { width, height, renderSlide, textReader };
}

const localChildren = (node, name) => [...(node?.children || [])].filter(child => !name || child.localName === name);
const child = (node, name) => localChildren(node, name)[0] || null;
const attribute = (node, name) => [...(node?.attributes || [])].find(item => item.localName === name)?.value ?? '';
const number = (node, name, fallback = 0) => {
  const raw = attribute(node, name), value = Number(raw);
  return raw !== '' && Number.isFinite(value) ? value : fallback;
};
const plainText = runs => (runs || []).map(run => run.text || '').join('').replace(/\r\n?/g, '\n');

function metroParagraphs(body, element, theme, warnings) {
  let basePosition = 0;
  const baseRuns = (element.runs || []).map(run => { const start = basePosition; basePosition += (run.text || '').length; return { start, end: basePosition, run }; });
  const paragraphs = [];
  let position = 0;
  for (const [index, paragraph] of localChildren(body, 'p').entries()) {
    const properties = child(paragraph, 'pPr'), defaults = child(properties, 'defRPr');
    const runs = [];
    for (const item of localChildren(paragraph)) {
      if (item.localName === 'br') { runs.push({ text: '\n' }); position++; continue; }
      if (!['r', 'fld'].includes(item.localName)) continue;
      const text = child(item, 't')?.textContent || '', own = child(item, 'rPr');
      const base = baseRuns.find(run => run.start <= position && run.end > position)?.run || element.paragraphs?.[index]?.runs?.[0] || {};
      const run = { ...base, text };
      const get = name => attribute(own, name) || attribute(defaults, name);
      for (const [key, target] of [['b', 'bold'], ['i', 'italic']]) {
        const value = get(key); if (value !== '') run[target] = value === '1' || value === 'true';
      }
      const size = Number(get('sz'));
      if (get('sz') !== '' && size > 0 && size <= 400000) run.fontSize = size / 75;
      const spacing = get('spc');
      if (spacing !== '' && Number.isFinite(Number(spacing))) run.letterSpacing = Number(spacing) / 75;
      const underline = get('u'); if (underline) run.underline = underline !== 'none';
      const strike = get('strike'); if (strike) run.strike = strike !== 'noStrike';
      const font = attribute(child(own, 'ea'), 'typeface') || attribute(child(own, 'latin'), 'typeface') || attribute(child(defaults, 'ea'), 'typeface') || attribute(child(defaults, 'latin'), 'typeface');
      if (font && !font.startsWith('+')) run.fontFamily = font;
      const color = readDrawingColor(child(own, 'solidFill') || child(defaults, 'solidFill'), theme, warnings);
      if (color) run.color = cssColor(color);
      runs.push(run); position += text.length;
    }
    const result = { ...(element.paragraphs?.[index] || {}), runs };
    const align = attribute(properties, 'algn');
    if (align) result.align = { l: 'left', ctr: 'center', r: 'right', just: 'justify' }[align] || 'left';
    const line = child(properties, 'lnSpc'), percent = child(line, 'spcPct'), points = child(line, 'spcPts');
    if (percent) { result.lineHeight = number(percent, 'val') / 100000; delete result.lineHeightPx; }
    if (points) { result.lineHeightPx = number(points, 'val') / 75; delete result.lineHeight; }
    const linePixels = result.lineHeightPx || Math.max(1, ...runs.map(run => run.fontSize || element.fontSize || 24)) * (result.lineHeight || 1.2);
    for (const [tag, target] of [['spcBef', 'spaceBefore'], ['spcAft', 'spaceAfter']]) {
      const spacing = child(properties, tag), pts = child(spacing, 'spcPts'), pct = child(spacing, 'spcPct');
      if (pts) result[target] = number(pts, 'val') / 75;
      else if (pct) result[target] = number(pct, 'val') / 100000 * linePixels;
    }
    paragraphs.push(result); position++;
  }
  return paragraphs;
}

// MS-ODRAW Appendix F stores newer DrawingML properties in the shape's
// metroBlob ZIP. Read that named compatibility part, never nearby PPTX files.
export async function enrichPptRoundtrip(model) {
  const cache = new Map();
  let expanded = 0;
  for (const slide of model.slides || []) for (const element of slide.elements || []) {
    const metadata = roundtripData.get(element);
    if (!metadata) continue;
    roundtripData.delete(element);
    let document;
    if (cache.has(metadata.bytes)) document = cache.get(metadata.bytes);
    else {
      try {
        const files = await readZip(metadata.bytes);
        for (const bytes of files.values()) expanded += bytes.length;
        if (expanded > 128 * 1024 * 1024) throw new Error('PPT 兼容属性解压后超过 128MB 容量限制。');
        const bytes = files.get('drs/shapexml.xml');
        document = bytes ? parseXml(bytes) : null;
      } catch (error) {
        if (/容量限制|超过/.test(error.message)) throw error;
        warn(model.warnings, '部分 PPT 新版兼容属性无法读取，保留基础二进制样式。');
        document = null;
      }
      cache.set(metadata.bytes, document);
    }
    if (!document || document.documentElement?.localName !== 'sp') continue;
    const root = document.documentElement;
    const scheme = metadata.scheme || [];
    const theme = { lt1: scheme[0], dk1: scheme[1], lt2: scheme[4], dk2: scheme[3], accent1: scheme[5], accent2: scheme[6], accent3: scheme[7] };
    if (element.type === 'shape') {
      const properties = child(root, 'spPr');
      const paint = readLinearFill(properties, theme, model.warnings);
      if (paint) element.fillPaint = paint;
      const solid = readDrawingColor(child(properties, 'solidFill'), theme, model.warnings);
      if (solid) element.fill = cssColor(solid);
      const geometry = child(properties, 'prstGeom');
      if (attribute(geometry, 'prst') === 'roundRect') {
        const adjust = localChildren(child(geometry, 'avLst'), 'gd').find(node => attribute(node, 'name') === 'adj');
        const match = /^val\s+(\d+(?:\.\d+)?)$/.exec(attribute(adjust, 'fmla'));
        if (match) element.cornerRadius = Math.min(element.width, element.height) * Math.max(0, Math.min(50000, Number(match[1]))) / 100000;
      }
      continue;
    }
    if (element.type !== 'text') continue;
    const body = child(root, 'txBody');
    if (!body) continue;
    const paragraphs = metroParagraphs(body, element, theme, model.warnings);
    const runs = paragraphs.flatMap((paragraph, index) => index ? [{ text: '\n' }, ...paragraph.runs] : paragraph.runs);
    if (plainText(runs) !== plainText(element.runs)) {
      warn(model.warnings, '部分 PPT 兼容属性与当前正文不同，继续使用当前二进制正文。');
      continue;
    }
    element.paragraphs = paragraphs; element.runs = runs;
    if (paragraphs[0]?.align) element.align = paragraphs[0].align;
    const properties = child(body, 'bodyPr');
    const anchor = attribute(properties, 'anchor');
    if (anchor) element.verticalAlign = { t: 'top', ctr: 'middle', b: 'bottom' }[anchor] || 'top';
    if (metadata.outer && properties) {
      const left = number(properties, 'lIns', 91440) / 9525, right = number(properties, 'rIns', 91440) / 9525;
      const top = number(properties, 'tIns', 45720) / 9525, bottom = number(properties, 'bIns', 45720) / 9525;
      element.x = metadata.outer.x + left; element.y = metadata.outer.y + top;
      element.width = Math.max(0, metadata.outer.width - left - right); element.height = Math.max(0, metadata.outer.height - top - bottom);
    }
  }
  return model;
}
