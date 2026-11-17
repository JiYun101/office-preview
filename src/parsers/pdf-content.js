import { PdfLexer, asName, isDictionary, isReference, isString, integer, numeric, matrixMultiply } from './pdf-syntax.js';
import { loadPdfFont } from './pdf-font.js';

const IDENTITY = () => [1, 0, 0, 1, 0, 0];
const unit = number => Math.max(0, Math.min(1, number));
function rgb(numbers) { return `#${numbers.map(value => Math.round(unit(value) * 255).toString(16).padStart(2, '0')).join('')}`; }
function rectCommands(box) { return [['M', box[0], box[1]], ['L', box[2], box[1]], ['L', box[2], box[3]], ['L', box[0], box[3]], ['Z']]; }
function jpegDimensions(bytes) {
  if (bytes[0] !== 0xff || bytes[1] !== 0xd8) throw new Error('PDF JPEG 图片头无效。');
  let offset = 2;
  while (offset < bytes.length) {
    if (bytes[offset++] !== 0xff) throw new Error('PDF JPEG 标记无效。');
    while (bytes[offset] === 0xff) offset += 1;
    const marker = bytes[offset++];
    if (marker === 0xda || marker === 0xd9 || marker === undefined) break;
    if (marker === 0x01 || marker >= 0xd0 && marker <= 0xd7) continue;
    const length = bytes[offset] * 256 + bytes[offset + 1];
    if (!Number.isInteger(length) || length < 2 || offset + length > bytes.length) throw new Error('PDF JPEG 标记长度无效。');
    if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
      if (length < 8 || bytes[offset + 2] !== 8) throw new Error('当前 PDF JPEG 图片仅支持 8 位样本。');
      return { width: bytes[offset + 5] * 256 + bytes[offset + 6], height: bytes[offset + 3] * 256 + bytes[offset + 4] };
    }
    offset += length;
  }
  throw new Error('PDF JPEG 缺少有效的图像尺寸头。');
}

export class PdfContent {
  constructor(document, operations) {
    this.document = document;
    this.operations = operations;
    this.state = { fill: '#000000', stroke: '#000000', fillOpacity: 1, strokeOpacity: 1, lineWidth: 1, lineCap: 0, lineJoin: 0, miterLimit: 10, dash: [], dashPhase: 0, font: null, fontSize: 12, characterSpace: 0, wordSpace: 0, horizontalScale: 1, leading: 0, rise: 0, textMode: 0, fillSpace: 'DeviceGray', strokeSpace: 'DeviceGray' };
    this.saved = [];
    this.path = [];
    this.pathPoint = null;
    this.pathStart = null;
    this.clipRule = null;
    this.textActive = false;
    this.textMatrix = IDENTITY();
    this.lineMatrix = IDENTITY();
    this.forms = new Set();
    if (!document.fontCache) document.fontCache = new Map();
  }

  add(operation) {
    this.document.operationCount = (this.document.operationCount || 0) + 1;
    if (this.document.operationCount > 250000) throw new Error('PDF 绘制操作超过当前 250000 项上限。');
    this.operations.push(operation);
  }

  command(command) {
    this.document.pathCount = (this.document.pathCount || 0) + 1;
    if (this.document.pathCount > 1000000) throw new Error('PDF 路径指令超过当前上限。');
    this.path.push(command);
  }

  save() {
    if (this.saved.length >= 128) throw new Error('PDF 图形状态栈过深。');
    this.saved.push({ ...this.state }); this.add({ type: 'save' });
  }

  restore() {
    if (!this.saved.length) throw new Error('PDF 图形状态恢复没有匹配的保存。');
    this.state = this.saved.pop(); this.add({ type: 'restore' });
  }

  strokeStyle() {
    const { lineWidth, lineCap, lineJoin, miterLimit, dash, dashPhase, strokeOpacity } = this.state;
    return { lineWidth, lineCap, lineJoin, miterLimit, dash, dashPhase, strokeOpacity };
  }

  setDash(pattern, phase) {
    if (!Array.isArray(pattern)) throw new Error('PDF 虚线数组无效。');
    const dash = pattern.map(value => numeric(value, '虚线长度'));
    if (dash.some(value => value < 0) || dash.length && !dash.some(value => value > 0)) throw new Error('PDF 虚线长度必须非负且不能全部为零。');
    this.state.dash = dash;
    this.state.dashPhase = numeric(phase, '虚线相位');
  }

  paint(operator) {
    if (['s', 'b', 'b*'].includes(operator)) this.command(['Z']);
    const fill = ['f', 'F', 'f*', 'B', 'B*', 'b', 'b*'].includes(operator);
    const stroke = ['S', 's', 'B', 'B*', 'b', 'b*'].includes(operator);
    if (this.path.length && (fill || stroke)) this.add({ type: 'path', commands: this.path, ...(fill ? { fill: this.state.fill, fillOpacity: this.state.fillOpacity } : {}), ...(stroke ? { stroke: this.state.stroke, ...this.strokeStyle() } : {}), fillRule: operator.endsWith('*') ? 'evenodd' : 'nonzero' });
    if (this.clipRule && this.path.length) this.add({ type: 'clip', commands: this.path, fillRule: this.clipRule });
    this.path = []; this.pathPoint = null; this.pathStart = null; this.clipRule = null;
  }

  moveLine(x, y) {
    this.lineMatrix = matrixMultiply(this.lineMatrix, [1, 0, 0, 1, x, y]);
    this.textMatrix = [...this.lineMatrix];
  }

  show(value) {
    if (!this.textActive || !isString(value) || !this.state.font) throw new Error('PDF 文本操作缺少 BT、字体或有效字符串。');
    const font = this.state.font;
    for (const glyph of this.state.font.decode(value.bytes)) {
      const width = glyph.width / 1000 * this.state.fontSize;
      const embedded = !!font.program && glyph.renderText !== undefined;
      if (glyph.text || embedded) this.add({ type: 'text', text: glyph.text, ...(embedded ? { renderText: glyph.renderText, fontId: font.id } : {}), matrix: matrixMultiply(this.textMatrix, [this.state.horizontalScale, 0, 0, 1, 0, this.state.rise]), fontSize: this.state.fontSize, fontFamily: embedded ? font.family : font.fallbackFamily || font.family, fallbackFamily: font.fallbackFamily, fontWeight: font.weight || 400, fontStyle: font.style || 'normal', color: this.state.fill, fillOpacity: this.state.fillOpacity, stroke: this.state.stroke, ...this.strokeStyle(), textMode: this.state.textMode, width });
      const advance = (width + this.state.characterSpace + (glyph.code === 32 ? this.state.wordSpace : 0)) * this.state.horizontalScale;
      this.textMatrix = matrixMultiply(this.textMatrix, [1, 0, 0, 1, advance, 0]);
    }
  }

  async font(resources, name) {
    const fonts = await this.document.resolve(resources.Font);
    const reference = fonts?.[name];
    if (!reference) throw new Error(`PDF 字体资源不存在：${name}。`);
    return this.loadFont(reference);
  }

  loadFont(reference) {
    const key = isReference(reference) ? `${reference.ref}:${reference.generation}` : reference;
    if (!this.document.fontCache.has(key)) this.document.fontCache.set(key, loadPdfFont(this.document, reference));
    return this.document.fontCache.get(key);
  }

  async image(record, resources) {
    if (!this.document.imageCache) this.document.imageCache = new Map();
    if (!this.document.imageCache.has(record)) this.document.imageCache.set(record, this.decodeImage(record, resources));
    const image = await this.document.imageCache.get(record);
    if (image) this.add({ type: 'image', ...image, fillOpacity: this.state.fillOpacity });
  }

  async decodeImage(record, resources) {
    const dictionary = record.dictionary;
    const width = integer(dictionary.Width, '图片宽度', 16000000);
    const height = integer(dictionary.Height, '图片高度', 16000000);
    if (!width || !height || width * height > 16000000) throw new Error('PDF 图片超过当前 1600 万像素上限。');
    const rawFilter = await this.document.resolve(dictionary.Filter);
    const filters = rawFilter ? Array.isArray(rawFilter) ? rawFilter : [rawFilter] : [];
    if (filters.length === 1 && ['DCTDecode', 'DCT'].includes(asName(filters[0]))) {
      const dimensions = jpegDimensions(record.stream);
      if (dimensions.width !== width || dimensions.height !== height) throw new Error('PDF JPEG 实际尺寸与图片字典不一致。');
      if (dictionary.SMask || dictionary.Mask) this.document.warn('部分 PDF JPEG 图片的透明蒙版尚未应用。');
      return { bytes: record.stream, mime: 'image/jpeg', x: 0, y: 0, width: 1, height: 1 };
    }
    if (filters.some(filter => !['FlateDecode', 'Fl', 'ASCII85Decode', 'A85', 'ASCIIHexDecode', 'AHx'].includes(asName(filter)))) {
      this.document.warn(`PDF 图片过滤器 ${filters.map(asName).join('/')} 尚未支持，该图片未显示。`); return;
    }
    if (dictionary.ImageMask || numeric(dictionary.BitsPerComponent, '图片颜色位数', 1) !== 8) {
      this.document.warn('PDF 模板图片及非 8 位图片尚未显示。'); return;
    }
    let colorSpace = await this.document.resolve(dictionary.ColorSpace);
    if (asName(colorSpace) && !['DeviceRGB', 'DeviceGray', 'DeviceCMYK'].includes(asName(colorSpace))) {
      const spaces = await this.document.resolve(resources.ColorSpace);
      colorSpace = await this.document.resolve(spaces?.[asName(colorSpace)]);
    }
    const space = asName(colorSpace);
    const channels = { DeviceRGB: 3, DeviceGray: 1, DeviceCMYK: 4 }[space];
    if (!channels) { this.document.warn('PDF 特殊图片颜色空间尚未转换，该图片未显示。'); return; }
    const bytes = await this.document.stream(record);
    if (bytes.length !== width * height * channels) throw new Error('PDF 图片数据长度与尺寸不匹配。');
    if (space === 'DeviceCMYK') this.document.warn('PDF CMYK 颜色使用近似 RGB 转换，尚未进行 ICC 色彩管理。');
    let mask;
    if (dictionary.SMask) {
      const recordMask = await this.document.resolve(dictionary.SMask);
      if (recordMask?.stream && recordMask.dictionary.Width === width && recordMask.dictionary.Height === height && recordMask.dictionary.BitsPerComponent === 8 && asName(recordMask.dictionary.ColorSpace) === 'DeviceGray') mask = await this.document.stream(recordMask);
      else this.document.warn('部分 PDF 图片透明蒙版尚未应用。');
      if (mask && mask.length !== width * height) throw new Error('PDF 图片透明蒙版长度无效。');
    }
    if (dictionary.Mask) this.document.warn('PDF 图片色键和显式蒙版尚未应用。');
    const decode = await this.document.resolve(dictionary.Decode);
    if (decode !== undefined && (!Array.isArray(decode) || decode.length !== channels * 2)) throw new Error('PDF 图片 Decode 参数无效。');
    this.document.rgbaBytes = (this.document.rgbaBytes || 0) + width * height * 4;
    if (this.document.rgbaBytes > 128 * 1024 * 1024) throw new Error('PDF 解码后图片合计超过当前 128MB 上限。');
    const rgba = new Uint8Array(width * height * 4);
    for (let pixel = 0; pixel < width * height; pixel += 1) {
      const values = [];
      for (let channel = 0; channel < channels; channel += 1) {
        let value = bytes[pixel * channels + channel] / 255;
        if (decode) value = numeric(decode[channel * 2], '图片 Decode') + value * (numeric(decode[channel * 2 + 1], '图片 Decode') - decode[channel * 2]);
        values.push(unit(value));
      }
      let color = channels === 1 ? [values[0], values[0], values[0]] : values;
      if (channels === 4) color = values.slice(0, 3).map(value => 1 - Math.min(1, value + values[3]));
      rgba.set([...color.map(value => Math.round(value * 255)), mask ? mask[pixel] : 255], pixel * 4);
    }
    return { rgba, pixelWidth: width, pixelHeight: height, x: 0, y: 0, width: 1, height: 1 };
  }

  async xobject(resources, name, depth) {
    const objects = await this.document.resolve(resources.XObject);
    const record = await this.document.resolve(objects?.[name]);
    if (!record?.stream) throw new Error(`PDF XObject 资源无效：${name}。`);
    const subtype = asName(record.dictionary.Subtype);
    if (subtype === 'Image') return this.image(record, resources);
    if (subtype !== 'Form') { this.document.warn(`PDF ${subtype || '未知'} XObject 尚未显示。`); return; }
    if (depth > 16 || this.forms.has(record)) throw new Error('PDF Form XObject 引用循环或嵌套过深。');
    if (this.path.length) throw new Error('PDF 路径尚未结束时调用 Form XObject，当前不支持此结构。');
    this.forms.add(record);
    this.save();
    const level = this.saved.length;
    const matrix = await this.document.resolve(record.dictionary.Matrix) || IDENTITY();
    if (!Array.isArray(matrix) || matrix.length !== 6) throw new Error('PDF Form 变换矩阵无效。');
    this.add({ type: 'transform', matrix: matrix.map(value => numeric(value, 'Form 矩阵')) });
    const box = await this.document.resolve(record.dictionary.BBox);
    if (!Array.isArray(box) || box.length !== 4) throw new Error('PDF Form BBox 无效。');
    this.add({ type: 'clip', commands: rectCommands(box.map(value => numeric(value, 'Form BBox'))) });
    const formResources = await this.document.resolve(record.dictionary.Resources) || resources;
    await this.run(await this.document.stream(record), formResources, depth + 1);
    if (this.saved.length !== level || this.path.length) throw new Error('PDF Form 的图形状态或路径没有正确结束。');
    this.restore(); this.forms.delete(record);
    if (record.dictionary.Group) this.document.warn('PDF 透明度组尚未完整合成。');
  }

  async run(bytes, resources, depth = 0) {
    if (!isDictionary(resources)) throw new Error('PDF 页面或 Form 资源字典无效。');
    const lexer = new PdfLexer(bytes, 0, this.document.budget);
    let operands = [];
    const numbers = count => {
      if (operands.length !== count) throw new Error('PDF 绘制操作的参数数量无效。');
      return operands.map(value => numeric(value, '绘制参数'));
    };
    const setColor = (values, space, stroke) => {
      let color;
      if (space === 'DeviceGray' && values.length === 1) color = rgb([values[0], values[0], values[0]]);
      else if (space === 'DeviceRGB' && values.length === 3) color = rgb(values);
      else if (space === 'DeviceCMYK' && values.length === 4) {
        color = rgb(values.slice(0, 3).map(value => 1 - Math.min(1, value + values[3])));
        this.document.warn('PDF CMYK 颜色使用近似 RGB 转换，尚未进行 ICC 色彩管理。');
      } else this.document.warn('PDF 特殊颜色空间或图案填充尚未完整转换。');
      if (color) this.state[stroke ? 'stroke' : 'fill'] = color;
    };
    while (true) {
      lexer.skip();
      if (lexer.position >= bytes.length) break;
      const token = lexer.value(0, false);
      if (!token?.operator) { operands.push(token); if (operands.length > 10000) throw new Error('PDF 绘制操作参数过多。'); continue; }
      const operator = token.operator;
      if (operator === 'q') { numbers(0); this.save(); }
      else if (operator === 'Q') { numbers(0); this.restore(); }
      else if (operator === 'cm') {
        const matrix = numbers(6);
        if (this.path.length) throw new Error('当前 PDF 解析器尚不支持路径构建期间修改 CTM。');
        this.add({ type: 'transform', matrix });
      } else if (operator === 'w') { this.state.lineWidth = numbers(1)[0]; if (this.state.lineWidth < 0) throw new Error('PDF 线宽无效。'); }
      else if (operator === 'J' || operator === 'j') this.state[operator === 'J' ? 'lineCap' : 'lineJoin'] = integer(numbers(1)[0], '描边端点或连接', 2);
      else if (operator === 'M') { this.state.miterLimit = numbers(1)[0]; if (this.state.miterLimit < 1) throw new Error('PDF 斜接限制无效。'); }
      else if (operator === 'd') { if (operands.length !== 2) throw new Error('PDF 虚线参数无效。'); this.setDash(operands[0], operands[1]); }
      else if (['g', 'G', 'rg', 'RG', 'k', 'K'].includes(operator)) {
        const count = { g: 1, G: 1, rg: 3, RG: 3, k: 4, K: 4 }[operator];
        const stroke = operator === operator.toUpperCase();
        const space = count === 1 ? 'DeviceGray' : count === 3 ? 'DeviceRGB' : 'DeviceCMYK';
        this.state[stroke ? 'strokeSpace' : 'fillSpace'] = space; setColor(numbers(count), space, stroke);
      } else if (operator === 'cs' || operator === 'CS') {
        if (operands.length !== 1 || !asName(operands[0])) throw new Error('PDF 颜色空间参数无效。');
        this.state[operator === 'CS' ? 'strokeSpace' : 'fillSpace'] = asName(operands[0]);
      } else if (['sc', 'SC', 'scn', 'SCN'].includes(operator)) {
        const stroke = operator === operator.toUpperCase();
        if (operands.some(value => typeof value !== 'number')) this.document.warn('PDF 图案颜色尚未绘制。');
        else setColor(numbers(operands.length), this.state[stroke ? 'strokeSpace' : 'fillSpace'], stroke);
      } else if (operator === 'm' || operator === 'l') {
        const [x, y] = numbers(2);
        if (operator === 'l' && !this.pathPoint) throw new Error('PDF 路径线段缺少起点。');
        this.command([operator === 'm' ? 'M' : 'L', x, y]); this.pathPoint = [x, y];
        if (operator === 'm') this.pathStart = [x, y];
      } else if (operator === 'c') {
        const values = numbers(6); if (!this.pathPoint) throw new Error('PDF 曲线缺少起点。');
        this.command(['C', ...values]); this.pathPoint = values.slice(-2);
      } else if (operator === 'v' || operator === 'y') {
        const [a, b, x, y] = numbers(4); if (!this.pathPoint) throw new Error('PDF 曲线缺少起点。');
        this.command(operator === 'v' ? ['C', ...this.pathPoint, a, b, x, y] : ['C', a, b, x, y, x, y]); this.pathPoint = [x, y];
      } else if (operator === 'h') { numbers(0); this.command(['Z']); this.pathPoint = this.pathStart; }
      else if (operator === 're') {
        const [x, y, width, height] = numbers(4);
        for (const command of rectCommands([x, y, x + width, y + height])) this.command(command);
        this.pathPoint = this.pathStart = [x, y];
      } else if (['S', 's', 'f', 'F', 'f*', 'B', 'B*', 'b', 'b*', 'n'].includes(operator)) { numbers(0); this.paint(operator); }
      else if (operator === 'W' || operator === 'W*') { numbers(0); this.clipRule = operator === 'W*' ? 'evenodd' : 'nonzero'; }
      else if (operator === 'BT') {
        numbers(0); if (this.textActive) throw new Error('PDF BT 文本对象不能嵌套。'); this.textActive = true; this.textMatrix = this.lineMatrix = IDENTITY();
      } else if (operator === 'ET') { numbers(0); if (!this.textActive) throw new Error('PDF ET 缺少 BT。'); this.textActive = false; }
      else if (operator === 'Tf') {
        if (operands.length !== 2 || !asName(operands[0])) throw new Error('PDF Tf 字体参数无效。');
        this.state.font = await this.font(resources, asName(operands[0])); this.state.fontSize = numeric(operands[1], '字号');
        if (this.state.fontSize <= 0) throw new Error('PDF 字号无效或当前尚不支持负字号。');
      } else if (operator === 'Tm') { this.textMatrix = numbers(6); this.lineMatrix = [...this.textMatrix]; }
      else if (operator === 'Td' || operator === 'TD') { const [x, y] = numbers(2); if (operator === 'TD') this.state.leading = -y; this.moveLine(x, y); }
      else if (operator === 'T*') { numbers(0); this.moveLine(0, -this.state.leading); }
      else if (['Tc', 'Tw', 'TL', 'Ts', 'Tz'].includes(operator)) {
        const value = numbers(1)[0]; const property = { Tc: 'characterSpace', Tw: 'wordSpace', TL: 'leading', Ts: 'rise', Tz: 'horizontalScale' }[operator];
        this.state[property] = operator === 'Tz' ? value / 100 : value;
      } else if (operator === 'Tr') {
        this.state.textMode = integer(numbers(1)[0], '文本渲染模式', 7);
        if (this.state.textMode >= 4) this.document.warn('PDF 文字裁剪尚未应用，仅绘制该模式的可见文字部分。');
      } else if (operator === 'Tj') { if (operands.length !== 1) throw new Error('PDF Tj 参数无效。'); this.show(operands[0]); }
      else if (operator === 'TJ') {
        if (operands.length !== 1 || !Array.isArray(operands[0])) throw new Error('PDF TJ 参数无效。');
        for (const item of operands[0]) {
          if (isString(item)) this.show(item);
          else this.textMatrix = matrixMultiply(this.textMatrix, [1, 0, 0, 1, -numeric(item, 'TJ 字距') / 1000 * this.state.fontSize * this.state.horizontalScale, 0]);
        }
      } else if (operator === "'" || operator === '"') {
        if (operands.length !== (operator === "'" ? 1 : 3)) throw new Error('PDF 换行文字操作参数无效。');
        if (operator === '"') { this.state.wordSpace = numeric(operands[0], '单词间距'); this.state.characterSpace = numeric(operands[1], '字符间距'); }
        this.moveLine(0, -this.state.leading); this.show(operands[operands.length - 1]);
      } else if (operator === 'Do') {
        if (operands.length !== 1 || !asName(operands[0])) throw new Error('PDF Do 资源名无效。');
        await this.xobject(resources, asName(operands[0]), depth);
      } else if (operator === 'gs') {
        if (operands.length !== 1 || !asName(operands[0])) throw new Error('PDF gs 参数无效。');
        const states = await this.document.resolve(resources.ExtGState);
        const state = await this.document.resolve(states?.[asName(operands[0])]);
        if (!isDictionary(state)) throw new Error('PDF ExtGState 资源无效。');
        if (state.LW !== undefined) { this.state.lineWidth = numeric(await this.document.resolve(state.LW), 'ExtGState 线宽'); if (this.state.lineWidth < 0) throw new Error('PDF 线宽无效。'); }
        if (state.LC !== undefined) this.state.lineCap = integer(await this.document.resolve(state.LC), 'ExtGState 端点', 2);
        if (state.LJ !== undefined) this.state.lineJoin = integer(await this.document.resolve(state.LJ), 'ExtGState 连接', 2);
        if (state.ML !== undefined) { this.state.miterLimit = numeric(await this.document.resolve(state.ML), 'ExtGState 斜接限制'); if (this.state.miterLimit < 1) throw new Error('PDF 斜接限制无效。'); }
        if (state.D !== undefined) {
          const dash = await this.document.resolve(state.D);
          if (!Array.isArray(dash) || dash.length !== 2) throw new Error('PDF ExtGState 虚线参数无效。');
          this.setDash(await this.document.resolve(dash[0]), await this.document.resolve(dash[1]));
        }
        for (const [entry, property] of [['ca', 'fillOpacity'], ['CA', 'strokeOpacity']]) if (state[entry] !== undefined) {
          const alpha = numeric(await this.document.resolve(state[entry]), 'ExtGState 透明度');
          if (alpha < 0 || alpha > 1) throw new Error('PDF 透明度必须位于 0 和 1 之间。');
          this.state[property] = alpha;
        }
        const blend = await this.document.resolve(state.BM);
        if (state.SMask && asName(await this.document.resolve(state.SMask)) !== 'None' || blend && !(Array.isArray(blend) ? blend : [blend]).some(mode => ['Normal', 'Compatible'].includes(asName(mode))) || state.AIS === true) this.document.warn('PDF 软蒙版、形状透明度和特殊混合模式尚未完整合成。');
        if (state.Font) {
          const font = await this.document.resolve(state.Font);
          if (!Array.isArray(font) || font.length !== 2) throw new Error('PDF ExtGState 字体参数无效。');
          this.state.font = await this.loadFont(font[0]); this.state.fontSize = numeric(await this.document.resolve(font[1]), '字号');
          if (this.state.fontSize <= 0) throw new Error('PDF 字号无效或当前尚不支持负字号。');
        }
      } else if (operator === 'BI') throw new Error('当前尚不支持 PDF 内容流中的内联图像；支持独立 Image XObject。');
      else if (['BMC', 'BDC', 'EMC', 'MP', 'DP', 'BX', 'EX', 'ri', 'i'].includes(operator)) {
        if (operator === 'BDC' && asName(operands[0]) === 'OC') this.document.warn('PDF 可选图层的可见性尚未计算。');
      } else if (operator === 'sh') this.document.warn('PDF 渐变着色尚未绘制。');
      else this.document.warn(`PDF 绘制操作 ${operator} 尚未实现，相关内容可能缺失。`);
      operands = [];
    }
    if (operands.length) throw new Error('PDF 内容流末尾残留了没有操作符的参数。');
  }

  finish() {
    if (this.saved.length || this.textActive || this.path.length) throw new Error('PDF 内容流中的图形状态、文字或路径没有正确结束。');
  }
}
