import { binary, warn } from './legacy-binary.js';

export function createDocImageReader(bytes, warnings) {
  const cache = new Map();
  let total = 0, restored = 0;
  if (!bytes) return { read() { return null; }, get restored() { return 0; } };
  const source = binary(bytes);
  const record = (position, boundary) => {
    source.need(position, 8);
    const flags = source.u16(position), size = source.u32(position + 4), end = position + 8 + size;
    if (end > boundary) throw new Error('DOC 图片绘图记录超出 PICF 边界。');
    return { position, start: position + 8, size, end, type: source.u16(position + 2), instance: flags >>> 4, version: flags & 15 };
  };
  function decodeBlip(item) {
    let mime, count;
    if (item.type === 0xf01e && [0x6e0, 0x6e1].includes(item.instance)) { mime = 'image/png'; count = item.instance === 0x6e1 ? 2 : 1; }
    else if ([0xf01d, 0xf02a].includes(item.type) && [0x46a, 0x46b, 0x6e2, 0x6e3].includes(item.instance)) { mime = 'image/jpeg'; count = item.instance & 1 ? 2 : 1; }
    else if (item.type === 0xf01f && [0x7a8, 0x7a9].includes(item.instance)) { mime = 'image/bmp'; count = item.instance === 0x7a9 ? 2 : 1; }
    else { warn(warnings, 'DOC 中部分 EMF / WMF / PICT 等图像编码尚未显示。'); return null; }
    const start = item.start + count * 16 + 1;
    if (start >= item.end) throw new Error('DOC 图片 BLIP 没有有效负载。');
    let data = source.slice(start, item.end - start);
    let width, height;
    if (mime === 'image/png') {
      if (data.length < 24 || ![137,80,78,71,13,10,26,10].every((value, index) => data[index] === value)) throw new Error('DOC PNG 图片标识无效。');
      const view = new DataView(data.buffer, data.byteOffset, data.byteLength); width = view.getUint32(16); height = view.getUint32(20);
    } else if (mime === 'image/jpeg') {
      if (data[0] !== 255 || data[1] !== 216) throw new Error('DOC JPEG 图片标识无效。');
      let position = 2;
      while (position + 3 < data.length) {
        if (data[position++] !== 255) throw new Error('DOC JPEG 标记无效。');
        while (position < data.length && data[position] === 255) position++;
        const marker = data[position++];
        if (marker === 0xd9 || marker === 0xda) break;
        if (marker === 0xd8 || marker >= 0xd0 && marker <= 0xd7 || marker === 1) continue;
        const length = data[position] * 256 + data[position + 1];
        if (length < 2 || position + length > data.length) throw new Error('DOC JPEG 段长度无效。');
        if (marker >= 0xc0 && marker <= 0xcf && ![0xc4,0xc8,0xcc].includes(marker)) {
          if (length < 8) throw new Error('DOC JPEG 尺寸记录不完整。');
          height = data[position + 3] * 256 + data[position + 4]; width = data[position + 5] * 256 + data[position + 6]; break;
        }
        position += length;
      }
    } else {
      const dib = binary(data), size = dib.u32(0), compression = dib.u32(16), bits = dib.u16(14);
      if (size < 40 || size > data.length || ![0,3].includes(compression)) { warn(warnings, 'DOC 中部分旧式或压缩 DIB 图片尚未显示。'); return null; }
      width = Math.abs(dib.u32(4) | 0); height = Math.abs(dib.u32(8) | 0);
      const colors = dib.u32(32) || (bits <= 8 ? 2 ** bits : 0);
      const offset = 14 + size + colors * 4 + (compression === 3 && size === 40 ? 12 : 0);
      if (offset > data.length + 14) throw new Error('DOC DIB 图片颜色表不完整。');
      const bmp = new Uint8Array(data.length + 14), view = new DataView(bmp.buffer);
      bmp[0] = 66; bmp[1] = 77; view.setUint32(2, bmp.length, true); view.setUint32(10, offset, true); bmp.set(data, 14); data = bmp;
    }
    if (!width || !height || width * height > 16000000) throw new Error('DOC 图片尺寸无效或超过当前像素上限。');
    total += data.length;
    if (data.length > 32 * 1024 * 1024 || total > 128 * 1024 * 1024) throw new Error('DOC 图片数据超过当前容量上限。');
    return { bytes: data, mime };
  }
  return {
    read(position) {
      if (position === undefined) return null;
      if (cache.has(position)) { const image = cache.get(position); if (image) restored++; return image; }
      source.need(position, 68);
      const length = source.u32(position), end = position + length, header = source.u16(position + 4), format = source.u16(position + 6);
      if (header !== 68 || length < 76 || end > bytes.length) throw new Error('DOC 图片 PICF 长度或头无效。');
      if (![100,102].includes(format)) { warn(warnings, 'DOC 中部分行内图像存储格式尚未支持。'); cache.set(position, null); return null; }
      const goalWidth = source.u16(position + 28), goalHeight = source.u16(position + 30);
      const width = goalWidth * source.u16(position + 32) / 15000, height = goalHeight * source.u16(position + 34) / 15000;
      if (!(width > 0 && height > 0 && width <= 2112 && height <= 2112)) throw new Error('DOC 图片显示尺寸无效。');
      let cursor = position + 68;
      if (format === 102) { const nameBytes = source.u8(cursor++); cursor += nameBytes; }
      const shape = record(cursor, end);
      if (shape.type !== 0xf004 || shape.version !== 15) throw new Error('DOC 图片缺少行内形状容器。');
      cursor = shape.end;
      let image = null;
      while (cursor < end) {
        let item = record(cursor, end);
        if (item.type === 0xf007) {
          if (item.size < 36) throw new Error('DOC 图片 BSE 目录头不完整。');
          const nameLength = source.u8(item.start + 33), embedded = item.start + 36 + nameLength;
          if (embedded < item.end) item = record(embedded, item.end);
          else { warn(warnings, 'DOC 中部分图片 BSE 没有内嵌图像负载。'); cursor = item.end; continue; }
        }
        if (!image) image = decodeBlip(item);
        cursor = record(cursor, end).end;
      }
      const result = image ? { ...image, width, height } : null;
      cache.set(position, result);
      if (result) restored++;
      return result;
    },
    get restored() { return restored; }
  };
}
