// ZIP container reader. Compression is provided by the browser's standard API.
const MAX_ENTRY = 32 * 1024 * 1024;
const MAX_TOTAL = 128 * 1024 * 1024;

export function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (0xedb88320 & -(crc & 1));
  }
  return (crc ^ 0xffffffff) >>> 0;
}

async function inflate(bytes, expected) {
  let stream;
  try {
    stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  } catch {
    throw new Error('当前浏览器不支持原生 DEFLATE 解压，请更新 Safari / WKWebView、Chrome 或 Edge / WebView2。');
  }
  const reader = stream.getReader();
  const chunks = [];
  let length = 0;
  try {
    while (true) {
      const { value, done } = await reader.read();
      if (done) break;
      length += value.length;
      if (length > expected || length > MAX_ENTRY) throw new Error('ZIP 解压大小超出声明值或首版限制。');
      chunks.push(value);
    }
  } catch (error) {
    await reader.cancel().catch(() => {});
    throw error;
  }
  const output = new Uint8Array(length);
  let offset = 0;
  for (const chunk of chunks) { output.set(chunk, offset); offset += chunk.length; }
  return output;
}

export async function readZip(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const requireRange = (offset, length) => {
    if (offset < 0 || length < 0 || offset + length > bytes.length) throw new Error('ZIP 文件结构不完整。');
  };
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === bytes.length) { end = i; break; }
  }
  if (end < 0) throw new Error('不是有效的 ZIP / OOXML 文档，或文件已损坏。');
  if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true)) throw new Error('首版不支持分卷 ZIP。');
  const count = view.getUint16(end + 10, true);
  const centralSize = view.getUint32(end + 12, true);
  let cursor = view.getUint32(end + 16, true);
  if (count === 65535 || cursor === 0xffffffff || centralSize === 0xffffffff) throw new Error('首版不支持 ZIP64。');
  if (count > 10000) throw new Error('文档内部文件数量超过首版 10000 项限制。');
  requireRange(cursor, centralSize);
  if (cursor + centralSize > end) throw new Error('ZIP 中央目录位置无效。');
  const centralStart = cursor;
  const centralEnd = cursor + centralSize;
  const files = new Map();
  let total = 0;
  for (let index = 0; index < count; index++) {
    requireRange(cursor, 46);
    if (view.getUint32(cursor, true) !== 0x02014b50) throw new Error('ZIP 中央目录损坏。');
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    const checksum = view.getUint32(cursor + 16, true);
    const compressed = view.getUint32(cursor + 20, true);
    const size = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    const local = view.getUint32(cursor + 42, true);
    requireRange(cursor + 46, nameLength + extraLength + commentLength);
    if (cursor + 46 + nameLength + extraLength + commentLength > centralEnd) throw new Error('ZIP 中央目录条目越界。');
    const name = new TextDecoder().decode(bytes.subarray(cursor + 46, cursor + 46 + nameLength)).replaceAll('\\', '/');
    if (flags & 1) throw new Error('暂不支持加密 ZIP 文档。');
    if (size === 0xffffffff || compressed === 0xffffffff || local === 0xffffffff) throw new Error('首版不支持 ZIP64 条目。');
    if (size > MAX_ENTRY || (total += size) > MAX_TOTAL) throw new Error('文档解压后超过首版容量限制（单项32MB，总计128MB）。');
    if (name.startsWith('/') || name.split('/').includes('..') || files.has(name)) throw new Error('ZIP 包含无效或重复的内部路径。');
    requireRange(local, 30);
    if (local + 30 > centralStart || view.getUint32(local, true) !== 0x04034b50 || view.getUint16(local + 8, true) !== method || view.getUint16(local + 6, true) !== flags) throw new Error('ZIP 条目头损坏。');
    const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
    requireRange(start, compressed);
    if (start + compressed > centralStart) throw new Error('ZIP 数据与目录重叠。');
    const localName = new TextDecoder().decode(bytes.subarray(local + 30, local + 30 + view.getUint16(local + 26, true))).replaceAll('\\', '/');
    if (localName !== name) throw new Error('ZIP 条目名称与中央目录不一致。');
    let output;
    if (method === 0) output = bytes.slice(start, start + compressed);
    else if (method === 8) output = await inflate(bytes.subarray(start, start + compressed), size);
    else throw new Error(`暂不支持 ZIP 压缩算法 ${method}。`);
    if (output.length !== size || crc32(output) !== checksum) throw new Error(`文档内部文件校验失败：${name}`);
    files.set(name, output);
    cursor += 46 + nameLength + extraLength + commentLength;
    if (cursor > centralEnd) throw new Error('ZIP 中央目录长度不匹配。');
  }
  return files;
}
