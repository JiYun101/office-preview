// MS-CFB sector allocation and mini-stream reader; no Office runtime is used.
const END = 0xfffffffe;
const FREE = 0xffffffff;
const SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];

export function readCompound(input) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  if (bytes.length < 512 || !SIGNATURE.every((value, i) => bytes[i] === value)) throw new Error('不是有效的 OLE 复合文档。');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const u16 = offset => view.getUint16(offset, true);
  const u32 = offset => view.getUint32(offset, true);
  const major = u16(26);
  const shift = u16(30);
  if (u16(28) !== 0xfffe || !((major === 3 && shift === 9) || (major === 4 && shift === 12)) || u16(32) !== 6) throw new Error('不支持的 OLE 容器版本。');
  const sectorSize = 2 ** shift;
  const sectorCount = Math.ceil(bytes.length / sectorSize) - 1;
  if (sectorCount < 1) throw new Error('OLE 扇区数据不完整。');
  const sector = (id, allowPartial = false) => {
    if (!Number.isInteger(id) || id < 0 || id >= sectorCount) throw new Error('OLE 扇区索引越界。');
    const value = bytes.subarray((id + 1) * sectorSize, (id + 2) * sectorSize);
    if (!allowPartial && value.length !== sectorSize) throw new Error('OLE 结构扇区数据不完整。');
    return value;
  };
  const ints = data => {
    const dv = new DataView(data.buffer, data.byteOffset, data.byteLength);
    return Array.from({ length: data.length / 4 }, (_, i) => dv.getUint32(i * 4, true));
  };
  const fatCount = u32(44);
  if (fatCount > sectorCount || u32(72) > sectorCount || u32(64) > sectorCount) throw new Error('OLE 分配表数量无效。');
  const fatSectors = [];
  for (let i = 0; i < 109; i++) if (u32(76 + i * 4) !== FREE) fatSectors.push(u32(76 + i * 4));
  let difat = u32(68);
  const difatSeen = new Set();
  for (let i = 0; i < u32(72); i++) {
    if (difatSeen.has(difat)) throw new Error('OLE DIFAT 包含循环引用。');
    difatSeen.add(difat);
    const list = ints(sector(difat));
    for (const id of list.slice(0, -1)) if (id !== FREE) fatSectors.push(id);
    difat = list.at(-1);
  }
  if (fatSectors.length !== fatCount || new Set(fatSectors).size !== fatCount) throw new Error('OLE FAT 数量不匹配。');
  const fat = fatSectors.flatMap(id => ints(sector(id)));
  const chain = (start, table, source, chunkSize, expected) => {
    const parts = [];
    const seen = new Set();
    let id = start;
    let total = 0;
    while (id !== END) {
      if (id === FREE || seen.has(id) || id >= table.length) throw new Error('OLE 扇区链损坏或循环。');
      seen.add(id);
      const part = source(id);
      // Some writers omit unused padding at the physical end of a user data
      // stream. Its declared bytes must still exist, and a short sector cannot
      // appear in the middle of a chain or replace a structural sector.
      if (part.length !== chunkSize && (expected === undefined || table[id] !== END || total + part.length < expected)) throw new Error('OLE 流被截断。');
      parts.push(part);
      total += part.length;
      if (parts.length > Math.ceil(bytes.length / chunkSize) || total > bytes.length + chunkSize) throw new Error('OLE 流长度无效。');
      id = table[id];
    }
    if (expected !== undefined && total < expected) throw new Error('OLE 流被截断。');
    const length = expected ?? total;
    const result = new Uint8Array(length);
    let offset = 0;
    for (const part of parts) {
      const count = Math.min(part.length, length - offset);
      if (count <= 0) break;
      result.set(part.subarray(0, count), offset);
      offset += count;
    }
    return result;
  };
  const directory = chain(u32(48), fat, sector, sectorSize);
  const entries = [];
  for (let offset = 0; offset + 128 <= directory.length; offset += 128) {
    const dv = new DataView(directory.buffer, directory.byteOffset + offset, 128);
    const type = dv.getUint8(66);
    const nameBytes = dv.getUint16(64, true);
    if (type === 0) { entries.push(null); continue; }
    if (nameBytes < 2 || nameBytes > 64 || nameBytes % 2) throw new Error('OLE 目录名称无效。');
    const size = dv.getUint32(120, true) + (major === 4 ? dv.getUint32(124, true) * 4294967296 : 0);
    if (size > bytes.length) throw new Error('OLE 流大小无效。');
    entries.push({
      name: new TextDecoder('utf-16le').decode(directory.subarray(offset, offset + nameBytes - 2)),
      type, size, start: dv.getUint32(116, true),
      left: dv.getUint32(68, true), right: dv.getUint32(72, true), child: dv.getUint32(76, true)
    });
  }
  const root = entries[0];
  if (!root || root.type !== 5) throw new Error('OLE 缺少根目录。');
  const dataSector = id => sector(id, true);
  const miniData = root.size ? chain(root.start, fat, dataSector, sectorSize, root.size) : new Uint8Array();
  const miniFat = u32(64) ? ints(chain(u32(60), fat, sector, sectorSize, u32(64) * sectorSize)) : [];
  if (u32(56) !== 4096) throw new Error('OLE mini-stream 阈值无效。');
  const streams = new Map();
  const visited = new Set();
  const pending = [[root.child, '']];
  while (pending.length) {
    const [id, parent] = pending.pop();
    if (id === FREE) continue;
    if (visited.has(id) || !entries[id]) throw new Error('OLE 目录树损坏或循环。');
    visited.add(id);
    const entry = entries[id];
    const path = parent + entry.name;
    pending.push([entry.left, parent], [entry.right, parent]);
    if (entry.type === 1) pending.push([entry.child, path + '/']);
    else if (entry.type === 2) streams.set(path, entry);
    else throw new Error('OLE 目录对象类型无效。');
  }
  const cache = new Map();
  return {
    names: [...streams.keys()],
    get(name) {
      const key = streams.has(name) ? name : [...streams.keys()].find(key => key.toLowerCase() === name.toLowerCase());
      if (!key) return undefined;
      if (cache.has(key)) return cache.get(key);
      const entry = streams.get(key);
      let value;
      if (!entry.size) value = new Uint8Array();
      else if (entry.size >= 4096) value = chain(entry.start, fat, dataSector, sectorSize, entry.size);
      else value = chain(entry.start, miniFat, id => {
        if ((id + 1) * 64 > miniData.length) throw new Error('OLE mini-stream 索引越界。');
        return miniData.subarray(id * 64, (id + 1) * 64);
      }, 64, entry.size);
      cache.set(key, value);
      return value;
    }
  };
}
