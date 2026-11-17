import { binary } from './legacy-binary.js';
import { createPptArtReader } from './legacy-ppt-art.js';

export function parsePpt(compound, name) {
  const stream = compound.get('PowerPoint Document');
  const currentUser = compound.get('Current User');
  if (!stream) throw new Error('文件缺少 PowerPoint Document 流，不是受支持的 PPT。');
  if (!currentUser) throw new Error('PPT 缺少 Current User 当前版本索引，无法可靠识别有效幻灯片。');
  const source = binary(stream);
  const user = binary(currentUser);
  if (user.u16(2) !== 0x0ff6) throw new Error('PPT Current User 记录标识无效。');
  const token = user.u32(12);
  if (token === 0xf3d1c4df) throw new Error('PPT 文稿已加密，当前版本不支持解密。');
  if (token !== 0xe391c05f) throw new Error('PPT Current User 加密标识未知。');
  const header = (offset, boundary = stream.length) => {
    source.need(offset, 8);
    const flags = source.u16(offset);
    const length = source.u32(offset + 4);
    const record = { offset, data: offset + 8, end: offset + 8 + length, length, type: source.u16(offset + 2), version: flags & 15, instance: flags >>> 4 };
    if (record.end > boundary) throw new Error('PPT 记录超出所属容器，文件可能损坏。');
    if (record.type === 0x2f14) throw new Error('PPT 文稿包含加密会话，当前版本不支持解密。');
    return record;
  };
  const children = record => {
    const result = [];
    for (let offset = record.data; offset < record.end;) {
      const child = header(offset, record.end);
      result.push(child);
      offset = child.end;
    }
    return result;
  };
  const persists = new Map();
  let editOffset = user.u32(16);
  let documentId;
  const seenEdits = new Set();
  while (editOffset) {
    if (seenEdits.has(editOffset)) throw new Error('PPT 当前编辑索引包含循环引用。');
    seenEdits.add(editOffset);
    const edit = header(editOffset);
    if (edit.type !== 0x0ff5 || (edit.length !== 28 && edit.length !== 32)) throw new Error('PPT UserEditAtom 当前编辑记录无效。');
    if (edit.length === 32 && source.u32(edit.data + 28)) throw new Error('PPT 文稿已加密，当前版本不支持解密。');
    if (documentId === undefined) documentId = source.u32(edit.data + 16);
    const directory = header(source.u32(edit.data + 12));
    if (directory.type !== 0x1772) throw new Error('PPT 持久对象索引记录无效。');
    let cursor = directory.data;
    while (cursor < directory.end) {
      const descriptor = source.u32(cursor);
      const firstId = descriptor & 0xfffff;
      const count = descriptor >>> 20;
      cursor += 4;
      if (!count || cursor + count * 4 > directory.end) throw new Error('PPT 持久对象索引长度无效。');
      for (let index = 0; index < count; index++) {
        const id = firstId + index;
        if (!persists.has(id)) persists.set(id, source.u32(cursor + index * 4));
      }
      cursor += count * 4;
    }
    const previous = source.u32(edit.data + 8);
    if (previous >= editOffset) throw new Error('PPT 当前编辑索引的前序偏移无效。');
    editOffset = previous;
  }
  if (!persists.has(documentId)) throw new Error('PPT 持久对象目录没有当前 DocumentContainer。');
  const document = header(persists.get(documentId));
  if (document.type !== 0x03e8 || document.version !== 15) throw new Error('PPT 当前文档容器无效。');
  const warnings = ['PPT 当前按原始坐标还原基础文字、形状和图片；复杂图表、OLE、动画与媒体尚未完整实现。'];
  const art = createPptArtReader({ source, header, children, document, compound, persists, warnings });
  const entries = [];
  for (const list of children(document).filter(record => record.type === 0x0ff0 && record.instance === 0)) {
    let entry;
    for (const record of children(list)) {
      if (record.type === 0x03f3) {
        if (record.length < 4) throw new Error('PPT SlidePersistAtom 记录不完整。');
        entry = { id: source.u32(record.data), records: [] };
        entries.push(entry);
      } else if (entry) entry.records.push(record);
    }
  }
  if (entries.length > 500) throw new Error('PPT 超过当前 500 页预览上限。');
  const slides = entries.map((entry, index) => {
    if (!persists.has(entry.id)) throw new Error('PPT 幻灯片引用了缺失的持久对象。');
    const record = header(persists.get(entry.id));
    if (record.type !== 0x03ee || record.version !== 15) throw new Error('PPT 幻灯片持久对象类型无效。');
    const result = art.renderSlide(record, { blocks: art.textReader.readBlocks(entry.records) });
    return { name: `幻灯片 ${index + 1}`, background: result.background, elements: result.elements };
  });
  return { kind: 'presentation', format: 'ppt', name, width: art.width, height: art.height, slides, warnings };
}
