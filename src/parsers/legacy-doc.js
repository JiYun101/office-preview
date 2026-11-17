import { binary, utf16, warn } from './legacy-binary.js';
import { createDocStyleReader } from './legacy-doc-style.js';
import { createDocLists } from './legacy-doc-list.js';
import { createDocImageReader } from './legacy-doc-image.js';

const wordAnsi = new TextDecoder('windows-1252');
const paragraphKeys = ['align', 'indentLeft', 'indentRight', 'firstLineIndent', 'spaceBefore', 'spaceAfter', 'lineHeight', 'lineHeightRatio', 'lineHeightRule', 'tabs', 'pageBreakBefore', 'keepNext', 'keepLines', 'snapToGrid'];

export function parseDoc(compound, name) {
  const bytes = compound.get('WordDocument');
  if (!bytes) throw new Error('文件缺少 WordDocument 流，不是受支持的 DOC 文档。');
  const word = binary(bytes);
  if (word.u16(0) !== 0xa5ec) throw new Error('DOC 文件的 FIB 标识无效。');
  const version = word.u16(2);
  if (version < 0x00c1 || version > 0x0112) throw new Error('当前 DOC 解析器仅支持 Word 97–2003 系列二进制结构。');
  const flags = word.u16(10);
  if (flags & 0x0100) throw new Error('DOC 文档已加密或经过密码混淆，当前版本不支持解密。');
  const tableBytes = compound.get(flags & 0x0200 ? '1Table' : '0Table');
  if (!tableBytes) throw new Error('DOC 文件缺少 FIB 指定的文字索引表流。');
  const table = binary(tableBytes), csw = word.u16(32), cslwPosition = 34 + csw * 2, cslw = word.u16(cslwPosition);
  if (csw !== 14 || cslw < 22) throw new Error('当前版本不能解析该 DOC 的 FIB 结构。');
  const longWords = cslwPosition + 2, characterCount = word.u32(longWords + 12), pairCountPosition = longWords + cslw * 4, pairCount = word.u16(pairCountPosition);
  if (characterCount > 5000000) throw new Error('DOC 正文超过当前五百万字符预览上限。');
  if (pairCount < 34) throw new Error('DOC 的文字索引字段不完整。');
  const pair = index => index < pairCount ? { offset: word.u32(pairCountPosition + 2 + index * 8), length: word.u32(pairCountPosition + 6 + index * 8) } : { offset: 0, length: 0 };
  const warnings = ['DOC 已读取基本字符和段落样式；浮动环绕、完整表格、页眉页脚及精确原生分页仍不完整。'];
  if (!characterCount) return { kind: 'document', format: 'doc', name, warnings, blocks: [] };
  const clx = pair(33); table.need(clx.offset, clx.length);
  let cursor = clx.offset;
  const clxEnd = cursor + clx.length;
  while (cursor < clxEnd && table.u8(cursor) === 1) cursor += 3 + table.u16(cursor + 1);
  if (cursor + 5 > clxEnd || table.u8(cursor) !== 2) throw new Error('DOC 中没有受支持的 CLX PieceTable，无法可靠读取正文。');
  const pieceLength = table.u32(cursor + 1), piecesPosition = cursor + 5;
  if (pieceLength < 4 || (pieceLength - 4) % 12 || piecesPosition + pieceLength > clxEnd) throw new Error('DOC 的 PieceTable 长度无效。');
  const count = (pieceLength - 4) / 12, descriptors = piecesPosition + (count + 1) * 4;
  if (table.u32(piecesPosition) !== 0 || table.u32(piecesPosition + count * 4) < characterCount) throw new Error('DOC 的 PieceTable 没有覆盖完整正文。');
  const pieces = [], fragments = [];
  for (let index = 0; index < count; index++) {
    const start = table.u32(piecesPosition + index * 4), fullEnd = table.u32(piecesPosition + (index + 1) * 4);
    if (fullEnd < start) throw new Error('DOC 的文字位置索引顺序无效。');
    if (start >= characterCount) continue;
    const end = Math.min(fullEnd, characterCount), compressedOffset = table.u32(descriptors + index * 8 + 2), compressed = !!(compressedOffset & 0x40000000), width = compressed ? 1 : 2;
    const offset = (compressedOffset & 0x3fffffff) / (compressed ? 2 : 1);
    if (!Number.isInteger(offset)) throw new Error('DOC 的压缩文字偏移无效。');
    const data = word.slice(offset, (end - start) * width);
    pieces.push({ start, end, offset, width }); fragments.push(compressed ? wordAnsi.decode(data) : utf16.decode(data));
  }
  const text = fragments.join('');
  if (text.length !== characterCount) throw new Error('DOC 正文字符数与索引不一致。');
  const styles = createDocStyleReader(bytes, tableBytes, pair, compound.get('Data'), warnings);
  const lists = createDocLists(tableBytes, pair, pairCount, warnings);
  const imageReader = createDocImageReader(compound.get('Data'), warnings);
  const pieceAt = cp => {
    let low = 0, high = pieces.length;
    while (low < high) { const middle = (low + high) >>> 1; if (pieces[middle].start <= cp) low = middle + 1; else high = middle; }
    const piece = pieces[low - 1];
    if (!piece || cp >= piece.end) throw new Error('DOC 字符位置不在正文 PieceTable 中。');
    return piece;
  };
  const fcAt = cp => { const piece = pieceAt(cp); return piece.offset + (cp - piece.start) * piece.width; };
  const visible = new Uint8Array(characterCount), tocContext = new Uint8Array(characterCount), fields = [];
  let hidden = 0, tocDepth = 0, pictures = 0, cells = 0;
  for (let cp = 0; cp < text.length; cp++) {
    const character = text[cp];
    if (character === '\x13') { if (fields.length >= 64) throw new Error('DOC 域的嵌套层数超过当前预览上限。'); fields.push({ start: cp, result: false, toc: false }); hidden++; }
    else if (character === '\x14') {
      const field = fields[fields.length - 1];
      if (field && !field.result) { field.result = true; hidden--; field.toc = /^\s*TOC(?:\s|$)/i.test(text.slice(field.start + 1, cp)); if (field.toc) tocDepth++; }
    } else if (character === '\x15') { const field = fields.pop(); if (field && !field.result) hidden--; if (field?.toc) tocDepth--; }
    else if (!hidden) { visible[cp] = 1; tocContext[cp] = tocDepth ? 1 : 0; if (character === '\x01') pictures++; if (character === '\x07') cells++; }
  }
  const blocks = [];
  let start = 0;
  for (let cp = 0; cp <= text.length; cp++) {
    if (cp < text.length && text[cp] !== '\r') continue;
    const last = Math.min(cp, text.length - 1), paragraph = styles.paragraph(fcAt(last));
    const runs = [];
    let position = start;
    while (position < cp) {
      const piece = pieceAt(position), fc = fcAt(position), character = styles.character(fc, paragraph.character);
      const tocCharacter = styles.character(fc, paragraph.character, true);
      const end = Math.min(cp, piece.end, position + Math.max(1, Math.floor((character.end - fc) / piece.width)));
      let segment = '', category, segmentProperties = character.properties;
      const flush = () => { if (segment) { runs.push({ text: segment, ...styles.publicCharacter(segmentProperties, segment) }); segment = ''; } };
      for (let current = position; current < end; current++) {
        const properties = tocContext[current] ? tocCharacter.properties : character.properties;
        if (segmentProperties !== properties) { flush(); segmentProperties = properties; }
        if (!visible[current] || properties._hidden) { flush(); continue; }
        const ch = text[current];
        if (ch === '\t' || ch === '\x0c') { flush(); runs.push({ type: ch === '\t' ? 'tab' : 'pageBreak', text: ch === '\t' ? '\t' : '\f', ...styles.publicCharacter(properties, ch === '\t' ? '\t' : '中') }); continue; }
        if (ch === '\x01') {
          flush();
          if (!properties._pictureIsData) {
            const image = imageReader.read(properties._pictureLocation);
            if (image) runs.push({ text: '', image });
          }
          continue;
        }
        if (ch === '\x08') { flush(); continue; }
        const output = ch === '\x0b' ? '\n' : ch === '\x07' ? '\t' : ch;
        if (output.charCodeAt(0) < 32 && output !== '\n' && output !== '\t') { flush(); continue; }
        const kind = /[\u2e80-\u9fff\uf900-\ufaff\uff00-\uffef]/u.test(ch);
        if (category !== undefined && category !== kind) flush();
        category = kind; segment += output;
      }
      flush(); position = end;
    }
    const style = styles.publicCharacter(paragraph.character, '中');
    const list = runs.length || text[cp] === '\r' && visible[cp] ? lists.apply(paragraph, paragraph.character) : null;
    const properties = list?.paragraph || paragraph.paragraph;
    const block = { type: 'paragraph', runs, style, styleId: String(paragraph.index), styleName: paragraph.name || '' };
    block.paragraphMarkStyle = styles.publicCharacter(styles.character(fcAt(last), paragraph.character, !!tocContext[last]).properties, '中');
    for (const key of paragraphKeys) if (properties[key] !== undefined) block[key] = properties[key];
    if (paragraph.sti >= 19 && paragraph.sti <= 27 || /^目录\s*\d|^toc\s*\d/i.test(paragraph.name || '')) {
      block.isToc = true; block.tocLevel = paragraph.sti >= 19 && paragraph.sti <= 27 ? paragraph.sti - 18 : Number(paragraph.name.match(/\d+/)?.[0] || 1);
    }
    if (properties._listId) { block.listId = properties._listId; block.listLevel = properties._listLevel || 0; }
    if (list) block.numbering = { text: list.text, suffix: list.suffix, align: list.align, style: styles.publicCharacter(list.character, list.text) };
    const sectionIndex = styles.sections.findIndex(section => section.start <= start && section.end > start);
    if (sectionIndex >= 0) block.sectionIndex = sectionIndex;
    if (runs.length || text[cp] === '\r' && visible[cp]) blocks.push(block);
    start = cp + 1;
  }
  if (blocks.length > 1 && !blocks[blocks.length - 1].runs.length) blocks.pop();
  if (pictures > imageReader.restored) warn(warnings, `DOC 中 ${pictures - imageReader.restored} 个行内图片位置尚未显示。`);
  if (cells) warn(warnings, 'DOC 表格目前按正文顺序读取，复杂行列和合并尚未还原。');
  const page = styles.sections[0]?.page || { width: 816, height: 1056, marginTop: 96, marginRight: 96, marginBottom: 96, marginLeft: 96 };
  return { kind: 'document', format: 'doc', name, warnings, blocks, page, defaultStyle: styles.defaultStyle, ...(styles.sections.length ? { sections: styles.sections.map(section => section.page) } : {}) };
}
