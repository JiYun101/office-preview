import { parseCsv, decodeText } from './parsers/csv.js';
import { parseOoxml } from './parsers/ooxml.js';
import { parseLegacy } from './parsers/legacy.js';

export const capabilities = Object.freeze([
  { formats: ['csv', 'tsv'], level: '基础数据预览', detail: '引号、换行、分隔符、UTF-8 / UTF-16 / GB18030；按文本显示，不执行公式。' },
  { formats: ['docx'], level: '基础纸页预览', detail: '样式继承、目录与制表位、行内/无环绕浮图、不等宽表格与合并、行高/边框/内边距、保存分页与自动排版；复杂环绕、页眉页脚和精确分页仍有限制。' },
  { formats: ['xlsx'], level: '基础数据预览', detail: '工作表、缓存值、共享字符串、合并格；尚无公式计算引擎、图表和完整数字格式。' },
  { formats: ['pptx'], level: '基础静态预览', detail: '幻灯片顺序、文字定位、部分图片和基础形状；尚无完整母版、动画及音视频。' },
  { formats: ['doc'], level: '97–2003 基础纸页预览', detail: '正文、字符与段落样式、目录制表位、自动编号、常见图片和页面设置；复杂表格、环绕及原生分页仍有限制。' },
  { formats: ['xls'], level: 'BIFF8 基础版式预览', detail: '单元格、字体、对齐、边框、填充、行高列宽与常见数字格式；图表和打印版式尚未还原。' },
  { formats: ['ppt'], level: '97–2003 基础静态版式', detail: '真实页面尺寸、定位文本、形状颜色和常见图片，基础母版；复杂图形、媒体和动画仍有限制。' },
  { formats: ['pdf'], level: '基础页面预览', detail: '自主解析页面、嵌入 TrueType/OpenType 字形、基础路径与填充/描边透明度；特殊字体、复杂透明合成和交互功能尚不完整。' },
  { formats: ['eml', 'msg'], level: '单封邮件预览', detail: '主题、收发件人、纯文本正文和附件清单；PST / OST 邮箱存档尚未实现。' }
]);

function validateMerges(model) {
  // Validate once in the worker, including ranges outside the first visible page.
  for (const sheet of model.sheets || []) {
    const ranges = [...(sheet.merges || [])].sort((a, b) => a.startRow - b.startRow || a.startCol - b.startCol);
    let active = [];
    for (const range of ranges) {
      const { startRow, endRow, startCol, endCol } = range;
      if (![startRow, endRow, startCol, endCol].every(Number.isInteger) || startRow < 0 || startCol < 0 || endRow < startRow || endCol < startCol || endRow >= 50000 || endCol >= 1024) throw new Error('工作表合并区域无效或超过首版行列限制。');
      active = active.filter(item => item.endRow >= startRow);
      if (active.some(item => item.startCol <= endCol && item.endCol >= startCol)) throw new Error('工作表包含重叠的合并区域，无法正确预览。');
      active.push(range);
    }
  }
  return model;
}

export async function parseDocument(input, options = {}) {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const name = options.name || '未命名';
  if (bytes.length > 32 * 1024 * 1024) throw new Error('文件超过首版 32MB 限制。');
  const format = (options.format || name.split('.').at(-1)).toLowerCase();
  if (['csv', 'tsv'].includes(format)) return parseCsv(bytes, format, name, options);
  if (format === 'txt') {
    const warnings = [];
    const text = decodeText(bytes, options.encoding, warnings);
    return { kind: 'document', format, name, warnings, blocks: text.split(/\r\n|\r|\n/).map(text => ({ type: 'paragraph', runs: [{ text }] })) };
  }
  if (['docx', 'xlsx', 'pptx'].includes(format)) {
    if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('该文件可能是密码加密的 Office 文档；首版尚不支持解密。');
    return validateMerges(await parseOoxml(bytes, format, name));
  }
  if (['doc', 'xls', 'ppt'].includes(format)) {
    const model = parseLegacy(bytes, format, name);
    if (format === 'ppt') {
      const { enrichPptRoundtrip } = await import('./parsers/legacy-ppt-art.js');
      await enrichPptRoundtrip(model);
    }
    return validateMerges(model);
  }
  if (format === 'pdf') {
    const { parsePdf } = await import('./parsers/pdf.js');
    return parsePdf(bytes, name);
  }
  if (['eml', 'msg'].includes(format)) {
    const { parseEmail } = await import('./parsers/email.js');
    return parseEmail(bytes, format, name);
  }
  if (['pst', 'ost'].includes(format)) throw new Error('PST / OST 邮箱存档读取尚未实现；当前可预览单封 EML / MSG 邮件。');
  throw new Error(`暂未实现 .${format} 格式。当前可读取 CSV、TSV、TXT、DOCX、XLSX、PPTX、PDF、EML、MSG 及 DOC/XLS/PPT 97–2003 基础子集。`);
}
