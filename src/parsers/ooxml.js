import { readZip } from '../containers/zip.js';
import { xmlPart, relationships, relatedPath } from './ooxml-common.js';
import { parseWord } from './ooxml-word.js';
import { parseWorkbook } from './ooxml-sheet.js';
import { parsePresentation } from './ooxml-slides.js';

export async function parseOoxml(bytes, format, name = '') {
  format = String(format).toLowerCase().replace(/^\./, '');
  const parts = await readZip(bytes);
  const warnings = [];
  const contentTypes = xmlPart(parts, '[Content_Types].xml');
  if (contentTypes.documentElement.localName !== 'Types') throw new Error('无效的 Office 内容类型清单。');
  const rels = relationships(parts, '', warnings);
  const main = relatedPath(rels, 'officeDocument');
  if (!main) throw new Error('文件中没有 Office 主文档关系。');
  const root = xmlPart(parts, main).documentElement.localName;
  const expected = /^(docx|docm|dotx|dotm)$/.test(format) ? 'document' : /^(xlsx|xlsm|xltx|xltm)$/.test(format) ? 'workbook' : /^(pptx|pptm|ppsx|ppsm|potx|potm)$/.test(format) ? 'presentation' : null;
  if (!expected || root !== expected) throw new Error('Office 文件内容与扩展名不一致，或格式尚不支持。');
  if (Array.from(parts.keys()).some(path => /vbaProject\.bin$/i.test(path))) warnings.push('宏代码不会执行。');
  const model = root === 'document' ? parseWord(parts, main, warnings) : root === 'workbook' ? parseWorkbook(parts, main, warnings) : parsePresentation(parts, main, warnings);
  return { ...model, format, name, warnings };
}
