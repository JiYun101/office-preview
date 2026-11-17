import { readCompound } from '../containers/cfb.js';
import { parseXls } from './legacy-xls.js';
import { parsePpt } from './legacy-ppt.js';
import { parseDoc } from './legacy-doc.js';

export function parseLegacy(bytes, format, name = '') {
  if (!(bytes instanceof Uint8Array)) bytes = new Uint8Array(bytes);
  const compound = readCompound(bytes);
  if (compound.get('EncryptedPackage') || compound.get('EncryptionInfo')) throw new Error('文档已加密，当前版本不支持解密。');
  if (format === 'doc') return parseDoc(compound, name);
  if (format === 'xls') return parseXls(compound, name);
  if (format === 'ppt') return parsePpt(compound, name);
  throw new Error(`不支持的旧版 Office 格式：${format}`);
}
