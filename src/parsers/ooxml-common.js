// Package and XML helpers for our own small OOXML readers. No Office engine is used.
import { parseXml } from '../xml.js';
export const EMU_PER_PIXEL = 9525;
export const TWIPS_PER_PIXEL = 15;

export function elements(node, name) {
  return node ? Array.from(node.getElementsByTagNameNS('*', name)) : [];
}

export function children(node, name) {
  return Array.from(node?.children || []).filter(child => !name || child.localName === name);
}

export function first(node, name) {
  return children(node, name)[0] || null;
}

export function attr(node, name) {
  return Array.from(node?.attributes || []).find(value => value.localName === name)?.value ?? '';
}

export function relationshipId(node) {
  // p:sldId has both an unqualified numeric id and r:id: their local names collide.
  return Array.from(node?.attributes || []).find(value => value.localName === 'id' && /relationships/.test(value.namespaceURI || ''))?.value || '';
}

export function numberAttr(node, name, fallback = 0) {
  const value = attr(node, name);
  return value !== '' && Number.isFinite(Number(value)) ? Number(value) : fallback;
}

export function warn(warnings, message) {
  if (!warnings.includes(message)) warnings.push(message);
}

export function xmlPart(parts, path, required = true) {
  const bytes = parts.get(path);
  if (!bytes) {
    if (required) throw new Error(`Office 文件缺少必要部件：${path}`);
    return null;
  }
  let encoding = 'utf-8';
  if (bytes[0] === 0xff && bytes[1] === 0xfe || bytes[0] === 0x3c && bytes[1] === 0) encoding = 'utf-16le';
  if (bytes[0] === 0xfe && bytes[1] === 0xff || bytes[0] === 0 && bytes[1] === 0x3c) encoding = 'utf-16be';
  const source = new TextDecoder(encoding, { fatal: true }).decode(bytes);
  try { return parseXml(source); }
  catch (error) { throw new Error(`${path}：${error.message}`); }
}

export function resolveTarget(source, target) {
  if (!target || /^(?:[a-z][a-z0-9+.-]*:|\/\/)/i.test(target)) return null;
  let decoded;
  try { decoded = decodeURIComponent(target); } catch { throw new Error('Office 关系路径编码无效。'); }
  if (/[\\\u0000?#]/.test(decoded)) throw new Error('Office 关系路径无效。');
  const segments = decoded.startsWith('/') ? [] : source.split('/').slice(0, -1);
  for (const segment of decoded.split('/')) {
    if (!segment || segment === '.') continue;
    if (segment === '..') {
      if (!segments.length) throw new Error('Office 关系路径超出文件包。');
      segments.pop();
    } else segments.push(segment);
  }
  return segments.join('/');
}

export function relationships(parts, source, warnings) {
  const slash = source.lastIndexOf('/');
  const path = source ? `${source.slice(0, slash + 1)}_rels/${source.slice(slash + 1)}.rels` : '_rels/.rels';
  const doc = xmlPart(parts, path, !source);
  const result = new Map();
  for (const rel of elements(doc, 'Relationship')) {
    const id = attr(rel, 'Id');
    if (!id || result.has(id)) throw new Error('Office 文件关系 ID 无效或重复。');
    const external = attr(rel, 'TargetMode').toLowerCase() === 'external';
    const target = external ? null : resolveTarget(source, attr(rel, 'Target'));
    if (!target) warn(warnings, '外部图片、链接和数据源未加载。');
    result.set(id, { type: attr(rel, 'Type'), target });
  }
  return result;
}

export function relatedPath(rels, type) {
  return Array.from(rels.values()).find(rel => rel.type.endsWith(`/${type}`))?.target || null;
}

export function imageFromRelationship(parts, rels, id, warnings) {
  const rel = rels.get(id);
  if (!rel?.target) {
    warn(warnings, '部分图片缺少内部关系或引用了外部资源，未显示。');
    return null;
  }
  const bytes = parts.get(rel.target);
  if (!bytes) throw new Error(`Office 图片部件丢失：${rel.target}`);
  const extension = rel.target.split('.').pop().toLowerCase();
  const mime = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp' }[extension];
  if (!mime) {
    warn(warnings, `尚未解码 ${extension.toUpperCase()} 图片；当前支持 PNG、JPEG、GIF、WebP、BMP。`);
    return null;
  }
  return { bytes, mime };
}

export function onOff(node) {
  return !!node && !['0', 'false', 'off'].includes(attr(node, 'val').toLowerCase());
}

export function rgb(value) {
  return /^(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value) ? `#${value.slice(-6)}` : undefined;
}

export function richText(node) {
  if (!node) return '';
  return children(node).map(child => child.localName === 't' ? child.textContent : child.localName === 'r' ? first(child, 't')?.textContent || '' : '').join('');
}

export function themeColors(parts, rels) {
  const path = relatedPath(rels, 'theme');
  const theme = path ? xmlPart(parts, path) : null;
  const scheme = elements(theme, 'clrScheme')[0];
  const colors = {};
  for (const child of children(scheme)) {
    const value = children(child)[0];
    colors[child.localName] = rgb(attr(value, 'val')) || rgb(attr(value, 'lastClr'));
  }
  return colors;
}

export function drawingColor(node, theme = {}) {
  const value = children(node).find(child => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(child.localName));
  if (!value) return undefined;
  if (value.localName === 'schemeClr') {
    const key = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2' }[attr(value, 'val')] || attr(value, 'val');
    return theme[key];
  }
  return rgb(attr(value, 'val')) || rgb(attr(value, 'lastClr'));
}
