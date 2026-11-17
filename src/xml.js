// Small, read-only XML 1.0 tree for Office package parts. Works in workers and Node.
// This reader intentionally has no DTD, entity declarations, XPath, or mutation API.
const XML_NS = 'http://www.w3.org/XML/1998/namespace';
const XMLNS_NS = 'http://www.w3.org/2000/xmlns/';
const MAX_NODES = 200000;
const MAX_DEPTH = 128;
const NAME = /[:A-Z_a-z\u00c0-\u00d6\u00d8-\u00f6\u00f8-\u02ff\u0370-\u037d\u037f-\u1fff\u200c-\u200d\u2070-\u218f\u2c00-\u2fef\u3001-\ud7ff\uf900-\ufdcf\ufdf0-\ufffd\u{10000}-\u{effff}][:A-Z_a-z0-9.\-\u00b7\u00c0-\u00d6\u00d8-\u00f6\u00f8-\u02ff\u0300-\u037d\u037f-\u1fff\u200c-\u200d\u203f-\u2040\u2070-\u218f\u2c00-\u2fef\u3001-\ud7ff\uf900-\ufdcf\ufdf0-\ufffd\u{10000}-\u{effff}]*/uy;
const INVALID_CHARACTER = /[^\u0009\u000a\u000d\u0020-\ud7ff\ue000-\ufffd\u{10000}-\u{10ffff}]/u;
const ENTITIES = { amp: '&', lt: '<', gt: '>', apos: "'", quot: '"' };

function xmlCharacter(code) {
  return code === 9 || code === 10 || code === 13 || code >= 0x20 && code <= 0xd7ff || code >= 0xe000 && code <= 0xfffd || code >= 0x10000 && code <= 0x10ffff;
}

function decodeEntities(source) {
  const chunks = [];
  let cursor = 0;
  while (true) {
    const ampersand = source.indexOf('&', cursor);
    if (ampersand < 0) { chunks.push(source.slice(cursor)); break; }
    chunks.push(source.slice(cursor, ampersand));
    const end = source.indexOf(';', ampersand + 1);
    if (end < 0) throw new Error('文档 XML 实体没有结束。');
    const reference = source.slice(ampersand + 1, end);
    if (Object.hasOwn(ENTITIES, reference)) chunks.push(ENTITIES[reference]);
    else if (/^#(?:[0-9]+|x[0-9a-fA-F]+)$/.test(reference)) {
      const code = reference[1] === 'x' ? Number.parseInt(reference.slice(2), 16) : Number(reference.slice(1));
      if (!xmlCharacter(code)) throw new Error('文档 XML 字符引用无效。');
      chunks.push(String.fromCodePoint(code));
    } else throw new Error('文档 XML 包含未知或自定义实体。');
    cursor = end + 1;
  }
  return chunks.join('');
}

class XmlNode {
  constructor(name = '', namespaceURI = '', attributes = []) {
    this.name = name;
    this.localName = name.split(':').pop();
    this.namespaceURI = namespaceURI;
    this.attributes = attributes;
    this.children = [];
    this.content = [];
  }

  get textContent() {
    const chunks = [];
    const pending = [...this.content].reverse();
    while (pending.length) {
      const item = pending.pop();
      if (typeof item === 'string') chunks.push(item);
      else for (let index = item.content.length - 1; index >= 0; index -= 1) pending.push(item.content[index]);
    }
    return chunks.join('');
  }

  getElementsByTagNameNS(namespaceURI, localName) {
    const found = [];
    const pending = [...this.children].reverse();
    while (pending.length) {
      const node = pending.pop();
      if ((localName === '*' || node.localName === localName) && (namespaceURI === '*' || node.namespaceURI === namespaceURI || !namespaceURI && !node.namespaceURI)) found.push(node);
      for (let index = node.children.length - 1; index >= 0; index -= 1) pending.push(node.children[index]);
    }
    return found;
  }
}

function decodeSource(input) {
  if (typeof input === 'string') return input;
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  let encoding = 'utf-8';
  if (bytes[0] === 0xff && bytes[1] === 0xfe || bytes[0] === 0x3c && bytes[1] === 0) encoding = 'utf-16le';
  if (bytes[0] === 0xfe && bytes[1] === 0xff || bytes[0] === 0 && bytes[1] === 0x3c) encoding = 'utf-16be';
  return new TextDecoder(encoding, { fatal: true }).decode(bytes);
}

export function parseXml(input) {
  const source = decodeSource(input).replace(/^\ufeff/, '').replace(/\r\n?/g, '\n');
  if (INVALID_CHARACTER.test(source)) throw new Error('文档 XML 包含无效字符。');
  const document = new XmlNode();
  document.documentElement = null;
  const stack = [];
  let cursor = 0;
  let nodeCount = 0;
  const fail = message => { throw new Error(`文档 XML 已损坏：${message}（位置 ${cursor}）。`); };
  function count() {
    nodeCount += 1;
    if (nodeCount > MAX_NODES) throw new Error(`文档 XML 超过 ${MAX_NODES} 个节点/属性的解析上限。`);
  }
  function skipSpace() {
    const start = cursor;
    while (cursor < source.length && /[ \t\n]/.test(source[cursor])) cursor += 1;
    return cursor > start;
  }
  function readName() {
    NAME.lastIndex = cursor;
    const match = NAME.exec(source);
    if (!match) fail('名称无效');
    cursor = NAME.lastIndex;
    const name = match[0];
    const qualified = name.split(':');
    if (qualified.length > 2 || qualified.some(part => !part || /^[0-9.\-\u00b7\u0300-\u036f\u203f-\u2040]/u.test(part))) fail('限定名称无效');
    return name;
  }
  function appendText(text, cdata = false) {
    if (!stack.length) {
      if (/[^ \t\n]/.test(text)) fail('根元素以外存在文本');
      return;
    }
    if (!text) return;
    if (!cdata && text.includes(']]>')) fail('普通文本中出现 CDATA 结束标记');
    count();
    stack[stack.length - 1].node.content.push(cdata ? text : decodeEntities(text));
  }

  while (cursor < source.length) {
    if (source[cursor] !== '<') {
      const end = source.indexOf('<', cursor);
      const textEnd = end < 0 ? source.length : end;
      appendText(source.slice(cursor, textEnd));
      cursor = textEnd;
      continue;
    }
    if (source.startsWith('<!--', cursor)) {
      const end = source.indexOf('-->', cursor + 4);
      if (end < 0) fail('注释没有结束');
      const text = source.slice(cursor + 4, end);
      if (text.includes('--') || text.endsWith('-')) fail('注释中包含双连字符');
      count();
      cursor = end + 3;
      continue;
    }
    if (source.startsWith('<![CDATA[', cursor)) {
      if (!stack.length) fail('CDATA 不在根元素内');
      const end = source.indexOf(']]>', cursor + 9);
      if (end < 0) fail('CDATA 没有结束');
      appendText(source.slice(cursor + 9, end), true);
      cursor = end + 3;
      continue;
    }
    if (source.startsWith('<?', cursor)) {
      const start = cursor;
      cursor += 2;
      const target = readName();
      const end = source.indexOf('?>', cursor);
      if (end < 0) fail('处理指令没有结束');
      if (cursor !== end && !/[ \t\n]/.test(source[cursor])) fail('处理指令格式无效');
      if (target.toLowerCase() === 'xml') {
        const declaration = source.slice(start + 2, end);
        if (start !== 0 || !/^xml[ \t\n]+version[ \t\n]*=[ \t\n]*(?:"1\.0"|'1\.0')(?:[ \t\n]+encoding[ \t\n]*=[ \t\n]*(?:"[A-Za-z][A-Za-z0-9._-]*"|'[A-Za-z][A-Za-z0-9._-]*'))?(?:[ \t\n]+standalone[ \t\n]*=[ \t\n]*(?:"(?:yes|no)"|'(?:yes|no)'))?[ \t\n]*$/.test(declaration)) fail('XML 声明无效或版本尚不支持');
      }
      count();
      cursor = end + 2;
      continue;
    }
    if (source.startsWith('<!', cursor)) throw new Error('文档 XML 不允许 DTD、自定义实体或未知声明。');
    if (source.startsWith('</', cursor)) {
      cursor += 2;
      const name = readName();
      skipSpace();
      if (source[cursor] !== '>') fail('结束标签无效');
      const frame = stack.pop();
      if (!frame || frame.node.name !== name) fail('标签未正确闭合');
      cursor += 1;
      continue;
    }

    cursor += 1;
    const name = readName();
    const rawAttributes = [];
    const attributeNames = new Set();
    let selfClosing = false;
    while (cursor < source.length) {
      const separated = skipSpace();
      if (source[cursor] === '>') { cursor += 1; break; }
      if (source.startsWith('/>', cursor)) { cursor += 2; selfClosing = true; break; }
      if (!separated) fail('属性前缺少空格，或开始标签没有结束');
      const attributeName = readName();
      if (attributeNames.has(attributeName)) fail('属性重复');
      attributeNames.add(attributeName);
      skipSpace();
      if (source[cursor] !== '=') fail('属性缺少等号');
      cursor += 1;
      skipSpace();
      const quote = source[cursor];
      if (quote !== '"' && quote !== "'") fail('属性必须有引号');
      const end = source.indexOf(quote, cursor + 1);
      if (end < 0) fail('属性引号没有结束');
      const raw = source.slice(cursor + 1, end);
      if (raw.includes('<')) fail('属性包含未转义的小于号');
      count();
      rawAttributes.push({ name: attributeName, value: decodeEntities(raw.replace(/[\t\n]/g, ' ')) });
      cursor = end + 1;
    }
    if (source[cursor - 1] !== '>') fail('开始标签没有结束');
    if (stack.length >= MAX_DEPTH) throw new Error(`文档 XML 嵌套超过 ${MAX_DEPTH} 层上限。`);
    let namespaces = stack.length ? stack[stack.length - 1].namespaces : new Map([['xml', XML_NS], ['', '']]);
    const declarations = rawAttributes.filter(item => item.name === 'xmlns' || item.name.startsWith('xmlns:'));
    if (declarations.length) {
      namespaces = new Map(namespaces);
      for (const item of declarations) {
        const prefix = item.name === 'xmlns' ? '' : item.name.slice(6);
        if (prefix === 'xmlns' || item.value === XMLNS_NS || prefix === 'xml' && item.value !== XML_NS || prefix !== 'xml' && item.value === XML_NS || prefix && !item.value) fail('命名空间声明无效');
        namespaces.set(prefix, item.value);
      }
    }
    const prefix = name.includes(':') ? name.split(':')[0] : '';
    if (prefix === 'xmlns' || !namespaces.has(prefix)) fail('元素前缀未声明');
    const expandedAttributes = new Set();
    const attributes = rawAttributes.map(item => {
      const parts = item.name.split(':');
      const attributePrefix = parts.length === 2 ? parts[0] : '';
      const namespaceURI = item.name === 'xmlns' || attributePrefix === 'xmlns' ? XMLNS_NS : attributePrefix ? namespaces.get(attributePrefix) : '';
      if (namespaceURI === undefined) fail('属性前缀未声明');
      const localName = parts[parts.length - 1];
      const key = `${namespaceURI}\u0000${localName}`;
      if (expandedAttributes.has(key)) fail('同一命名空间中的属性重复');
      expandedAttributes.add(key);
      return { ...item, localName, namespaceURI };
    });
    count();
    const node = new XmlNode(name, namespaces.get(prefix), attributes);
    const parent = stack.length ? stack[stack.length - 1].node : document;
    if (!stack.length) {
      if (document.documentElement) fail('包含多个根元素');
      document.documentElement = node;
    }
    parent.children.push(node);
    parent.content.push(node);
    if (!selfClosing) stack.push({ node, namespaces });
  }
  if (stack.length || !document.documentElement) fail('根元素缺失或标签未闭合');
  return document;
}

export function elements(node, name) { return node?.getElementsByTagNameNS('*', name) || []; }
export function children(node, name) { return (node?.children || []).filter(child => !name || child.localName === name); }
export function first(node, name) { return elements(node, name)[0] || null; }
export function attr(node, name) { return (node?.attributes || []).find(item => item.localName === name)?.value ?? null; }
