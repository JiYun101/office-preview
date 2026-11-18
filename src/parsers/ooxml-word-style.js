import { attr, children, elements, first, numberAttr, onOff, relatedPath, rgb, warn, xmlPart } from './ooxml-common.js';

const px = value => value / 15;
const alignment = value => ({ both: 'justify', distribute: 'justify', start: 'left', end: 'right' }[value] || value);

export function mergeWordParagraph(base = {}, own = {}) {
  const result = { ...base, ...own };
  if (base.numbering || own.numbering) result.numbering = { ...base.numbering, ...own.numbering };
  if (own.lineHeight !== undefined) delete result.lineHeightRatio;
  if (own.lineHeightRatio !== undefined) { delete result.lineHeight; delete result.lineHeightRule; }
  if (base.tabs || own.tabs) {
    const tabs = new Map((base.tabs || []).map(tab => [tab.position, { ...tab }]));
    for (const tab of own.tabs || []) {
      if (tab.align === 'clear') tabs.delete(tab.position); else tabs.set(tab.position, { ...tab });
    }
    result.tabs = [...tabs.values()].sort((a, b) => a.position - b.position);
  }
  return result;
}

export function createWordStyles(parts, rels, warnings) {
  const stylePath = relatedPath(rels, 'styles');
  const doc = stylePath ? xmlPart(parts, stylePath) : null;
  const settingsPath = relatedPath(rels, 'settings');
  const settings = settingsPath ? xmlPart(parts, settingsPath) : null;
  const themePath = relatedPath(rels, 'theme');
  const theme = themePath ? xmlPart(parts, themePath) : null;
  const scheme = elements(theme, 'fontScheme')[0];
  const language = attr(elements(settings, 'themeFontLang')[0], 'eastAsia');
  const script = /^ja/i.test(language) ? 'Jpan' : /^ko/i.test(language) ? 'Hang' : /(?:TW|HK|MO)/i.test(language) ? 'Hant' : 'Hans';
  const themeFonts = {};
  for (const kind of ['major', 'minor']) {
    const entry = first(scheme, `${kind}Font`);
    const eastAsia = attr(first(entry, 'ea'), 'typeface') || attr(children(entry, 'font').find(node => attr(node, 'script') === script), 'typeface');
    themeFonts[`${kind}HAnsi`] = attr(first(entry, 'latin'), 'typeface');
    themeFonts[`${kind}Ascii`] = themeFonts[`${kind}HAnsi`];
    themeFonts[`${kind}EastAsia`] = eastAsia;
    themeFonts[`${kind}Bidi`] = attr(first(entry, 'cs'), 'typeface');
  }
  const colorScheme = elements(theme, 'clrScheme')[0], colors = {};
  for (const entry of children(colorScheme)) {
    const value = children(entry)[0]; colors[entry.localName] = rgb(attr(value, 'val')) || rgb(attr(value, 'lastClr'));
  }
  const styleNodes = new Map(), resolved = new Map();
  let defaultParagraphId = '', defaultCharacterId = '';
  for (const node of children(doc?.documentElement, 'style')) {
    const id = attr(node, 'styleId');
    if (!id || styleNodes.has(id)) throw new Error('Word 样式 ID 缺失或重复。');
    styleNodes.set(id, node);
    if (onOff(node) && ['1', 'true', 'on'].includes(attr(node, 'default'))) {
      if (attr(node, 'type') === 'paragraph') defaultParagraphId = id;
      if (attr(node, 'type') === 'character') defaultCharacterId = id;
    }
  }

  function run(properties) {
    const style = {};
    for (const [tag, key] of [['b', 'bold'], ['i', 'italic'], ['strike', 'strike'], ['vanish', 'hidden']]) {
      const value = first(properties, tag); if (value) style[key] = onOff(value);
    }
    const underline = first(properties, 'u');
    if (underline) style.underline = !['none', '0', 'false'].includes(attr(underline, 'val'));
    const color = first(properties, 'color');
    if (color) {
      const themeName = { text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2' }[attr(color, 'themeColor')] || attr(color, 'themeColor');
      const value = rgb(attr(color, 'val')) || colors[themeName];
      if (value) style.color = value;
    }
    const size = first(properties, 'sz');
    if (size) { const value = numberAttr(size, 'val'); if (value > 0 && value <= 8000) style.fontSize = value * 2 / 3; }
    const fonts = first(properties, 'rFonts');
    if (fonts) {
      const latin = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || themeFonts[attr(fonts, 'asciiTheme')] || themeFonts[attr(fonts, 'hAnsiTheme')];
      const eastAsia = attr(fonts, 'eastAsia') || themeFonts[attr(fonts, 'eastAsiaTheme')];
      if (latin) style.fontFamilyLatin = latin;
      if (eastAsia) style.fontFamilyEastAsia = eastAsia;
      if (eastAsia || latin) style.fontFamily = eastAsia || latin;
    }
    const spacing = first(properties, 'spacing'); if (spacing) style.letterSpacing = px(numberAttr(spacing, 'val'));
    const position = first(properties, 'position'); if (position) style.baselineShift = numberAttr(position, 'val') * 2 / 3;
    const highlight = first(properties, 'highlight');
    const background = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', lightGray: '#C0C0C0', darkGray: '#808080', black: '#000000', white: '#FFFFFF' }[attr(highlight, 'val')] || rgb(attr(first(properties, 'shd'), 'fill'));
    if (background) style.background = background;
    const vertical = attr(first(properties, 'vertAlign'), 'val'); if (vertical) style.verticalAlign = vertical;
    return style;
  }

  function paragraph(properties) {
    const result = {};
    const justify = first(properties, 'jc'); if (justify) result.align = alignment(attr(justify, 'val'));
    const indent = first(properties, 'ind');
    if (indent) {
      const left = attr(indent, 'start') || attr(indent, 'left'), right = attr(indent, 'end') || attr(indent, 'right');
      if (left !== '') result.indentLeft = px(Number(left));
      if (right !== '') result.indentRight = px(Number(right));
      if (attr(indent, 'firstLine') !== '') result.firstLineIndent = px(numberAttr(indent, 'firstLine'));
      if (attr(indent, 'hanging') !== '') result.firstLineIndent = -px(numberAttr(indent, 'hanging'));
      for (const [name, target] of [['leftChars', 'indentLeftChars'], ['rightChars', 'indentRightChars'], ['firstLineChars', 'firstLineChars'], ['hangingChars', 'hangingChars']]) if (attr(indent, name) !== '') result[target] = numberAttr(indent, name) / 100;
    }
    const spacing = first(properties, 'spacing');
    if (spacing) {
      for (const [name, target] of [['before', 'spaceBefore'], ['after', 'spaceAfter']]) if (attr(spacing, name) !== '') result[target] = px(numberAttr(spacing, name));
      for (const [name, target] of [['beforeLines', 'spaceBeforeLines'], ['afterLines', 'spaceAfterLines']]) if (attr(spacing, name) !== '') result[target] = numberAttr(spacing, name) / 100;
      if (attr(spacing, 'line') !== '') {
        const value = numberAttr(spacing, 'line');
        if (attr(spacing, 'lineRule') === 'exact' || attr(spacing, 'lineRule') === 'atLeast') { result.lineHeight = px(value); result.lineHeightRule = attr(spacing, 'lineRule'); }
        else result.lineHeightRatio = value / 240;
      }
    }
    for (const [tag, key] of [['pageBreakBefore', 'pageBreakBefore'], ['keepNext', 'keepNext'], ['keepLines', 'keepLines'], ['snapToGrid', 'snapToGrid'], ['widowControl', 'widowControl'], ['contextualSpacing', 'contextualSpacing']]) {
      const value = first(properties, tag); if (value) result[key] = onOff(value);
    }
    const outline = first(properties, 'outlineLvl'); if (outline) result.outlineLevel = numberAttr(outline, 'val', 9);
    const tabs = first(properties, 'tabs');
    if (tabs) result.tabs = children(tabs, 'tab').map(tab => ({ position: px(numberAttr(tab, 'pos')), align: ({ start: 'left', end: 'right', num: 'decimal' }[attr(tab, 'val')]) || attr(tab, 'val') || 'left', leader: attr(tab, 'leader') || 'none' }));
    const numbering = first(properties, 'numPr');
    if (numbering) {
      result.numbering = {};
      if (first(numbering, 'numId')) result.numbering.numId = attr(first(numbering, 'numId'), 'val');
      if (first(numbering, 'ilvl')) result.numbering.level = numberAttr(first(numbering, 'ilvl'), 'val');
    }
    return result;
  }

  function resolve(id, visiting = new Set()) {
    if (!id || !styleNodes.has(id)) return { run: {}, paragraph: {}, name: '', type: '' };
    if (resolved.has(id)) return resolved.get(id);
    if (visiting.has(id) || visiting.size >= 128) throw new Error('Word 样式继承包含循环或层级过深。');
    const chain = new Set(visiting); chain.add(id);
    const node = styleNodes.get(id), parent = resolve(attr(first(node, 'basedOn'), 'val'), chain);
    const style = { run: { ...parent.run, ...run(first(node, 'rPr')) }, paragraph: mergeWordParagraph(parent.paragraph, paragraph(first(node, 'pPr'))), name: attr(first(node, 'name'), 'val'), type: attr(node, 'type') };
    resolved.set(id, style); return style;
  }

  const defaults = first(doc?.documentElement, 'docDefaults');
  const defaultRun = { fontSize: 14, color: '#000000', ...run(first(first(defaults, 'rPrDefault'), 'rPr')) };
  const defaultParagraph = { snapToGrid: true, ...paragraph(first(first(defaults, 'pPrDefault'), 'pPr')) };
  const defaultTabStop = numberAttr(elements(settings, 'defaultTabStop')[0], 'val', 720) / 15;
  const evenAndOddHeaders = onOff(first(settings?.documentElement, 'evenAndOddHeaders'));
  return { run, paragraph, resolve, defaultRun, defaultParagraph, defaultParagraphId, defaultCharacterId, defaultTabStop, evenAndOddHeaders };
}
