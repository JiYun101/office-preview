import { EMU_PER_PIXEL, TWIPS_PER_PIXEL, elements, children, first, attr, numberAttr, warn, xmlPart, relationships, imageFromRelationship } from './ooxml-common.js';
import { createWordStyles, mergeWordParagraph } from './ooxml-word-style.js';
import { createWordNumbering } from './ooxml-word-numbering.js';
import { parseWordTable } from './ooxml-word-table.js';

export function parseWord(parts, path, warnings) {
  const document = xmlPart(parts, path), body = first(document.documentElement, 'body');
  if (!body) throw new Error('DOCX 缺少文档正文。');
  const rels = relationships(parts, path, warnings);
  const styles = createWordStyles(parts, rels, warnings), numbering = createWordNumbering(parts, rels, styles, warnings);
  warn(warnings, 'DOCX 已读取样式、制表位和保存的分页标记；字体度量、浮动环绕与复杂分页仍可能不同，尚不保证精确分页。');
  const fields = [];
  let sectionIndex = 0;

  function runStyle(base, own) {
    const result = { ...base, ...own };
    result.fontFamily = result.fontFamilyEastAsia || result.fontFamilyLatin || result.fontFamily;
    return result;
  }
  function fieldVisible() { return fields.every(field => field.result); }
  function pageSettings(section) {
    const size = first(section, 'pgSz'), margin = first(section, 'pgMar'), grid = first(section, 'docGrid');
    const page = { width: numberAttr(size, 'w', 11906) / TWIPS_PER_PIXEL, height: numberAttr(size, 'h', 16838) / TWIPS_PER_PIXEL };
    for (const side of ['top', 'right', 'bottom', 'left']) page[`margin${side[0].toUpperCase()}${side.slice(1)}`] = numberAttr(margin, side, 1440) / TWIPS_PER_PIXEL;
    page.headerDistance = numberAttr(margin, 'header', 720) / TWIPS_PER_PIXEL;
    page.footerDistance = numberAttr(margin, 'footer', 720) / TWIPS_PER_PIXEL;
    if (['lines', 'linesAndChars'].includes(attr(grid, 'type')) && numberAttr(grid, 'linePitch') > 0) {
      page.lineGrid = numberAttr(grid, 'linePitch') / TWIPS_PER_PIXEL; page.lineGridType = attr(grid, 'type');
    }
    const count = numberAttr(first(section, 'cols'), 'num', 1);
    if (count > 1) warn(warnings, 'Word 多栏正文尚未按原始栏位排版。');
    return page;
  }
  const sectionNodes = elements(body, 'sectPr');
  const sections = (sectionNodes.length ? sectionNodes : [null]).map(pageSettings);

  function picture(node) {
    const blip = elements(node, 'blip')[0], legacy = elements(node, 'imagedata')[0];
    const id = attr(blip, 'embed') || attr(legacy, 'id');
    const image = id || blip || legacy ? imageFromRelationship(parts, rels, id, warnings) : null;
    if (!image) return null;
    const extent = elements(node, 'extent')[0];
    const width = numberAttr(extent, 'cx') / EMU_PER_PIXEL, height = numberAttr(extent, 'cy') / EMU_PER_PIXEL;
    if (width > 0) image.width = width;
    if (height > 0) image.height = height;
    if (!extent && legacy) {
      const style = attr(elements(node, 'shape')[0], 'style');
      for (const key of ['width', 'height']) {
        const match = new RegExp(`(?:^|;)\\s*${key}\\s*:\\s*([0-9.]+)(pt|px|in|cm|mm)`).exec(style);
        if (match) image[key] = Number(match[1]) * ({ pt: 4 / 3, px: 1, in: 96, cm: 96 / 2.54, mm: 96 / 25.4 }[match[2]]);
      }
    }
    const anchor = elements(node, 'anchor')[0];
    if (anchor) {
      const horizontal = first(anchor, 'positionH'), vertical = first(anchor, 'positionV');
      image.anchor = {
        x: Number(first(horizontal, 'posOffset')?.textContent || 0) / EMU_PER_PIXEL,
        y: Number(first(vertical, 'posOffset')?.textContent || 0) / EMU_PER_PIXEL,
        relativeFromH: attr(horizontal, 'relativeFrom') || 'column', relativeFromV: attr(vertical, 'relativeFrom') || 'paragraph',
        alignH: first(horizontal, 'align')?.textContent || '', alignV: first(vertical, 'align')?.textContent || '',
        behindText: ['1', 'true'].includes(attr(anchor, 'behindDoc')),
        wrap: children(anchor).find(child => /^wrap/.test(child.localName))?.localName || 'wrapNone',
      };
      warn(warnings, '浮动图片已保留位置参考，复杂文字环绕尚未完整还原。');
    }
    return image;
  }

  function paragraph(node, topLevel) {
    const properties = first(node, 'pPr'), id = attr(first(properties, 'pStyle'), 'val') || styles.defaultParagraphId;
    const definition = styles.resolve(id), direct = styles.paragraph(properties);
    const inherited = mergeWordParagraph(styles.defaultParagraph, definition.paragraph);
    const initial = mergeWordParagraph(inherited, direct), number = numbering.next(initial.numbering);
    const layout = mergeWordParagraph(mergeWordParagraph(styles.defaultParagraph, number?.paragraph), definition.paragraph);
    const finalLayout = mergeWordParagraph(layout, direct);
    const defaultRun = runStyle(styles.defaultRun, definition.run);
    const styleName = definition.name || id;
    const toc = /^(?:toc|目录)\s*([1-9])$/i.exec(styleName) || /^(?:TOC)([1-9])$/i.exec(id);
    const runs = [];
    const add = (value, style) => {
      if (!fieldVisible() || style.hidden) return;
      const { hidden, ...visible } = style;
      if (value.text && !value.type && !value.image && style.fontFamilyLatin && style.fontFamilyEastAsia && style.fontFamilyLatin !== style.fontFamilyEastAsia) {
        // Word selects the Western and East Asian font slots per character.
        let chunk = '', eastAsia;
        const flush = () => { if (chunk) runs.push({ ...visible, ...value, text: chunk, fontFamily: eastAsia ? style.fontFamilyEastAsia : style.fontFamilyLatin }); chunk = ''; };
        for (const character of value.text) {
          const isEastAsia = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\u3000-\u303f\uff00-\uffef]/u.test(character);
          if (eastAsia !== undefined && isEastAsia !== eastAsia) flush();
          eastAsia = isEastAsia; chunk += character;
        }
        flush();
      } else runs.push({ ...visible, ...value });
    };
    function visit(current, style = defaultRun) {
      for (const child of children(current)) {
        const tag = child.localName;
        if (['pPr', 'rPr', 'del', 'moveFrom'].includes(tag)) continue;
        if (tag === 'r') {
          const properties = first(child, 'rPr'), characterId = attr(first(properties, 'rStyle'), 'val') || styles.defaultCharacterId;
          const character = styles.resolve(characterId);
          let characterRun = character.run;
          const tocField = fields.some(field => field.result && /^\s*TOC(?:\s|$)/i.test(field.instruction));
          if (toc && tocField && /^(?:Hyperlink|FollowedHyperlink|超链接|已访问的超链接)$/i.test(character.name)) {
            // Word's paginated TOC uses its paragraph appearance instead of
            // the hyperlink style's blue/underlined web appearance. Explicit
            // rPr below remains authoritative.
            const { color, underline, ...printStyle } = characterRun; characterRun = printStyle;
          }
          visit(child, runStyle(runStyle(style, characterRun), styles.run(properties)));
        } else if (tag === 'fldChar') {
          const kind = attr(child, 'fldCharType');
          if (kind === 'begin') { if (fields.length >= 128) throw new Error('Word 域嵌套层数超过当前上限。'); fields.push({ result: false, instruction: '' }); }
          else if (kind === 'separate' && fields.length) fields[fields.length - 1].result = true;
          else if (kind === 'end' && fields.length) fields.pop();
        } else if (tag === 'instrText') { if (fields.length) fields[fields.length - 1].instruction += child.textContent; }
        else if (tag === 'fldSimple') {
          if (fields.length >= 128) throw new Error('Word 域嵌套层数超过当前上限。');
          const depth = fields.length; fields.push({ result: true, instruction: attr(child, 'instr') });
          visit(child, style); fields.length = depth;
        }
        else if (tag === 't') add({ text: child.textContent }, style);
        else if (tag === 'tab' || tag === 'ptab') add({ type: 'tab', text: '\t' }, style);
        else if (tag === 'lastRenderedPageBreak') add({ type: 'pageBreak', text: '\f', source: 'rendered', cached: true }, style);
        else if (tag === 'br' || tag === 'cr') {
          if (attr(child, 'type') === 'page') add({ type: 'pageBreak', text: '\f', source: 'explicit', cached: false }, style);
          else if (attr(child, 'type') === 'column') { warn(warnings, 'Word 分栏符目前按换行显示。'); add({ text: '\n' }, style); }
          else add({ text: '\n' }, style);
        } else if (tag === 'noBreakHyphen') add({ text: '\u2011' }, style);
        else if (tag === 'softHyphen') add({ text: '\u00ad' }, style);
        else if (tag === 'drawing' || tag === 'pict') { const image = picture(child); if (image) add({ text: '', image }, style); }
        else if (tag === 'oMath' || tag === 'oMathPara') warn(warnings, '数学公式尚未渲染。');
        else if (tag === 'AlternateContent') {
          const choice = children(child, 'Choice').find(value => elements(value, 'drawing').length || elements(value, 'pict').length) || first(child, 'Fallback');
          if (choice) visit(choice, style);
        } else if (tag !== 'txbxContent') visit(child, style);
      }
    }
    visit(node);
    const { numbering: unusedNumber, ...paragraphLayout } = finalLayout;
    const result = { type: 'paragraph', runs, style: defaultRun, styleId: id, styleName, ...paragraphLayout, sectionIndex };
    const mark = first(properties, 'rPr'); if (mark) result.paragraphMarkStyle = runStyle(defaultRun, styles.run(mark));
    if (toc) { result.isToc = true; result.tocLevel = Number(toc[1]); }
    if (/^(?:toc heading|目录标题)$/i.test(styleName)) result.isTocHeading = true;
    if (number) {
      for (const key of ['numberingLabel', 'numberingSuffix', 'numberingLevel', 'numberingStyle']) result[key] = number[key];
    }
    if (topLevel && first(properties, 'sectPr')) {
      result.sectionBreakAfter = attr(first(first(properties, 'sectPr'), 'type'), 'val') || 'nextPage';
      sectionIndex = Math.min(sectionIndex + 1, sections.length - 1);
    }
    return result;
  }

  function blocks(container, topLevel = false) {
    const result = [];
    for (const node of children(container)) {
      if (node.localName === 'p') result.push(paragraph(node, topLevel));
      else if (node.localName === 'tbl') {
        result.push(parseWordTable(node, { blocks, sectionIndex, warnings }));
      } else if (node.localName === 'sdt') result.push(...blocks(first(node, 'sdtContent'), topLevel));
      else if (['customXml', 'ins', 'moveTo'].includes(node.localName)) result.push(...blocks(node, topLevel));
      else if (node.localName === 'altChunk') warn(warnings, '嵌入的 HTML/外部文档内容尚未解析。');
    }
    return result;
  }

  const content = blocks(body, true);
  if (elements(body, 'headerReference').length || elements(body, 'footerReference').length) warn(warnings, '页眉和页脚尚未显示。');
  if (elements(body, 'del').length || elements(body, 'ins').length) warn(warnings, '修订内容按最终版本读取：保留插入、忽略删除。');
  if (elements(body, 'chart').length || elements(body, 'graphicData').some(node => /chart|diagram/.test(attr(node, 'uri')))) warn(warnings, '图表和 SmartArt 尚未渲染。');
  if (elements(body, 'footnoteReference').length || elements(body, 'endnoteReference').length) warn(warnings, '脚注、尾注尚未显示。');
  if (elements(body, 'txbxContent').length) warn(warnings, '文本框内容尚未显示。');
  if (elements(body, 'sym').length) warn(warnings, '特殊符号字体映射尚未实现。');
  if (content.some(block => block.isToc)) warn(warnings, '目录与域页码显示文件保存的结果，不在预览中重新计算。');
  return { kind: 'document', blocks: content, page: sections[0], sections, defaultStyle: styles.defaultRun, defaultTabStop: styles.defaultTabStop, hasSavedPageBreaks: elements(body, 'lastRenderedPageBreak').length > 0 };
}
