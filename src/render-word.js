function node(tag, className, text) {
  const element = document.createElement(tag); element.className = className;
  if (text !== undefined) element.textContent = String(text);
  return element;
}

export function renderWord(model, viewport, content, auxiliary, tools = {}) {
  let disposed = false, current, sectionIndex = 0, pageCount = 0, savedFlow = false;
  const defaults = { width: 794, height: 1123, marginTop: 96, marginBottom: 96, marginLeft: 96, marginRight: 96, ...model.page };
  const pages = node('div', 'op-word-pages'); content.append(pages);
  const allPages = [], tabs = new Map(), metrics = new Map(), rowHeights = new Map();
  const canvas = document.createElement('canvas'), context = canvas.getContext('2d');
  const hasSaved = !!model.hasSavedPageBreaks || (model.blocks || []).some(block => block.runs?.some(run => run.cached && run.type === 'pageBreak'));
  const specFor = index => ({ ...defaults, ...(model.sections?.[index]?.page || model.sections?.[index] || {}) });
  const setColumns = () => {
    const width = defaults.width;
    const columns = viewport.clientWidth >= width * 2 + 80 ? 2 : 1;
    pages.style.width = `${width * columns + 24 * (columns - 1)}px`;
    pages.dataset.columns = String(columns);
    return width * columns + 24 * (columns - 1) + 56;
  };
  const pageScale = page => page.paper.getBoundingClientRect().width / page.spec.width || 1;

  function newPage() {
    if (current && !current.body.children.length) {
      if (current.sectionIndex === sectionIndex) return current;
      current.paper.remove(); allPages.pop(); pageCount--;
    }
    if (++pageCount > 1000) throw new Error('Word 文档分页超过当前 1000 页预览上限。');
    const previous = allPages.at(-1), section = model.sections?.[sectionIndex];
    const firstPage = previous?.sectionIndex !== sectionIndex;
    const number = (firstPage ? section?.pageNumberStart : undefined) ?? ((previous?.number ?? 0) + 1);
    const variant = firstPage && section?.titlePage ? 'first' : model.evenAndOddHeaders && number % 2 === 0 ? 'even' : 'default';
    const spec = specFor(sectionIndex), paper = node('article', 'op-paper op-word-page'), body = node('div', 'op-word-page-body');
    Object.assign(paper.style, { width: `${spec.width}px`, height: `${spec.height}px`, minHeight: `${spec.height}px`, padding: `${spec.marginTop}px ${spec.marginRight}px ${spec.marginBottom}px ${spec.marginLeft}px` });
    paper.dataset.page = String(pageCount); paper.setAttribute('aria-label', `第 ${pageCount} 页`);
    paper.append(body); pages.append(paper);
    current = { paper, body, spec, sectionIndex, number }; allPages.push(current);
    current.header = renderPageArea('header', section?.headers?.[variant], current);
    current.footer = renderPageArea('footer', section?.footers?.[variant], current);
    current.height = Math.max(1, spec.height - spec.marginTop - spec.marginBottom);
    body.style.minHeight = `${current.height}px`;
    paper.style.height = `${Math.max(spec.height, spec.marginTop + spec.marginBottom + current.height)}px`;
    return current;
  }
  function fontMetrics(element) {
    const style = getComputedStyle(element);
    const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    const eastAsian = /[\u2e80-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/u.test(element.textContent);
    const key = font + (eastAsian ? '/ea' : '/latin');
    if (metrics.has(key)) return metrics.get(key);
    context.font = font;
    const measure = context.measureText(eastAsian ? 'Hg中' : 'Hg'), size = parseFloat(style.fontSize) || 14;
    const ascent = measure.fontBoundingBoxAscent || size * 0.95, descent = measure.fontBoundingBoxDescent || size * 0.25;
    const value = { font, ascent, descent, height: ascent + descent, dot: context.measureText('.').width || size / 3 };
    metrics.set(key, value); return value;
  }
  function applyParagraphStyle(p, data) {
    const style = { ...(model.defaultStyle || {}), ...(data.style || {}) };
    tools.runStyle(p, style);
    p.style.fontSize = `${style.fontSize || model.defaultStyle?.fontSize || 14}px`;
    if (data.align) p.style.textAlign = data.align === 'both' ? 'justify' : data.align;
    p.style.marginLeft = `${data.indentLeft || 0}px`; p.style.marginRight = `${data.indentRight || 0}px`;
    p.style.textIndent = `${data.firstLineIndent || 0}px`;
    p.style.marginTop = `${data.spaceBefore || 0}px`; p.style.marginBottom = `${data.spaceAfter || 0}px`;
    p.style.whiteSpace = 'pre-wrap';
    if (data.isToc) { p.classList.add('op-word-toc'); p.dataset.level = String(data.tocLevel || 1); }
  }
  function paragraph(data, runs = data.runs || []) {
    if (data.numbering) data = { ...data, numberingLabel: data.numbering.text, numberingSuffix: data.numbering.suffix, numberingStyle: data.numbering.style };
    const p = node('p', 'op-word-paragraph'); applyParagraphStyle(p, data);
    const groups = [], pending = [], floating = [];
    let pieces = [];
    const flush = () => { groups.push(pieces); pieces = []; };
    if (data.numberingLabel) {
      pieces.push({ text: data.numberingLabel, ...(data.numberingStyle || {}) });
      if (data.numberingSuffix === 'tab') { flush(); pending.push({ numbering: true }); }
      else if (data.numberingSuffix !== 'nothing') pieces.push({ text: ' ' });
    }
    for (const original of runs) {
      const anchored = original.image?.anchor;
      if (anchored) floating.push({ image: original.image });
      const run = anchored ? { ...original, image: undefined } : original;
      if (run.type === 'pageBreak') continue;
      if (anchored && !run.text) continue;
      if (run.image) { pieces.push(run); continue; }
      const text = String(run.text ?? '');
      const split = text.split('\t');
      for (let i = 0; i < split.length; i++) {
        if (i) { flush(); pending.push({ style: run }); }
        if (split[i]) pieces.push({ ...run, text: split[i] });
      }
      if (run.type === 'tab' && !text.includes('\t')) { flush(); pending.push({ style: run }); }
    }
    flush();
    const tabItems = [];
    const right = (data.tabs || []).findLast?.(tab => tab.align === 'right');
    groups.forEach((group, index) => {
      if (index) {
        const gap = node('span', 'op-word-tab');
        const tabStyle = pending[index - 1]?.style || {};
        tools.runStyle(gap, { ...tabStyle, ...(tabStyle.fontFamilyLatin ? { fontFamily: tabStyle.fontFamilyLatin } : {}) }); gap.setAttribute('aria-hidden', 'true');
        p.append(gap); tabItems.push({ element: gap, numbering: pending[index - 1]?.numbering });
      }
      const tail = data.isToc && right && index === groups.length - 1 && index > 0;
      const span = node('span', tail ? 'op-word-page-number' : 'op-word-tab-content');
      if (tail) span.style.whiteSpace = 'nowrap';
      tools.appendRuns(span, group);
      const images = group.filter(run => run.image).map(run => run.image);
      for (const [imageIndex, image] of [...span.querySelectorAll('img')].entries()) {
        image.loading = 'lazy'; image.decoding = 'async';
        const original = images[imageIndex];
        if (original?.width && original?.height) {
          const available = Math.max(1, current.spec.width - current.spec.marginLeft - current.spec.marginRight - (data.indentLeft || 0) - (data.indentRight || 0));
          const scale = Math.min(1, available / original.width);
          image.style.width = `${original.width * scale}px`; image.style.height = `${original.height * scale}px`;
        }
      }
      p.append(span); if (index) tabItems[index - 1].following = span;
    });
    tabs.set(p, { data, items: tabItems, floating });
    return p;
  }
  function layoutParagraph(p, page) {
    const entry = tabs.get(p);
    if (!entry) return;
    const { data, items } = entry;
    const sizes = [...(data.runs || []).filter(run => !run.image?.anchor || run.text).map(run => run.fontSize), data.paragraphMarkStyle?.fontSize, data.numberingStyle?.fontSize].filter(value => value > 0);
    if (sizes.length) p.style.fontSize = `${Math.max(...sizes)}px`;
    let normal = fontMetrics(p).height;
    for (const span of p.querySelectorAll('span:not(.op-word-tab):not(.op-word-tab-leader)')) normal = Math.max(normal, fontMetrics(span).height);
    if (data.paragraphMarkStyle) {
      const probe = node('span'); tools.runStyle(probe, data.paragraphMarkStyle); p.append(probe); normal = Math.max(normal, fontMetrics(probe).height); probe.remove();
    }
    let line = data.lineHeight > 0 ? data.lineHeight : normal * (data.lineHeightRatio || 1);
    if (data.lineHeightRule === 'atLeast') line = Math.max(line, normal);
    const grid = page.spec.lineGrid || page.spec.linePitch || model.page?.lineGrid || model.page?.linePitch;
    if (grid && data.snapToGrid !== false) line = Math.ceil((line - 0.01) / grid) * grid;
    p.style.lineHeight = `${line}px`;
    p.style.minHeight = `${line}px`;
    const scale = pageScale(page), left = page.body.getBoundingClientRect().left;
    for (let i = 0; i < items.length; i++) {
      const tab = items[i], element = tab.element;
      element.style.width = '0px'; element.replaceChildren(); element.className = 'op-word-tab';
      const currentX = (element.getBoundingClientRect().left - left) / scale;
      let stop;
      if (tab.numbering) {
        const position = data.indentLeft || 0;
        const explicit = (data.tabs || []).find(value => value.align === 'left' && value.position > currentX + 0.1);
        stop = { position: position > currentX + 0.1 ? position : explicit?.position ?? (Math.floor(currentX / (model.defaultTabStop || 48)) + 1) * (model.defaultTabStop || 48), align: 'left', leader: 'none' };
      } else stop = [...(data.tabs || [])].sort((a, b) => a.position - b.position).find(value => value.position > currentX + 0.1);
      stop ||= { position: (Math.floor(currentX / (model.defaultTabStop || 48)) + 1) * (model.defaultTabStop || 48), align: 'left', leader: 'none' };
      let width = tab.following.getBoundingClientRect().width / scale;
      if (stop.align === 'left') width = 0;
      else if (stop.align === 'center') width /= 2;
      else if (stop.align === 'decimal') {
        const text = tab.following.textContent.split(/[.,]/)[0]; context.font = fontMetrics(tab.following).font; width = context.measureText(text).width;
      }
      const gap = Math.max(0, stop.position - currentX - width);
      element.style.width = `${gap}px`;
      if (stop.leader && stop.leader !== 'none' && gap > 0) {
        element.classList.add('op-word-tab-leader');
        const metric = fontMetrics(element), text = node('span', 'op-word-leader-glyphs');
        const character = ({ dot: '.', hyphen: '-', underscore: '_', middleDot: '·' })[stop.leader] || '.';
        context.font = metric.font;
        const advance = context.measureText(character).width || metric.dot;
        text.textContent = character.repeat(Math.min(10000, Math.ceil(gap / advance) + 1));
        text.style.height = `${metric.height}px`; text.style.lineHeight = `${metric.height}px`; text.style.bottom = `${-metric.descent}px`;
        element.append(text);
      }
    }
  }
  function positionFloating(element, page) {
    const paragraphs = element.matches('.op-word-paragraph') ? [element] : [...element.querySelectorAll('.op-word-paragraph')];
    const scale = pageScale(page), paper = page.paper.getBoundingClientRect(), spec = page.spec;
    for (const p of paragraphs) {
      const floating = tabs.get(p)?.floating || [], cell = p.closest('td');
      if (!cell || !floating.length) continue;
      const style = getComputedStyle(cell);
      if (!['middle', 'bottom'].includes(style.verticalAlign)) continue;
      const borderScale = getComputedStyle(cell.closest('table')).borderCollapse === 'collapse' ? 2 : 1;
      const available = cell.getBoundingClientRect().height / scale - parseFloat(style.paddingTop) - parseFloat(style.paddingBottom) -
        (parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth)) / borderScale;
      // Oversized paragraph-anchored floats keep the row's normal flow height,
      // but clamp vertical centering to the cell's top content edge.
      if (floating.some(({ image }) => image.anchor.relativeFromV === 'paragraph' && image.anchor.wrap === 'wrapNone' && image.height > available)) cell.style.verticalAlign = 'top';
    }
    for (const p of paragraphs) {
      const entry = tabs.get(p);
      if (!entry?.floating.length) continue;
      const paragraph = p.getBoundingClientRect(), cell = p.closest('td'), cellBox = cell?.getBoundingClientRect(), cellStyle = cell && getComputedStyle(cell);
      const columnLeft = cell ? (cellBox.left - paper.left) / scale + parseFloat(cellStyle.paddingLeft) : spec.marginLeft;
      const columnWidth = cell ? cellBox.width / scale - parseFloat(cellStyle.paddingLeft) - parseFloat(cellStyle.paddingRight) : spec.width - spec.marginLeft - spec.marginRight;
      for (const item of entry.floating) {
        const image = item.image, anchor = image.anchor;
        if (!item.frame) {
          item.frame = node('div', 'op-word-float'); tools.appendRuns(item.frame, [{ image }]);
          Object.assign(item.frame.style, { position: 'absolute', width: `${image.width || 0}px`, height: `${image.height || 0}px`, zIndex: anchor.behindText ? '0' : '2' });
          for (const img of item.frame.querySelectorAll('img')) { img.style.maxWidth = 'none'; img.style.display = 'block'; }
        }
        const horizontal = {
          page: [0, spec.width], margin: [spec.marginLeft, spec.width - spec.marginLeft - spec.marginRight],
          leftMargin: [0, spec.marginLeft], rightMargin: [spec.width - spec.marginRight, spec.marginRight],
          column: [columnLeft, columnWidth], character: [(paragraph.left - paper.left) / scale, 0],
        }[anchor.relativeFromH] || [columnLeft, columnWidth];
        const vertical = {
          page: [0, spec.height], margin: [spec.marginTop, spec.height - spec.marginTop - spec.marginBottom],
          topMargin: [0, spec.marginTop], bottomMargin: [spec.height - spec.marginBottom, spec.marginBottom],
          paragraph: [(paragraph.top - paper.top) / scale, paragraph.height / scale], line: [(paragraph.top - paper.top) / scale, paragraph.height / scale],
        }[anchor.relativeFromV] || [(paragraph.top - paper.top) / scale, paragraph.height / scale];
        const offset = (box, align, size, value) => box[0] + (align === 'center' ? (box[1] - size) / 2 : ['right', 'bottom', 'outside'].includes(align) ? box[1] - size : align ? 0 : value || 0);
        item.frame.style.left = `${offset(horizontal, anchor.alignH, image.width || 0, anchor.x)}px`;
        item.frame.style.top = `${offset(vertical, anchor.alignV, image.height || 0, anchor.y)}px`;
        page.body.style.zIndex = '1'; page.paper.append(item.frame);
      }
    }
  }
  function table(data) {
    const element = node('table', 'op-word-table');
    element.style.margin = '0';
    if (data.width > 0) element.style.width = `${data.width}px`;
    else if (data.widthPercent > 0) element.style.width = `${data.widthPercent}%`;
    else if (data.columnWidths?.length) element.style.width = `${data.columnWidths.reduce((sum, width) => sum + width, 0)}px`;
    if (data.align === 'center') element.style.margin = '0 auto';
    else if (data.align === 'right') element.style.marginLeft = 'auto';
    else if (data.indent) element.style.marginLeft = `${data.indent}px`;
    if (data.cellSpacing > 0) { element.style.borderCollapse = 'separate'; element.style.borderSpacing = `${data.cellSpacing}px`; }
    if (data.columnWidths?.length) {
      const columns = node('colgroup');
      for (const width of data.columnWidths) { const column = node('col'); column.style.width = `${width}px`; columns.append(column); }
      element.append(columns);
    }
    const borderCss = border => {
      const style = ({ single: 'solid', nil: 'none', thick: 'solid', dashSmallGap: 'dashed', dotDash: 'dashed', dotDotDash: 'dashed' })[border.style] || border.style || 'solid';
      return `${border.width || 0}px ${style} ${border.color || '#000'}`;
    };
    for (const [rowIndex, row] of (data.rows || []).entries()) {
      const tr = node('tr');
      const properties = data.rowProperties?.[rowIndex] || {};
      if (properties.height > 0 && properties.heightRule !== 'auto') rowHeights.set(tr, properties);
      for (const cell of row) {
        const td = node('td'); if (cell.colSpan) td.colSpan = cell.colSpan; if (cell.rowSpan) td.rowSpan = cell.rowSpan;
        if (cell.width > 0 && !data.columnWidths?.length) td.style.width = `${cell.width}px`;
        if (cell.verticalAlign) td.style.verticalAlign = cell.verticalAlign === 'center' ? 'middle' : cell.verticalAlign;
        if (cell.shading) td.style.backgroundColor = cell.shading;
        if (cell.noWrap) td.style.whiteSpace = 'nowrap';
        for (const side of ['top', 'right', 'bottom', 'left']) {
          const margin = cell.margins?.[side] ?? data.cellMargins?.[side];
          if (Number.isFinite(margin)) td.style[`padding${side[0].toUpperCase()}${side.slice(1)}`] = `${margin}px`;
          const outer = side === 'top' ? rowIndex === 0 : side === 'bottom' ? rowIndex + (cell.rowSpan || 1) === data.rows.length : side === 'left' ? cell.column === 0 : cell.column + (cell.colSpan || 1) === data.columnWidths?.length;
          const border = cell.borders?.[side] || data.borders?.[outer ? side : side === 'left' || side === 'right' ? 'insideV' : 'insideH'];
          if (border) td.style[`border${side[0].toUpperCase()}${side.slice(1)}`] = borderCss(border);
        }
        for (const block of cell.blocks || []) td.append(block.type === 'table' ? table(block) : paragraph(block));
        tr.append(td);
      }
      element.append(tr);
    }
    return element;
  }
  function layoutRows(element) {
    for (const row of element.querySelectorAll('tr')) {
      const properties = rowHeights.get(row);
      if (!properties) continue;
      let extra = 0;
      for (const cell of row.cells) {
        const style = getComputedStyle(cell), bottom = parseFloat(style.paddingBottom) || 0;
        // Native layout adds cell padding to an atLeast content floor. For
        // exact rows it adds only the largest bottom padding (MS-OI29500
        // 17.4.80). Measure effective styles, including inherited margins.
        const padding = properties.heightRule === 'exact' ? bottom :
          (parseFloat(style.paddingTop) || 0) + bottom +
          (parseFloat(style.borderTopWidth) + parseFloat(style.borderBottomWidth)) / (getComputedStyle(row.closest('table')).borderCollapse === 'collapse' ? 2 : 1);
        extra = Math.max(extra, padding);
      }
      row.style.height = `${properties.height + extra}px`;
    }
  }
  function renderPageArea(tag, blocks, page) {
    if (!blocks?.length) return;
    const region = node(tag, `op-word-page-${tag}`), { spec, paper, body } = page;
    const header = tag === 'header', edge = header ? 'top' : 'bottom';
    const distance = Math.max(0, spec[`${tag}Distance`] ?? 48);
    Object.assign(region.style, { position: 'absolute', left: `${spec.marginLeft}px`, right: `${spec.marginRight}px`, [edge]: `${distance}px` });
    for (const block of blocks) region.append(block.type === 'table' ? table(block) : paragraph(block));
    paper.insertBefore(region, header ? body : null);
    layoutRows(region);
    for (const p of region.querySelectorAll('.op-word-paragraph')) layoutParagraph(p, page);
    const margin = header ? 'marginTop' : 'marginBottom';
    if (spec[margin] >= 0) {
      spec[margin] = Math.max(spec[margin], distance + region.getBoundingClientRect().height / pageScale(page));
      paper.style[header ? 'paddingTop' : 'paddingBottom'] = `${spec[margin]}px`;
    }
    return region;
  }
  function place(element, data) {
    current.body.append(element);
    layoutRows(element);
    if (element.matches('.op-word-paragraph')) layoutParagraph(element, current);
    else for (const p of element.querySelectorAll('.op-word-paragraph')) layoutParagraph(p, current);
    const limit = hasSaved && savedFlow && !current.footer ? current.spec.height - current.spec.marginTop : current.height;
    if (current.body.children.length > 1 && current.body.scrollHeight > limit + 2) {
      const moving = [element];
      let previous = element.previousElementSibling;
      while (previous && tabs.get(previous)?.data.keepNext) { moving.unshift(previous); previous = previous.previousElementSibling; }
      if (moving.length < current.body.children.length) {
        for (const item of moving) item.remove();
        newPage();
        for (const item of moving) {
          current.body.append(item);
          if (item.matches('.op-word-paragraph')) layoutParagraph(item, current);
          else for (const p of item.querySelectorAll('.op-word-paragraph')) layoutParagraph(p, current);
        }
      }
    }
    positionFloating(element, current);
  }
  function addParagraph(data) {
    if (data.pageBreakBefore) newPage();
    let runs = [], continuation = false;
    const flush = () => {
      if (runs.length) {
        const copy = { ...data, runs, pageBreakBefore: false, ...(continuation ? { numbering: undefined, numberingLabel: undefined, firstLineIndent: 0, spaceBefore: 0 } : {}) };
        place(paragraph(copy, runs), copy); runs = []; continuation = true;
      }
    };
    for (const run of data.runs || []) {
      if (run.type === 'pageBreak' || run.text === '\f') { if (run.cached) savedFlow = true; flush(); newPage(); continue; }
      if (typeof run.text === 'string' && run.text.includes('\f')) {
        const parts = run.text.split('\f');
        parts.forEach((text, index) => { if (index) { flush(); newPage(); } if (text) runs.push({ ...run, text }); });
      } else runs.push(run);
    }
    if (!runs.length && !data.runs?.length) runs.push({ text: '' });
    flush();
  }
  const ready = (async () => {
    await document.fonts.ready;
    if (disposed) return;
    newPage(); setColumns();
    let count = 0;
    for (const data of model.blocks || []) {
      if (disposed) return;
      if (Number.isInteger(data.sectionIndex) && data.sectionIndex !== sectionIndex) { sectionIndex = data.sectionIndex; newPage(); }
      if (data.type === 'pageBreak') newPage();
      else if (data.type === 'table') place(table(data), data);
      else if (data.type === 'image') place(paragraph({ runs: [{ image: data }] }), data);
      else addParagraph(data);
      if (++count % 40 === 0) await new Promise(resolve => setTimeout(resolve, 0));
    }
    if (current && !current.body.children.length && allPages.length > 1) { current.paper.remove(); allPages.pop(); pageCount--; }
    for (const page of allPages) {
      positionFloating(page.body, page);
      // Tall inline pictures may occupy the bottom margin without exceeding
      // the physical sheet. Padding must not enlarge those paper pages.
      const limit = page.footer ? page.height : page.spec.height - page.spec.marginTop;
      if (page.body.scrollHeight > limit + 3) {
        page.paper.dataset.overflow = 'true';
        page.paper.style.height = `${page.spec.marginTop + page.body.scrollHeight + page.spec.marginBottom}px`;
      }
      if (page.header) positionFloating(page.header, page);
      if (page.footer) positionFloating(page.footer, page);
    }
    pages.dataset.ready = 'true';
  })();
  return { ready, get fitWidth() { return setColumns(); }, dispose() { disposed = true; tabs.clear(); metrics.clear(); rowHeights.clear(); pages.remove(); } };
}
