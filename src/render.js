import { renderPdf } from './render-pdf.js';
import { renderEmail } from './render-email.js';
import { renderSheet } from './render-sheet.js';
import { gradientCss, colorOpacity } from './slide-paint.js';
import { renderWord } from './render-word.js';

function node(tag, className, text) {
  const element = document.createElement(tag);
  if (className) element.className = className;
  if (text !== undefined) element.textContent = String(text);
  return element;
}
function fontStack(name) {
  const quoted = '"' + String(name).replaceAll('\\', '\\\\').replaceAll('"', '\\"') + '"';
  if (/楷体|KaiTi/i.test(name)) return `${quoted},KaiTi,"Kaiti SC",STKaiti,serif`;
  if (/仿宋|FangSong/i.test(name)) return `${quoted},FangSong,STFangsong,serif`;
  if (/宋体|小标宋|SimSun/i.test(name)) return `${quoted},SimSun,"Songti SC",serif`;
  if (/黑体|SimHei/i.test(name)) return `${quoted},SimHei,"Heiti SC","PingFang SC",sans-serif`;
  return `${quoted},"PingFang SC","Microsoft YaHei",sans-serif`;
}
function runStyle(element, style = {}) {
  if (style.bold) element.style.fontWeight = '700';
  if (style.italic) element.style.fontStyle = 'italic';
  if (style.underline || style.strike) element.style.textDecoration = [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ');
  if (style.color) element.style.color = style.color;
  if (style.background) element.style.backgroundColor = style.background;
  if (Number.isFinite(style.fontSize)) element.style.fontSize = `${Math.max(1, Math.min(500, style.fontSize))}px`;
  if (style.fontFamily) element.style.fontFamily = fontStack(style.fontFamily);
  if (Number.isFinite(style.letterSpacing)) element.style.letterSpacing = `${style.letterSpacing}px`;
  if (Number.isFinite(style.baselineShift) && style.baselineShift) { element.style.position = 'relative'; element.style.top = `${-style.baselineShift}px`; }
}

export function renderModel(model, viewport, content, auxiliary) {
  const urls = new Set();
  const wordImageUrls = model.kind === 'document' ? new WeakMap() : null;
  let activeUrls = null;
  const image = (data, lazy = false) => {
    const element = node('img');
    if (lazy) { element.loading = 'lazy'; element.decoding = 'async'; }
    if (!data.bytes || !/^image\/(png|jpeg|jpg|gif|webp|bmp)$/i.test(data.mime ?? '')) {
      const fallback = node('span', '', '〔暂不支持的图片格式〕');
      return fallback;
    }
    let url = wordImageUrls?.get(data);
    if (!url) {
      url = URL.createObjectURL(new Blob([data.bytes], { type: data.mime }));
      urls.add(url); activeUrls?.add(url); wordImageUrls?.set(data, url);
    }
    element.src = url; element.alt = '文档内嵌图片';
    if (data.width) element.style.width = `${data.width}px`;
    if (data.height) element.style.height = `${data.height}px`;
    return element;
  };
  const appendRuns = (parent, runs = [], options = {}) => {
    for (const run of runs) {
      if (run.image) parent.append(image(run.image, options.lazy));
      if (run.text !== undefined) { const span = node('span', '', run.text); runStyle(span, run); parent.append(span); }
    }
  };
  content.replaceChildren(); auxiliary.replaceChildren();
  if (model.kind === 'pdf') return renderPdf(model, viewport, content, auxiliary);
  if (model.kind === 'email') return renderEmail(model, content);
  if (model.kind === 'document') {
    const rendered = renderWord(model, viewport, content, auxiliary, { appendRuns: (parent, runs) => appendRuns(parent, runs, { lazy: true }), runStyle });
    return { ready: rendered.ready, get fitWidth() { return rendered.fitWidth; }, dispose() { rendered.dispose(); for (const url of urls) URL.revokeObjectURL(url); urls.clear(); } };
  }
  if (model.kind === 'spreadsheet') return renderSheet(model, viewport, content, auxiliary, { runStyle, fontStack });
  if (model.kind === 'presentation') {
    const slides = model.slides || [], width = model.width || 960, height = model.height || 540;
    const stack = node('div', 'op-slide-stack'); stack.style.width = `${width}px`;
    content.append(stack);
    let disposed = false;
    const states = slides.map((slide, index) => {
      const slot = node('article', 'op-slide-page'), canvas = node('div', 'op-slide');
      slot.dataset.page = String(index + 1); slot.setAttribute('aria-label', `幻灯片 ${index + 1}`);
      canvas.style.width = `${width}px`; canvas.style.height = `${height}px`;
      if (slide.background) canvas.style.backgroundColor = slide.background;
      if (slide.backgroundPaint) canvas.style.backgroundImage = gradientCss(slide.backgroundPaint, width, height);
      slot.append(canvas); stack.append(slot);
      return { slot, canvas, urls: new Set(), hydrated: false };
    });
    const release = state => {
      for (const url of state.urls) { URL.revokeObjectURL(url); urls.delete(url); }
      state.urls.clear(); state.canvas.replaceChildren(); state.hydrated = false; delete state.slot.dataset.ready;
    };
    const slideText = (item, className = 'op-slide-text') => {
      const element = node('div', className); element.style.fontSize = `${item.fontSize || 24}px`;
      runStyle(element, item);
      element.style.overflow = item.verticalOverflow === 'clip' ? 'hidden' : 'visible';
      if (item.align) element.style.textAlign = item.align;
      element.style.justifyContent = ({ middle: 'center', center: 'center', bottom: 'flex-end' }[item.verticalAlign]) || 'flex-start';
      const body = node('div', 'op-slide-text-content');
      if (item.paragraphs?.length) {
        for (const paragraph of item.paragraphs) {
          const p = node('p', 'op-slide-paragraph'); appendRuns(p, paragraph.runs);
          const sizes = (paragraph.runs || []).map(run => run.fontSize).filter(value => value > 0);
          const fontSize = sizes.length ? Math.max(...sizes) : item.fontSize || 24;
          p.style.fontSize = `${fontSize}px`;
          const family = paragraph.runs?.find(run => run.fontFamily)?.fontFamily;
          if (family) p.style.fontFamily = fontStack(family);
          if (paragraph.align) p.style.textAlign = paragraph.align;
          const normalLine = fontSize * 1.2;
          const lineHeight = paragraph.lineHeightPx > 0 ? paragraph.lineHeightPx : normalLine * (paragraph.lineHeight > 0 ? paragraph.lineHeight : 1);
          p.style.lineHeight = `${lineHeight}px`;
          if (paragraph.lineHeight > 0 && !paragraph.lineHeightPx) {
            // Keep the extra initial leading rather than letting CSS halve it.
            const leading = (paragraph.lineHeight - 1) * fontSize - (lineHeight - normalLine) / 2;
            p.style.position = 'relative'; p.style.top = `${leading}px`;
          }
          const before = paragraph.spaceBefore ?? (paragraph.spaceBeforeRatio || 0) * lineHeight;
          const after = paragraph.spaceAfter ?? (paragraph.spaceAfterRatio || 0) * lineHeight;
          if (before) p.style.marginTop = `${before}px`;
          if (after) p.style.marginBottom = `${after}px`;
          body.append(p);
        }
      } else {
        const sizes = (item.runs || []).map(run => run.fontSize).filter(value => value > 0);
        if (sizes.length) body.style.fontSize = `${Math.max(...sizes)}px`;
        appendRuns(body, item.runs);
      }
      element.append(body);
      return element;
    };
    const slideTable = item => {
      const table = node('div', 'op-slide-table'); table.setAttribute('role', 'table');
      Object.assign(table.style, {
        position: 'absolute', display: 'grid', color: '#000',
        gridTemplateColumns: item.columns.map(width => `${width}px`).join(' '),
        gridTemplateRows: item.rows.map(row => `${row.height}px`).join(' '),
      });
      const gridline = item.styled ? '1px solid #d1d5db' : '';
      item.rows.forEach((row, rowIndex) => {
        const element = node('div'); element.setAttribute('role', 'row'); element.style.display = 'contents';
        for (const cell of row.cells) {
          const box = slideText({ ...cell, fontSize: cell.fontSize ?? 18, verticalOverflow: 'clip' }, 'op-slide-table-cell');
          box.setAttribute('role', 'cell'); box.setAttribute('aria-colindex', String(cell.column + 1));
          if (cell.colSpan > 1) box.setAttribute('aria-colspan', String(cell.colSpan));
          if (cell.rowSpan > 1) box.setAttribute('aria-rowspan', String(cell.rowSpan));
          Object.assign(box.style, {
            display: 'flex', flexDirection: 'column', overflowWrap: 'anywhere', minWidth: '0', minHeight: '0',
            gridRow: `${rowIndex + 1} / span ${cell.rowSpan}`, gridColumn: `${cell.column + 1} / span ${cell.colSpan}`,
            padding: cell.padding.map(size => `${size}px`).join(' '), backgroundColor: cell.fill || '',
            borderTop: cell.borders.top || gridline, borderLeft: cell.borders.left || gridline,
            borderRight: cell.borders.right || (cell.column + cell.colSpan === item.columns.length ? gridline : ''),
            borderBottom: cell.borders.bottom || (rowIndex + cell.rowSpan === item.rows.length ? gridline : ''),
          });
          element.append(box);
        }
        table.append(element);
      });
      return table;
    };
    const hydrate = index => {
      const state = states[index];
      if (disposed || !state || state.hydrated) return;
      const slide = slides[index], canvas = state.canvas;
      activeUrls = state.urls;
      try {
        for (const item of slide.elements || []) {
          let element;
          if (item.type === 'image') {
            element = image(item, true); element.className = 'op-slide-image';
            if (model.format === 'ppt') element.style.objectFit = 'fill';
            const crop = item.crop;
            if (crop && [crop.left, crop.top, crop.right, crop.bottom].some(value => value)) {
              const frame = node('div', 'op-slide-image-frame');
              const width = Math.max(0.01, 1 - (crop.left || 0) - (crop.right || 0));
              const height = Math.max(0.01, 1 - (crop.top || 0) - (crop.bottom || 0));
              Object.assign(element.style, { left: `${-(crop.left || 0) / width * 100}%`, top: `${-(crop.top || 0) / height * 100}%`, width: `${100 / width}%`, height: `${100 / height}%`, objectFit: 'fill' });
              frame.append(element); element = frame;
            }
          }
          else if (item.type === 'shape') {
            element = node('div', 'op-slide-shape'); element.style.backgroundColor = colorOpacity(item.fill || 'transparent', item.fillOpacity);
            if (item.fillPaint) element.style.backgroundImage = gradientCss(item.fillPaint, item.width, item.height);
            if (item.shape === 'ellipse') element.style.borderRadius = '50%';
            if (item.shape === 'roundRect') element.style.borderRadius = `${item.cornerRadius ?? Math.min(item.width, item.height) / 6}px`;
            if (item.stroke && (item.strokeWidth ?? 1) > 0) {
              const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
              const width = Math.max(1, item.width), height = Math.max(1, item.height);
              svg.setAttribute('viewBox', `0 0 ${width} ${height}`); svg.setAttribute('aria-hidden', 'true');
              Object.assign(svg.style, { position: 'absolute', left: '0', top: '0', width: `${width}px`, height: `${height}px`, overflow: 'visible', pointerEvents: 'none' });
              const tag = item.shape === 'ellipse' ? 'ellipse' : item.shape === 'line' ? 'line' : 'rect';
              const path = document.createElementNS('http://www.w3.org/2000/svg', tag);
              const attributes = tag === 'ellipse' ? { cx: item.width / 2, cy: item.height / 2, rx: item.width / 2, ry: item.height / 2 } : tag === 'line' ? { x1: 0, y1: 0, x2: item.width, y2: item.height } : { x: 0, y: 0, width: item.width, height: item.height, ...(item.shape === 'roundRect' ? { rx: item.cornerRadius ?? Math.min(item.width, item.height) / 6 } : {}) };
              for (const [key, value] of Object.entries(attributes)) path.setAttribute(key, String(value));
              path.setAttribute('fill', 'none'); path.setAttribute('stroke', colorOpacity(item.stroke, item.strokeOpacity));
              path.setAttribute('stroke-width', String(item.strokeWidth ?? 1));
              svg.append(path); element.append(svg); element.style.overflow = 'visible';
            }
          } else if (item.type === 'table') element = slideTable(item);
          else element = slideText(item);
          Object.assign(element.style, { left: `${item.x || 0}px`, top: `${item.y || 0}px`, width: `${item.width || 0}px`, height: `${item.height || 0}px` });
          const transform = [];
          if (item.rotation) transform.push(`rotate(${item.rotation}deg)`);
          if (item.flipH) transform.push('scaleX(-1)');
          if (item.flipV) transform.push('scaleY(-1)');
          if (transform.length) element.style.transform = transform.join(' ');
          if (item.transformOrigin) element.style.transformOrigin = `${item.transformOrigin.x}px ${item.transformOrigin.y}px`;
          canvas.append(element);
        }
        state.hydrated = true; state.slot.dataset.ready = 'true';
      } finally { activeUrls = null; }
    };
    hydrate(0);
    const observer = new IntersectionObserver(entries => {
      if (disposed) return;
      for (const entry of entries) {
        const index = Number(entry.target.dataset.page) - 1;
        if (entry.isIntersecting) hydrate(index); else release(states[index]);
      }
    }, { root: viewport, rootMargin: '800px 0px' });
    for (const state of states) observer.observe(state.slot);
    if (!slides.length) stack.append(node('div', 'op-empty', '没有可显示的幻灯片。'));
    return { ready: Promise.resolve(), fitWidth: width + 24, dispose() { disposed = true; observer.disconnect(); for (const state of states) release(state); stack.remove(); } };
  }
  return { fitWidth: 850, dispose() { for (const url of urls) URL.revokeObjectURL(url); urls.clear(); } };
}
