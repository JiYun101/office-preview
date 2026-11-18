function element(tag, className, text) {
  const node = document.createElement(tag); node.className = className;
  if (text !== undefined) node.textContent = String(text);
  return node;
}
function trace(context, commands = []) {
  context.beginPath();
  for (const [op, ...args] of commands) {
    if (op === 'M') context.moveTo(...args);
    else if (op === 'L') context.lineTo(...args);
    else if (op === 'C') context.bezierCurveTo(...args);
    else if (op === 'Z') context.closePath();
  }
}
const cancelled = () => new DOMException('PDF 页面绘制已取消。', 'AbortError');
let fontSession = 0;

function strokeStyle(context, operation, scale) {
  context.strokeStyle = operation.stroke || operation.color || '#000';
  context.lineWidth = operation.lineWidth > 0 ? operation.lineWidth : 1 / scale;
  context.lineCap = ['butt', 'round', 'square'][operation.lineCap] || 'butt';
  context.lineJoin = ['miter', 'round', 'bevel'][operation.lineJoin] || 'miter';
  context.miterLimit = operation.miterLimit ?? 10;
  context.setLineDash(operation.dash || []); context.lineDashOffset = operation.dashPhase || 0;
}

async function decodeImage(operation, state) {
  if (operation.rgba) {
    const width = operation.pixelWidth, height = operation.pixelHeight;
    if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || width * height > 16000000 || operation.rgba.length !== width * height * 4) throw new Error('PDF 图片尺寸无效或超过绘制上限。');
    const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
    const pixels = new Uint8ClampedArray(operation.rgba.buffer, operation.rgba.byteOffset, operation.rgba.byteLength);
    canvas.getContext('2d').putImageData(new ImageData(pixels, width, height), 0, 0);
    return canvas;
  }
  if (!/^image\/(jpeg|png|webp|gif)$/i.test(operation.mime || '')) throw new Error('PDF 内嵌图片类型尚未支持。');
  const url = URL.createObjectURL(new Blob([operation.bytes], { type: operation.mime }));
  const image = new Image();
  return new Promise((resolve, reject) => {
    let settled = false;
    const finish = (error) => {
      if (settled) return;
      settled = true; URL.revokeObjectURL(url);
      if (state.abort === abort) state.abort = null;
      if (error) { image.removeAttribute('src'); reject(error); }
      else resolve(image);
    };
    const abort = () => finish(cancelled());
    state.abort = abort;
    image.src = url;
    image.decode().then(() => finish(), error => finish(error));
  });
}

export function renderPdf(model, viewport, content, auxiliary) {
  let disposed = false, active = false, firstSettled = false, fontsPrepared = false;
  const fonts = new Map(), fontFaces = new Set(), session = ++fontSession;
  const pages = model.pages || [];
  const cssScale = 96 / 72;
  const readZoom = () => {
    const value = getComputedStyle(content).zoom, number = parseFloat(value);
    if (Number.isFinite(number) && number > 0) return String(value).trim().endsWith('%') ? number / 100 : number;
    return parseFloat(content.style.getPropertyValue('--op-scale')) || 1;
  };
  const readDensity = () => globalThis.devicePixelRatio || 1;
  let zoom = readZoom(), density = readDensity();
  const stack = element('div', 'op-pdf-stack');
  Object.assign(stack.style, { display: 'flex', flexDirection: 'column', alignItems: 'center', gap: '24px', margin: '0 auto' });
  content.append(stack);
  const states = pages.map((page, index) => {
    if (!(page.width > 0) || !(page.height > 0) || !Number.isFinite(page.width + page.height)) throw new Error('PDF 页面尺寸无效。');
    const slot = element('section', 'op-pdf-slot');
    slot.dataset.page = String(index + 1); slot.dataset.state = 'idle';
    slot.setAttribute('aria-label', `PDF 第 ${index + 1} 页`);
    Object.assign(slot.style, { position: 'relative', flex: 'none', width: `${page.width * cssScale}px`, height: `${page.height * cssScale}px`, background: 'white' });
    stack.append(slot);
    return { page, index, slot, token: 0, status: 'idle', wanted: false, canvas: null, drawingCanvas: null, abort: null };
  });
  const width = pages.reduce((maximum, page) => Math.max(maximum, page.width * cssScale), 0);
  stack.style.width = `${width || 816}px`;
  const queued = new Set();
  let resolveFirst, rejectFirst;
  const ready = new Promise((resolve, reject) => { resolveFirst = resolve; rejectFirst = reject; });
  const settleFirst = error => {
    if (firstSettled) return;
    firstSettled = true;
    error ? rejectFirst(error) : resolveFirst();
  };
  const release = (state, keepCanvas = false) => {
    state.token++; queued.delete(state); state.abort?.(); state.abort = null;
    if (state.drawingCanvas && state.drawingCanvas !== state.canvas) {
      state.drawingCanvas.width = 1; state.drawingCanvas.height = 1; state.drawingCanvas.remove();
    }
    state.drawingCanvas = null;
    if (state.canvas && (!keepCanvas || state.canvas.dataset.ready !== 'true')) { state.canvas.width = 1; state.canvas.height = 1; state.canvas.remove(); state.canvas = null; }
    if (state.status !== 'error') { state.status = 'idle'; state.slot.dataset.state = 'idle'; }
  };
  const draw = async (state, token) => {
    const page = state.page;
    const current = () => !disposed && token === state.token;
    const scale = Math.min(cssScale * density * zoom, 8192 / Math.max(page.width, page.height), Math.sqrt(16000000 / (page.width * page.height)));
    const canvas = element('canvas', 'op-pdf-page');
    canvas.width = Math.max(1, Math.floor(page.width * scale)); canvas.height = Math.max(1, Math.floor(page.height * scale));
    canvas.style.width = `${page.width * cssScale}px`; canvas.style.height = `${page.height * cssScale}px`;
    canvas.setAttribute('aria-label', `PDF 第 ${state.index + 1} 页`);
    state.drawingCanvas = canvas;
    if (!state.canvas) { state.canvas = canvas; state.slot.replaceChildren(canvas); }
    const context = canvas.getContext('2d');
    context.fillStyle = 'white'; context.fillRect(0, 0, canvas.width, canvas.height);
    context.setTransform(scale, 0, 0, -scale, 0, canvas.height);
    for (const operation of page.operations || []) {
      if (!current()) throw cancelled();
      if (operation.type === 'save') context.save();
      else if (operation.type === 'restore') context.restore();
      else if (operation.type === 'transform') context.transform(...operation.matrix);
      else if (operation.type === 'clip') { trace(context, operation.commands); context.clip(operation.fillRule || 'nonzero'); }
      else if (operation.type === 'path') {
        if (operation.blend === 'multiply') { context.save(); context.globalCompositeOperation = 'multiply'; }
        trace(context, operation.commands);
        if (operation.fill) { context.globalAlpha = operation.fillOpacity ?? 1; context.fillStyle = operation.fill; context.fill(operation.fillRule || 'nonzero'); }
        if (operation.stroke) { context.globalAlpha = operation.strokeOpacity ?? 1; strokeStyle(context, operation, scale); context.stroke(); }
        if (operation.blend === 'multiply') context.restore();
      } else if (operation.type === 'text') {
        const mode = operation.textMode ?? 0;
        if (mode === 3 || mode === 7) continue;
        const font = fonts.get(operation.fontId), embedded = font && operation.renderText !== undefined;
        const text = embedded ? operation.renderText : operation.text || '';
        const family = embedded ? `"${font}"` : operation.fallbackFamily || operation.fontFamily || 'sans-serif';
        context.save(); context.transform(...operation.matrix); context.scale(1, -1);
        context.font = `${operation.fontStyle || 'normal'} ${operation.fontWeight || 400} ${operation.fontSize || 12}px ${family}`;
        context.fillStyle = operation.color || '#000';
        // PDF widths advance the next glyph's origin; embedded outlines already
        // have the correct proportions and must not be stretched to that width.
        if (!embedded) {
          const measured = context.measureText(text).width;
          if (operation.width > 0 && measured > 0) context.scale(operation.width / measured, 1);
        }
        if (mode % 4 === 0 || mode % 4 === 2) { context.globalAlpha = operation.fillOpacity ?? 1; context.fillText(text, 0, 0); }
        if (mode % 4 === 1 || mode % 4 === 2) { context.globalAlpha = operation.strokeOpacity ?? 1; strokeStyle(context, operation, scale); context.strokeText(text, 0, 0); }
        context.restore();
      } else if (operation.type === 'image') {
        const image = await decodeImage(operation, state);
        try {
          if (!current()) throw cancelled();
          context.save();
          context.globalAlpha = operation.fillOpacity ?? 1;
          if (operation.matrix) context.transform(...operation.matrix);
          context.translate(operation.x || 0, (operation.y || 0) + (operation.height ?? 1)); context.scale(1, -1);
          context.drawImage(image, 0, 0, operation.width ?? 1, operation.height ?? 1); context.restore();
        } finally {
          if (operation.rgba) { image.width = 1; image.height = 1; }
          else image.removeAttribute('src');
        }
      }
    }
    if (!current()) throw cancelled();
    canvas.dataset.ready = 'true';
    if (state.canvas && state.canvas !== canvas) { state.canvas.width = 1; state.canvas.height = 1; state.canvas.remove(); }
    state.canvas = canvas; state.drawingCanvas = null; state.slot.replaceChildren(canvas);
  };
  const pump = async () => {
    if (disposed || active || !fontsPrepared || !queued.size) return;
    const state = queued.values().next().value; queued.delete(state);
    if (!state.wanted && !(state.index === 0 && !firstSettled)) { pump(); return; }
    active = true; state.status = 'drawing'; state.slot.dataset.state = 'drawing';
    const token = ++state.token;
    try {
      await draw(state, token);
      if (disposed || token !== state.token) return;
      state.status = 'ready'; state.slot.dataset.state = 'ready';
      if (state.index === 0) settleFirst();
    } catch (error) {
      if (disposed || token !== state.token || error.name === 'AbortError') return;
      if (state.drawingCanvas && state.drawingCanvas !== state.canvas) { state.drawingCanvas.width = 1; state.drawingCanvas.height = 1; }
      state.drawingCanvas = null;
      if (state.canvas) { state.canvas.width = 1; state.canvas.height = 1; state.canvas = null; }
      state.status = 'error'; state.slot.dataset.state = 'error';
      state.slot.replaceChildren(element('div', 'op-error', error.message));
      if (state.index === 0) settleFirst(error);
    } finally {
      active = false;
      if (!state.wanted && !(state.index === 0 && !firstSettled)) release(state);
      pump();
    }
  };
  const enqueue = state => {
    if (disposed || state.status !== 'idle') return;
    state.status = 'queued'; queued.add(state); pump();
  };
  const observer = new IntersectionObserver(entries => {
    for (const entry of entries) {
      const state = states[Number(entry.target.dataset.page) - 1];
      state.wanted = entry.isIntersecting;
      if (state.wanted) enqueue(state);
      else if (!(state.index === 0 && !firstSettled)) release(state);
    }
  }, { root: viewport, rootMargin: '600px 0px' });
  const refreshScale = () => {
    if (disposed) return;
    const value = readZoom(), nextDensity = readDensity();
    const changedDensity = Math.abs(nextDensity - density) >= 0.000001;
    if (Math.abs(value - zoom) < 0.000001 && !changedDensity) return;
    zoom = value; density = nextDensity;
    if (changedDensity) watchDensity();
    for (const state of states) {
      if (state.status === 'error' || !state.wanted && !(state.index === 0 && !firstSettled)) continue;
      // Keep the last completed bitmap visible until its sharper replacement
      // finishes; only near-viewport pages participate in a zoom redraw.
      release(state, true); enqueue(state);
    }
  };
  const zoomObserver = new MutationObserver(refreshScale);
  zoomObserver.observe(content, { attributes: true, attributeFilter: ['style'] });
  let resolutionQuery;
  const watchDensity = () => {
    resolutionQuery?.removeEventListener('change', densityChanged);
    resolutionQuery = globalThis.matchMedia?.(`(resolution: ${readDensity()}dppx)`);
    resolutionQuery?.addEventListener('change', densityChanged);
  };
  const densityChanged = () => refreshScale();
  watchDensity(); globalThis.addEventListener('resize', refreshScale);
  for (const state of states) observer.observe(state.slot);
  if (states.length) { states[0].wanted = true; enqueue(states[0]); }
  else { stack.append(element('div', 'op-empty', '没有可预览的 PDF 页面。')); settleFirst(); }
  const prepareFonts = async () => {
    await Promise.all((model.fonts || []).filter(font => font.program).map(async font => {
      try {
        const family = `OfficePreviewPdf_${session}_${font.id}`;
        const face = new FontFace(family, font.program.bytes, { weight: String(font.weight || 400), style: font.style || 'normal' });
        await face.load();
        if (disposed) return;
        document.fonts.add(face); fontFaces.add(face); fonts.set(font.id, family);
      } catch {
        if (disposed) return;
        const warning = 'PDF 内嵌字体未能加载，已使用系统字体近似显示。';
        model.warnings ||= [];
        if (!model.warnings.includes(warning)) model.warnings.push(warning);
      }
    }));
    if (disposed) return;
    fontsPrepared = true; pump();
  };
  prepareFonts().catch(settleFirst);
  return {
    ready, fitWidth: (width || 816) + 56,
    dispose() {
      if (disposed) return;
      disposed = true; observer.disconnect(); zoomObserver.disconnect();
      resolutionQuery?.removeEventListener('change', densityChanged); globalThis.removeEventListener('resize', refreshScale);
      queued.clear(); for (const state of states) release(state);
      for (const face of fontFaces) document.fonts.delete(face);
      fontFaces.clear(); fonts.clear(); settleFirst(cancelled()); stack.remove();
    }
  };
}
