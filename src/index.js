import { viewerStyles } from './styles.js';
import { renderModel } from './render.js';
import { capabilities } from './parse.js';
export { capabilities, parseDocument } from './parse.js';

export class OfficePreview extends EventTarget {
  constructor(container, options = {}) {
    super();
    if (!(container instanceof HTMLElement)) throw new TypeError('OfficePreview 需要一个 HTML 容器。');
    this.host = document.createElement('div'); this.host.style.height = '100%'; container.append(this.host);
    this.root = this.host.attachShadow({ mode: 'open' });
    const style = document.createElement('style'); style.textContent = viewerStyles;
    this.root.append(style);
    this.shell = this.el('div', 'op-shell');
    this.shell.style.setProperty('--op-zoom-safe-area', options.showZoom === false ? '0px' : '64px');
    this.shell.setAttribute('role', 'region'); this.shell.setAttribute('aria-label', '文件预览');
    const zoomControls = this.el('div', 'op-zoom-controls');
    zoomControls.setAttribute('role', 'group'); zoomControls.setAttribute('aria-label', '文档缩放');
    this.scaleLabel = this.el('span', 'op-scale', '100%');
    const control = (label, action, title = label) => { const b = this.el('button', '', label); b.type = 'button'; b.title = title; b.setAttribute('aria-label', title); b.onclick = action; return b; };
    zoomControls.append(control('−', () => this.setZoom(this.zoom - 0.1), '缩小'), this.scaleLabel, control('+', () => this.setZoom(this.zoom + 0.1), '放大'), control('↔', () => this.fit(), '适合宽度'));
    this.auxiliary = this.el('div', 'op-auxiliary');
    this.viewport = this.el('div', 'op-viewport');
    this.viewport.tabIndex = 0; this.viewport.setAttribute('aria-label', '文件内容');
    this.content = this.el('div', 'op-content'); this.viewport.append(this.content);
    this.shell.append(this.auxiliary, this.viewport);
    if (options.showZoom !== false) this.shell.append(zoomControls);
    this.root.append(this.shell);
    this.zoom = 1; this._serial = 0; this._destroyed = false;
    this._initialZoom = Number.isFinite(options.zoom) && options.zoom > 0 ? options.zoom : null;
    this._autoFit = true;
    this._resize = new ResizeObserver(() => { if (this._autoFit && this.rendered) this.fit(); });
    this._resize.observe(this.viewport);
    this.message('等待打开文件', '');
    if (this._initialZoom) this.setZoom(this._initialZoom);
  }
  el(tag, className, text) { const element = document.createElement(tag); element.className = className; if (text !== undefined) element.textContent = text; return element; }
  message(title, detail, error = false) { const box = this.el('div', error ? 'op-error' : 'op-empty'); box.append(this.el('strong', '', title), this.el('span', '', detail)); this.content.replaceChildren(box); }
  setZoom(value) { this._autoFit = false; this.zoom = Math.max(0.2, Math.min(3, Number(value) || 1)); this.content.style.setProperty('--op-scale', String(this.zoom)); this.scaleLabel.textContent = `${Math.round(this.zoom * 100)}%`; }
  fit() { this.setZoom(this.rendered?.fitWidth ? Math.min(1.5, this.viewport.clientWidth / this.rendered.fitWidth) : 1); this._autoFit = true; }
  getDocument() { return this.model ?? null; }
  getCapabilities() { return capabilities; }
  async open(input, options = {}) {
    if (this._destroyed) throw new Error('预览实例已销毁。');
    const serial = ++this._serial;
    this._cancel?.();
    this.rendered?.dispose(); this.rendered = null; this.model = null;
    this.content.removeAttribute('style'); this.setZoom(this._initialZoom || 1); this.auxiliary.replaceChildren();
    this.shell.setAttribute('aria-label', options.name || input?.name || '未命名文件');
    this.viewport.scrollTo(0, 0);
    this.message('正在读取文件…', '');
    try {
      if ((input?.size ?? input?.byteLength ?? 0) > 32 * 1024 * 1024) throw new Error('文件超过首版 32MB 限制。');
      let buffer;
      if (input instanceof Blob) buffer = await input.arrayBuffer();
      else if (input instanceof ArrayBuffer) buffer = input.slice(0);
      else if (ArrayBuffer.isView(input)) buffer = input.buffer.slice(input.byteOffset, input.byteOffset + input.byteLength);
      else throw new TypeError('请传入 File、Blob、ArrayBuffer 或 Uint8Array。');
      if (serial !== this._serial) throw new DOMException('已取消打开文件。', 'AbortError');
      const parseOptions = { ...options, name: options.name || input?.name || '未命名' };
      const model = await new Promise((resolve, reject) => {
        const worker = new Worker(new URL('./worker.js', import.meta.url), { type: 'module' });
        const finish = () => { worker.terminate(); if (serial === this._serial) this._cancel = null; };
        this._cancel = () => { finish(); reject(new DOMException('已取消打开文件。', 'AbortError')); };
        worker.onmessage = event => { finish(); event.data.ok ? resolve(event.data.model) : reject(new Error(event.data.error)); };
        worker.onerror = () => { finish(); reject(new Error('解析工作线程加载失败，请通过 HTTP / HTTPS 服务加载 SDK，并检查 CSP worker-src 配置。')); };
        try { worker.postMessage({ buffer, options: parseOptions }, [buffer]); }
        catch (error) { finish(); reject(error); }
      });
      if (serial !== this._serial) throw new DOMException('已取消打开文件。', 'AbortError');
      this.model = model; this.shell.setAttribute('aria-label', model.name || '文件预览');
      this.rendered = renderModel(model, this.viewport, this.content, this.auxiliary);
      const cancelled = new Promise((_, reject) => {
        this._cancel = () => reject(new DOMException('已取消打开文件。', 'AbortError'));
      });
      try { await Promise.race([Promise.resolve(this.rendered.ready), cancelled]); }
      finally { if (serial === this._serial) this._cancel = null; }
      if (serial !== this._serial) throw new DOMException('已取消打开文件。', 'AbortError');
      if (this._initialZoom) this.setZoom(this._initialZoom); else this.fit();
      this.dispatchEvent(new CustomEvent('load', { detail: model })); return model;
    } catch (error) {
      if (serial === this._serial && error.name !== 'AbortError') {
        this.model = null;
        this.rendered?.dispose(); this.rendered = null;
        this.message('暂时无法预览此文件', error.message, true);
        this.dispatchEvent(new CustomEvent('error', { detail: error }));
      }
      throw error;
    }
  }
  destroy() { if (this._destroyed) return; this._destroyed = true; this._serial++; this._cancel?.(); this.rendered?.dispose(); this._resize.disconnect(); this.host.remove(); this.model = null; }
}
