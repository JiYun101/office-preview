import { OfficePreview } from '../src/index.js';

const params = new URLSearchParams(location.search);
const viewer = new OfficePreview(document.querySelector('#viewer'), { showZoom: params.get('showZoom') !== '0' });
const configuredOrigin = params.get('parentOrigin') || location.origin;
const parentOrigin = new URL(configuredOrigin).origin;
if (parentOrigin !== configuredOrigin || !/^https?:/.test(parentOrigin)) throw new Error('parentOrigin 必须是精确的 HTTP / HTTPS 来源。');
const describe = model => ({ name: model.name, kind: model.kind, format: model.format, warnings: model.warnings });
const notifyNative = message => window.webkit?.messageHandlers?.officePreview?.postMessage(message);
viewer.addEventListener('load', event => notifyNative({ type: 'office-preview:loaded', ...describe(event.detail) }));
viewer.addEventListener('error', event => notifyNative({ type: 'office-preview:error', error: event.detail.message }));

window.addEventListener('message', async event => {
  if (event.source !== window.parent || event.origin !== parentOrigin) return;
  if (event.data?.type === 'office-preview:setZoom' || event.data?.type === 'office-preview:fit') {
    const { type, zoom, requestId } = event.data;
    try {
      if (type === 'office-preview:setZoom') {
        if (typeof zoom !== 'number' || !Number.isFinite(zoom)) throw new TypeError('缩放值必须是有限数字。');
        viewer.setZoom(zoom);
      } else viewer.fit();
      window.parent.postMessage({ type: 'office-preview:zoom', requestId, zoom: viewer.zoom }, parentOrigin);
    } catch (error) {
      window.parent.postMessage({ type: 'office-preview:error', requestId, error: error.message }, parentOrigin);
    }
    return;
  }
  if (event.data?.type !== 'office-preview:open') return;
  const { requestId, buffer, name, options } = event.data;
  try {
    if (!(buffer instanceof ArrayBuffer)) throw new Error('嵌入消息需要 ArrayBuffer。');
    const model = await viewer.open(buffer, { ...options, name });
    window.parent.postMessage({ type: 'office-preview:loaded', requestId, name: model.name, format: model.format, warnings: model.warnings }, parentOrigin);
  } catch (error) {
    if (error.name !== 'AbortError') window.parent.postMessage({ type: 'office-preview:error', requestId, error: error.message }, parentOrigin);
  }
});
// A desktop WebView can call this object with a byte buffer or a base64 string.
window.officePreview = {
  open: (buffer, name, options = {}) => viewer.open(buffer, { ...options, name }).then(describe),
  openBase64: async (base64, name) => {
    if (typeof base64 !== 'string') throw new TypeError('文件内容必须是 base64 字符串。');
    if (base64.length > Math.ceil(32 * 1024 * 1024 / 3) * 4) throw new Error('文件超过首版 32MB 限制。');
    const binary = atob(base64);
    return window.officePreview.open(Uint8Array.from(binary, ch => ch.charCodeAt(0)), name);
  },
  getCapabilities: () => viewer.getCapabilities(),
  setZoom: value => viewer.setZoom(value),
  fit: () => viewer.fit(),
  destroy: () => viewer.destroy()
};
window.parent.postMessage({ type: 'office-preview:ready' }, parentOrigin);
notifyNative({ type: 'office-preview:ready' });
