import { parseDocument } from './parse.js';
self.onmessage = async event => {
  try {
    const model = await parseDocument(event.data.buffer, event.data.options);
    self.postMessage({ ok: true, model });
  } catch (error) {
    self.postMessage({ ok: false, error: error.message || String(error) });
  }
};
