import { OfficePreview } from '../src/index.js';

const viewer = new OfficePreview(document.querySelector('#viewer'));
const input = document.querySelector('#file-input');
const drop = document.querySelector('#drop-zone');
const status = document.querySelector('#load-status');
let currentFile;
let request = 0;
const options = () => ({ encoding: document.querySelector('#encoding').value, delimiter: document.querySelector('#delimiter').value.replace('tab', '\t') });
async function open(file) {
  const id = ++request;
  currentFile = file; status.textContent = '正在解析…';
  try {
    const model = await viewer.open(file, options());
    if (id === request) { status.textContent = `${model.format.toUpperCase()} · 已读取${model.warnings?.length ? `（${model.warnings.length}项兼容提示）` : ''}`; status.title = (model.warnings || []).join('\n'); }
  } catch (error) { if (id === request && error.name !== 'AbortError') status.textContent = error.message; }
}
document.querySelector('#open-file').onclick = () => input.click();
input.onchange = () => { if (input.files[0]) open(input.files[0]); input.value = ''; };
for (const event of ['dragenter', 'dragover']) drop.addEventListener(event, event => { event.preventDefault(); drop.classList.add('dragover'); });
drop.addEventListener('dragleave', () => drop.classList.remove('dragover'));
drop.addEventListener('drop', event => { event.preventDefault(); drop.classList.remove('dragover'); if (event.dataTransfer.files[0]) open(event.dataTransfer.files[0]); });
async function sample(format) {
  const id = ++request;
  try {
    status.textContent = '正在读取开发样本…';
    const { samples } = await import('../fixtures/samples.js');
    if (!samples[format]) throw new Error('开发样本未生成，请运行 npm run fixtures。');
    const buffer = Uint8Array.from(atob(samples[format].base64), character => character.charCodeAt(0));
    if (id === request) await open(new File([buffer], `开发样本.${format}`));
  } catch (error) { if (id === request) status.textContent = error.message; }
}
for (const button of document.querySelectorAll('[data-sample]')) button.onclick = () => sample(button.dataset.sample);
document.querySelector('#legacy-sample').onchange = event => { if (event.target.value) sample(event.target.value); event.target.value = ''; };
for (const selector of ['#encoding', '#delimiter']) document.querySelector(selector).onchange = () => { if (currentFile) open(currentFile); };
window.addEventListener('pagehide', () => viewer.destroy());
