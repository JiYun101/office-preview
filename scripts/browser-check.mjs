// Browser integration checks over the standard Chrome DevTools Protocol.
// No browser automation dependency is installed or shipped with the SDK.
import { spawn } from 'node:child_process';
import { access, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { startServer } from './serve.mjs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const delay = ms => new Promise(accept => setTimeout(accept, ms));
const candidates = [process.env.OFFICE_PREVIEW_BROWSER, 'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe', 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe', '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome', '/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge', '/Applications/Chromium.app/Contents/MacOS/Chromium', '/usr/bin/chromium', '/usr/bin/google-chrome'].filter(Boolean);
let executable;
for (const candidate of candidates) { try { await access(candidate); executable = candidate; break; } catch {} }
if (!executable) throw new Error('未发现 Chromium / Edge。可通过 OFFICE_PREVIEW_BROWSER 指定可执行文件。');
const profile = await mkdtemp(join(tmpdir(), 'office-preview-check-'));
const server = await startServer({ port: 0 });
const origin = `http://127.0.0.1:${server.address().port}`;
const child = spawn(executable, ['--headless=new', '--disable-gpu', '--no-first-run', '--no-default-browser-check', '--remote-debugging-port=0', `--user-data-dir=${profile}`, '--window-size=1440,1100', 'about:blank'], { windowsHide: true, stdio: 'ignore' });
let ws;
let failed = false;
try {
  let portInfo;
  const bootDeadline = Date.now() + 20000;
  while (Date.now() < bootDeadline) { try { portInfo = (await readFile(join(profile, 'DevToolsActivePort'), 'utf8')).trim().split('\n'); break; } catch { await delay(150); } }
  if (!portInfo) throw new Error('浏览器调试接口启动超时。');
  ws = new WebSocket(`ws://127.0.0.1:${portInfo[0]}${portInfo[1]}`);
  await new Promise((accept, reject) => { ws.addEventListener('open', accept, { once: true }); ws.addEventListener('error', reject, { once: true }); });
  let sequence = 0;
  const pending = new Map();
  const browserErrors = [];
  ws.addEventListener('message', event => {
    const message = JSON.parse(event.data);
    if (message.method === 'Runtime.exceptionThrown') browserErrors.push(message.params.exceptionDetails.exception?.description || message.params.exceptionDetails.text);
    if (message.id && pending.has(message.id)) {
      const { accept, reject, timer } = pending.get(message.id); pending.delete(message.id); clearTimeout(timer);
      message.error ? reject(new Error(message.error.message)) : accept(message.result);
    }
  });
  const call = (method, params = {}, sessionId) => new Promise((accept, reject) => {
    const id = ++sequence;
    const timer = setTimeout(() => { pending.delete(id); reject(new Error(`CDP 请求超时：${method}`)); }, 20000);
    pending.set(id, { accept, reject, timer }); ws.send(JSON.stringify({ id, method, params, ...(sessionId ? { sessionId } : {}) }));
  });
  const target = await call('Target.createTarget', { url: 'about:blank' });
  const { sessionId } = await call('Target.attachToTarget', { targetId: target.targetId, flatten: true });
  const send = (method, params) => call(method, params, sessionId);
  await send('Page.enable');
  await send('Runtime.enable');
  const evaluate = async expression => {
    const result = await send('Runtime.evaluate', { expression, returnByValue: true, awaitPromise: true });
    if (result.exceptionDetails) throw new Error(result.exceptionDetails.text + ': ' + (result.exceptionDetails.exception?.description || ''));
    return result.result.value;
  };
  const waitFor = async expression => {
    const deadline = Date.now() + 30000;
    while (Date.now() < deadline) { const value = await evaluate(expression); if (value) return value; await delay(150); }
    const diagnostic = await evaluate("({url:location.href,title:document.title,results:document.querySelector('#results')?.textContent?.slice(0,2000)})");
    throw new Error(`浏览器检查超时：${expression}\n${JSON.stringify(diagnostic)}\n${browserErrors.slice(-5).join('\n')}`);
  };
  for (const page of ['/tests/browser.html', '/tests/sdk-browser.html', '/tests/formats-browser.html', '/tests/legacy-style-browser.html', '/tests/slide-render-browser.html', '/tests/word-browser.html', '/tests/word-report-browser.html', '/tests/pdf-scroll-browser.html', '/tests/pdf-quality-browser.html', '/tests/pdf-paint-browser.html', '/tests/sheet-scroll-browser.html', '/tests/embed-minimal-browser.html', '/tests/ppt-scroll-browser.html']) {
    await send('Page.navigate', { url: origin + page });
    if (page === '/tests/formats-browser.html') {
      await waitFor("location.pathname==='/tests/formats-browser.html' && !!document.querySelector('#fixture-files')");
      const documentNode = await send('DOM.getDocument');
      const input = await send('DOM.querySelector', { nodeId: documentNode.root.nodeId, selector: '#fixture-files' });
      // Use actual local File objects, as the native host does. Document HTTP
      // downloads can be rewritten by system proxies or inline PDF viewers.
      await send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: ['sample.pdf', 'sample.eml', 'sample.msg'].map(name => join(root, 'fixtures', name)) });
    }
    if (page === '/tests/legacy-style-browser.html') {
      await waitFor("location.pathname==='/tests/legacy-style-browser.html' && !!document.querySelector('#fixture-files')");
      const documentNode = await send('DOM.getDocument');
      const input = await send('DOM.querySelector', { nodeId: documentNode.root.nodeId, selector: '#fixture-files' });
      await send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: ['styled.xls', 'hidden.xls', 'styled.ppt'].map(name => join(root, 'fixtures', name)) });
    }
    if (page === '/tests/word-browser.html') {
      await waitFor("location.pathname==='/tests/word-browser.html' && !!document.querySelector('#fixture-files')");
      const documentNode = await send('DOM.getDocument');
      const input = await send('DOM.querySelector', { nodeId: documentNode.root.nodeId, selector: '#fixture-files' });
      await send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: ['styled.docx', 'page-areas.docx', 'page-areas.doc'].map(name => join(root, 'fixtures', name)) });
    }
    if (page === '/tests/pdf-scroll-browser.html' || page === '/tests/pdf-quality-browser.html') {
      await waitFor(`location.pathname===${JSON.stringify(page)} && !!document.querySelector('#fixture-files')`);
      const documentNode = await send('DOM.getDocument');
      const input = await send('DOM.querySelector', { nodeId: documentNode.root.nodeId, selector: '#fixture-files' });
      await send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [join(root, 'fixtures', 'sample.pdf')] });
    }
    const result = await waitFor(`location.pathname === ${JSON.stringify(page)} && document.documentElement.dataset.status ? window.officeBrowserTestResults : null`);
    const entries = result.results || [];
    const passed = entries.filter(item => item.ok ?? item.passed).length;
    const total = result.total ?? entries.length;
    console.log(`${page}: ${passed}/${total} passed`);
    for (const item of entries) if (!(item.ok ?? item.passed)) { failed = true; console.error(`  FAIL ${item.name}: ${item.error || item.message || ''}`); }
    if (passed !== total) failed = true;
  }
  if (process.argv.includes('--screenshot')) {
    await send('Emulation.setDeviceMetricsOverride', { width: 1440, height: 1100, deviceScaleFactor: 1, mobile: false });
    await send('Page.navigate', { url: origin + '/' });
    await waitFor("typeof document.querySelector('[data-sample=docx]')?.onclick === 'function'");
    await evaluate("document.querySelector('[data-sample=docx]').click()");
    await waitFor("document.querySelector('#load-status').textContent.includes('已读取')");
    await mkdir(join(root, 'artifacts'), { recursive: true });
    const screenshot = await send('Page.captureScreenshot', { format: 'png' });
    await writeFile(join(root, 'artifacts', 'demo-desktop.png'), Buffer.from(screenshot.data, 'base64'));
    for (const [format, filename] of [['pdf', 'demo-pdf.png'], ['eml', 'demo-email.png']]) {
      await evaluate(`document.querySelector('[data-sample=${format}]').click()`);
      await waitFor(`document.querySelector('#load-status').textContent.includes(${JSON.stringify(format.toUpperCase())}) && document.querySelector('#load-status').textContent.includes('已读取')`);
      const capture = await send('Page.captureScreenshot', { format: 'png' });
      await writeFile(join(root, 'artifacts', filename), Buffer.from(capture.data, 'base64'));
    }
    await evaluate("document.querySelector('[data-sample=docx]').click()");
    await waitFor("document.querySelector('#load-status').textContent.includes('DOCX') && document.querySelector('#load-status').textContent.includes('已读取')");
    await send('Emulation.setDeviceMetricsOverride', { width: 390, height: 844, deviceScaleFactor: 1, mobile: true });
    await delay(200);
    const mobile = await send('Page.captureScreenshot', { format: 'png' });
    await writeFile(join(root, 'artifacts', 'demo-mobile.png'), Buffer.from(mobile.data, 'base64'));
    console.log('Screenshots: artifacts/demo-desktop.png, demo-mobile.png, demo-pdf.png, demo-email.png');
  }
  const inspectAt = process.argv.indexOf('--inspect-files');
  if (inspectAt >= 0) {
    const paths = [];
    for (const argument of process.argv.slice(inspectAt + 1)) { if (argument.startsWith('--')) break; paths.push(resolve(argument)); }
    const tagAt = process.argv.indexOf('--inspect-tag');
    const tag = tagAt < 0 ? '' : process.argv[tagAt + 1];
    if (tag && !/^[a-z0-9-]+$/i.test(tag)) throw new Error('Invalid inspection tag.');
    const prefix = tag ? `user-preview-${tag}` : 'user-preview';
    await mkdir(join(root, 'artifacts'), { recursive: true });
    await send('Emulation.setDeviceMetricsOverride', { width: 1600, height: 1300, deviceScaleFactor: 1, mobile: false });
    await send('Page.navigate', { url: origin + '/examples/iframe.html' });
    await waitFor("location.pathname==='/examples/iframe.html' && document.querySelector('#file')?.disabled===false");
    const viewer = "document.querySelector('#preview').contentDocument.querySelector('#viewer').firstElementChild.shadowRoot";
    const inspection = [];
    for (let index = 0; index < paths.length; index++) {
      const name = basename(paths[index]);
      const documentNode = await send('DOM.getDocument');
      const input = await send('DOM.querySelector', { nodeId: documentNode.root.nodeId, selector: '#file' });
      await send('DOM.setFileInputFiles', { nodeId: input.nodeId, files: [paths[index]] });
      await waitFor(`document.querySelector('#status').textContent===${JSON.stringify('已读取 ' + name)}`);
      const slotCount = await evaluate(`(${viewer}).querySelectorAll('.op-slide-page').length`);
      const pages = [];
      for (let page = 0; page < 500; page++) {
        const scope = slotCount ? `(${viewer}).querySelectorAll('.op-slide-page')[${page}]` : `(${viewer})`;
        if (slotCount) {
          await evaluate(`(${scope}).scrollIntoView({block:'start'})`);
          await waitFor(`(${scope}).dataset.ready==='true'`);
        }
        await evaluate(`Promise.all([...(${scope}).querySelectorAll('img')].map(image=>{image.loading='eager';return image.decode().catch(()=>{})}))`);
        const stats = await evaluate(`(()=>{const r=${scope}; const canvas=r.querySelector('.op-slide'); const table=r.querySelector('.op-grid');const cell=table?.querySelector('tbody td');const style=cell?getComputedStyle(cell):null;return {slide:!!canvas,background:canvas?getComputedStyle(canvas).backgroundColor:null,shapes:r.querySelectorAll('.op-slide-shape').length,textBoxes:r.querySelectorAll('.op-slide-text').length,images:[...r.querySelectorAll('img')].map(image=>({width:image.naturalWidth,height:image.naturalHeight})),table:table?{visibleRows:table.querySelectorAll('tbody tr').length,firstRows:[...table.querySelectorAll('tbody tr')].slice(0,3).map(row=>row.getBoundingClientRect().height),columns:[...table.querySelectorAll('thead th')].slice(1,7).map(cell=>cell.getBoundingClientRect().width),titleStyle:{fontFamily:style.fontFamily,fontSize:style.fontSize,textAlign:style.textAlign,verticalAlign:style.verticalAlign,borderBottom:style.borderBottom}}:null,error:r.querySelector('.op-error')?.textContent||null}})()`);
        if (stats.error || stats.images.some(image => !image.width)) { failed = true; console.error(`User preview failed: ${name}, page ${page + 1}`); }
        const next = page + 1 < slotCount;
        if (page === 0 || page === 4 || !next) {
          const clip = await evaluate("(()=>{const box=document.querySelector('#preview').getBoundingClientRect();return {x:box.x,y:box.y,width:box.width,height:box.height,scale:1}})()");
          const file = `${prefix}-${index + 1}-page-${String(page + 1).padStart(2, '0')}.png`;
          const capture = await send('Page.captureScreenshot', { format: 'png', clip, captureBeyondViewport: true });
          await writeFile(join(root, 'artifacts', file), Buffer.from(capture.data, 'base64'));
          stats.screenshot = file;
        }
        pages.push(stats);
        if (!next) break;
      }
      inspection.push({ name, pages });
      console.log(`Inspected ${name}: ${pages.length} page(s)`);
    }
    await writeFile(join(root, 'artifacts', `${prefix}-inspection.json`), JSON.stringify(inspection, null, 2));
  }
  await call('Browser.close');
} finally {
  ws?.close();
  await new Promise(accept => server.close(accept));
  if (child.exitCode === null) { await Promise.race([new Promise(accept => child.once('exit', accept)), delay(1500)]); if (child.exitCode === null) child.kill(); }
  // Delete only the temporary profile created by this check, never a user profile.
  if (dirname(resolve(profile)) === resolve(tmpdir()) && basename(profile).startsWith('office-preview-check-')) await rm(profile, { recursive: true, force: true, maxRetries: 8, retryDelay: 200 }).catch(() => {});
}
if (failed) process.exitCode = 1;
