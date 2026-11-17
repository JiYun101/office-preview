// Real local File -> Worker -> SDK. Reports only geometry, never document text.
import {spawn} from 'node:child_process';
import {access,copyFile,mkdir,mkdtemp,readFile,rm,writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {basename,dirname,join,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {startServer} from './serve.mjs';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..'),args=process.argv.slice(2);
const option=name=>{const at=args.indexOf(name);return at<0?undefined:args[at+1];};
const file=option('--file'),tag=option('--tag')||'current',source=option('--source'),reference=option('--reference'),before=option('--before');
if(!file||!/^[a-z0-9-]+$/i.test(tag))throw new Error('Usage: node scripts/word-report-check.mjs --file <docx> [--tag baseline] [--source /artifacts/word-report-baseline-src/index.js] [--reference <screenshot>] [--before <inspection.json>]');
await Promise.all([file,...reference?[reference]:[]].map(path=>access(resolve(path))));
const candidates=[process.env.OFFICE_PREVIEW_BROWSER,'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe','C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/usr/bin/chromium'].filter(Boolean);
let executable;for(const path of candidates){try{await access(path);executable=path;break;}catch{}}
if(!executable)throw new Error('No Chromium browser found. Set OFFICE_PREVIEW_BROWSER.');
const artifacts=join(root,'artifacts'),prefix=`word-report-${tag}`;await mkdir(artifacts,{recursive:true});
const profile=await mkdtemp(join(tmpdir(),'office-word-report-')),server=await startServer({port:0}),origin=`http://127.0.0.1:${server.address().port}`;
const browser=spawn(executable,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--remote-debugging-port=0',`--user-data-dir=${profile}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay=ms=>new Promise(resolve=>setTimeout(resolve,ms));let ws;
try{
  let address;const deadline=Date.now()+20000;
  while(Date.now()<deadline){try{address=(await readFile(join(profile,'DevToolsActivePort'),'utf8')).trim().split('\n');break;}catch{await delay(100);}}
  if(!address)throw new Error('Browser startup timed out.');
  ws=new WebSocket(`ws://127.0.0.1:${address[0]}${address[1]}`);await new Promise((resolve,reject)=>{ws.addEventListener('open',resolve,{once:true});ws.addEventListener('error',reject,{once:true});});
  let sequence=0;const pending=new Map();
  ws.addEventListener('message',event=>{const message=JSON.parse(event.data);if(pending.has(message.id)){const task=pending.get(message.id);pending.delete(message.id);clearTimeout(task.timer);message.error?task.reject(new Error(message.error.message)):task.resolve(message.result);}});
  const call=(method,params={},sessionId)=>new Promise((resolve,reject)=>{const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP timeout: '+method));},30000);pending.set(id,{resolve,reject,timer});ws.send(JSON.stringify({id,method,params,...sessionId?{sessionId}:{}}));});
  const target=await call('Target.createTarget',{url:'about:blank'}),{sessionId}=await call('Target.attachToTarget',{targetId:target.targetId,flatten:true}),send=(method,params)=>call(method,params,sessionId);
  await send('Page.enable');await send('Runtime.enable');await send('Emulation.setDeviceMetricsOverride',{width:1400,height:1700,deviceScaleFactor:1,mobile:false});
  const evaluate=async expression=>{const response=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(response.exceptionDetails)throw new Error(response.exceptionDetails.exception?.description||response.exceptionDetails.text);return response.result.value;};
  await send('Page.navigate',{url:origin+'/tests/word-report-browser.html?inspect=1'+(source?'&source='+encodeURIComponent(source):'')});
  const startup=Date.now()+20000;while(!await evaluate('!!window.wordReport')){if(Date.now()>startup)throw new Error('QA page startup timed out.');await delay(100);}
  const document=await send('DOM.getDocument'),input=await send('DOM.querySelector',{nodeId:document.root.nodeId,selector:'#fixture-files'});await send('DOM.setFileInputFiles',{nodeId:input.nodeId,files:[resolve(file)]});
  const metadata=await evaluate('window.wordReport.open()'),box=await evaluate('window.wordReport.page()');
  if(box.width>8192||box.height>8192)throw new Error('Paper dimensions exceed capture bounds.');
  await send('Page.bringToFront');await delay(100);
  const capture=await send('Page.captureScreenshot',{format:'png',clip:{x:box.x,y:box.y,width:box.width,height:box.height,scale:1},captureBeyondViewport:false}),image=prefix+'-page-1.png';
  await writeFile(join(artifacts,image),Buffer.from(capture.data,'base64'));
  const report={tag,method:'Local File -> production Worker -> SDK; 100% CSS zoom; only geometry/style metadata stored. Independent native measurements are development-only oracles, never the production parsing/rendering path.',...metadata,image};
  if(reference){report.referenceImage=prefix+'-reference.png';await copyFile(resolve(reference),join(artifacts,report.referenceImage));}
  if(before){const baseline=JSON.parse(await readFile(resolve(before),'utf8'));report.before={image:baseline.image,paper:baseline.paper};}
  await writeFile(join(artifacts,prefix+'-inspection.json'),JSON.stringify(report,null,2));
  const data=JSON.stringify(report).replaceAll('<','\\u003c');
  const html=`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Word 表格与浮动图片对照</title><style>body{font:14px/1.7 system-ui;margin:24px;background:#edf1f5;color:#25334a}.reference{max-width:100%;height:auto}.row{display:flex;gap:20px;align-items:flex-start}.row figure{margin:0;flex:1;min-width:0}.row img{width:100%;height:auto;display:block}pre{background:white;padding:16px;white-space:pre-wrap}h1{font-size:22px}</style><h1>Word 表格与浮动图片对照</h1><p>参考截图同时包含旧预览与 WPS，右侧原件纸页被窗口截断。保留完整截图供人工核对，不将其拉伸为完整纸页；下方 SDK 纸页保持各自原始宽高比例。原生独立样本测量仅作为开发核验依据，实际解析和渲染始终经自研 Worker / SDK。</p><div id="reference"></div><div class="row"><figure id="before" hidden><figcaption>修复前 · 100%</figcaption><img></figure><figure><figcaption>当前 · 100%</figcaption><img id="current"></figure></div><pre id="stats"></pre><script type="application/json" id="data">${data}</script><script>const data=JSON.parse(document.querySelector('#data').textContent);if(data.referenceImage){const image=document.createElement('img');image.className='reference';image.src=data.referenceImage;document.querySelector('#reference').append(image);}if(data.before){document.querySelector('#before').hidden=false;document.querySelector('#before img').src=data.before.image;}document.querySelector('#current').src=data.image;document.querySelector('#stats').textContent=JSON.stringify(data,null,2);</script></html>`;
  await writeFile(join(artifacts,prefix+'-gallery.html'),html);
  console.log(JSON.stringify({tag,pageCount:report.pageCount,tableCount:report.tables.length,table:report.tables[0]?.box,images:report.images.map(image=>({box:image.box,naturalWidth:image.naturalWidth,naturalHeight:image.naturalHeight})),gallery:`artifacts/${prefix}-gallery.html`}));
  if(args.includes('--self-test')){
    await send('Page.navigate',{url:origin+'/tests/word-report-browser.html'});let results;const deadline=Date.now()+30000;
    while(Date.now()<deadline){results=await evaluate('window.officeBrowserTestResults');if(results)break;await delay(100);}
    if(!results)throw new Error('Word report browser checks timed out.');console.log(JSON.stringify(results));if(!results.passed)throw new Error('Word report browser checks failed.');
  }
  await call('Browser.close');
}finally{
  ws?.close();if(browser.exitCode===null){await Promise.race([new Promise(resolve=>browser.once('exit',resolve)),delay(1500)]);if(browser.exitCode===null)browser.kill();}
  await new Promise(resolve=>server.close(resolve));
  if(dirname(resolve(profile))===resolve(tmpdir())&&basename(profile).startsWith('office-word-report-'))await rm(profile,{recursive:true,force:true,maxRetries:8,retryDelay:200}).catch(()=>{});
}
