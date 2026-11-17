// Native local File -> production Worker -> SDK. No external PDF engine.
import {spawn} from 'node:child_process';
import {access,cp,mkdir,mkdtemp,readFile,rm,writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {basename,dirname,join,resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {startServer} from './serve.mjs';

const root=resolve(dirname(fileURLToPath(import.meta.url)),'..'),args=process.argv.slice(2);
const option=name=>{const at=args.indexOf(name);return at<0?null:args[at+1];};
const file=option('--file')||(args.includes('--self-test')?join(root,'fixtures','sample.pdf'):null),tag=option('--tag')||'current',reference=option('--reference');
if(!file||!/^[a-z0-9-]+$/i.test(tag))throw new Error('Usage: node scripts/pdf-preview-check.mjs --file <pdf> [--reference <image>] [--tag baseline] [--snapshot-source | --source <snapshot index URL>] [--before <inspection JSON>] [--quick] [--self-test]');
await Promise.all([file,...reference?[reference]:[]].map(path=>access(resolve(path))));
const artifacts=join(root,'artifacts'),prefix=`pdf-preview-${tag}`;await mkdir(artifacts,{recursive:true});
let source=option('--source')?`?source=${encodeURIComponent(option('--source'))}`:'';
if(args.includes('--snapshot-source')){const destination=join(artifacts,`${prefix}-src`);await cp(join(root,'src'),destination,{recursive:true});source=`?source=/artifacts/${prefix}-src/index.js`;}
const candidates=[process.env.OFFICE_PREVIEW_BROWSER,'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe','C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge','/usr/bin/chromium','/usr/bin/google-chrome'].filter(Boolean);
let executable;for(const candidate of candidates){try{await access(candidate);executable=candidate;break;}catch{}}
if(!executable)throw new Error('Chromium / Edge not found. Set OFFICE_PREVIEW_BROWSER.');
const profile=await mkdtemp(join(tmpdir(),'office-pdf-preview-')),server=await startServer({port:0});
const origin=`http://127.0.0.1:${server.address().port}`;
const browser=spawn(executable,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--remote-debugging-port=0',`--user-data-dir=${profile}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay=ms=>new Promise(resolve=>setTimeout(resolve,ms));let ws;
try{
  let address;const deadline=Date.now()+20000;
  while(Date.now()<deadline){try{address=(await readFile(join(profile,'DevToolsActivePort'),'utf8')).trim().split('\n');break;}catch{await delay(100);}}
  if(!address)throw new Error('Browser debug endpoint startup timed out.');
  ws=new WebSocket(`ws://127.0.0.1:${address[0]}${address[1]}`);
  await new Promise((resolve,reject)=>{ws.addEventListener('open',resolve,{once:true});ws.addEventListener('error',reject,{once:true});});
  let sequence=0;const pending=new Map(),errors=[];
  ws.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.exception?.description||message.params.exceptionDetails.text);if(pending.has(message.id)){const task=pending.get(message.id);pending.delete(message.id);clearTimeout(task.timer);message.error?task.reject(new Error(message.error.message)):task.resolve(message.result);}});
  const call=(method,params={},sessionId)=>new Promise((resolve,reject)=>{const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP timeout: '+method));},60000);pending.set(id,{resolve,reject,timer});ws.send(JSON.stringify({id,method,params,...sessionId?{sessionId}:{}}));});
  const target=await call('Target.createTarget',{url:'about:blank'}),{sessionId}=await call('Target.attachToTarget',{targetId:target.targetId,flatten:true});
  const send=(method,params)=>call(method,params,sessionId);
  await send('Page.enable');await send('Runtime.enable');
  const evaluate=async expression=>{const result=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
  const choose=async(selector,paths)=>{const document=await send('DOM.getDocument'),node=await send('DOM.querySelector',{nodeId:document.root.nodeId,selector});await send('DOM.setFileInputFiles',{nodeId:node.nodeId,files:paths.map(path=>resolve(path))});};
  const report={tag,method:'Native File -> Worker -> SDK. Actual CSS bounds and backing pixels after completed redraw; private file contents are not logged.',samples:[]};
  const matrix=args.includes('--quick')?[[1,1.5],[2,1.5]]:[[1,1],[1,1.5],[1,2],[2,1],[2,1.5],[2,2]];
  for(const [dpr,zoom] of matrix){
    await send('Emulation.setDeviceMetricsOverride',{width:1800,height:2600,deviceScaleFactor:dpr,mobile:false});
    await send('Page.navigate',{url:origin+'/tests/pdf-quality-browser.html?inspect=1'+source.replace(/^\?/,'&')});
    const deadline=Date.now()+20000;while(!await evaluate('!!window.pdfPreview')){if(Date.now()>deadline)throw new Error('QA page startup failed: '+errors.slice(-3).join('\n'));await delay(100);}
    await choose('#fixture-files',[file]);
    const metadata=await evaluate(`window.pdfPreview.open(0,${zoom})`),box=metadata.paper;
    const image=`${prefix}-dpr-${dpr}-zoom-${Math.round(zoom*100)}.png`;
    if(box.width>8192||box.height>8192||box.width*box.height*dpr*dpr>48000000)throw new Error('Capture dimensions exceed QA capacity.');
    await send('Page.bringToFront');
    let capture,rasterSample;
    for(let attempt=0;attempt<4;attempt++){
      if(attempt)await delay(150);
      capture=await send('Page.captureScreenshot',{format:'png',clip:{x:Math.max(0,box.x),y:Math.max(0,box.y),width:Math.min(box.width,1800),height:Math.min(box.height,2600-box.y),scale:1},captureBeyondViewport:false});
      rasterSample=await evaluate(`window.pdfPreview.rasterSample(${JSON.stringify(capture.data)})`);
      if(!metadata.textOperationCount||rasterSample.sampledInkRatio>.002)break;
    }
    if(metadata.textOperationCount&&rasterSample.sampledInkRatio<=.002)throw new Error('Captured compositor surface is blank despite text operations.');
    await writeFile(join(artifacts,image),Buffer.from(capture.data,'base64'));
    report.samples.push({...metadata,image,rasterSample});
    console.log(JSON.stringify({tag,dpr,zoom,paper:metadata.paper,canvas:metadata.canvas,pixelDensity:metadata.pixelDensity,fonts:metadata.fonts,warnings:metadata.warnings}));
  }
  if(args.includes('--dpr-transition')){
    await send('Emulation.setDeviceMetricsOverride',{width:1800,height:2600,deviceScaleFactor:1,mobile:false});
    await send('Page.navigate',{url:origin+'/tests/pdf-quality-browser.html?inspect=1'});
    const deadline=Date.now()+20000;while(!await evaluate('!!window.pdfPreview')){if(Date.now()>deadline)throw new Error('DPR transition page failed to start.');await delay(100);}
    await choose('#fixture-files',[file]);const before=await evaluate('window.pdfPreview.open(0,1)');
    await evaluate("window.pdfDprEvents=[];window.addEventListener('resize',()=>window.pdfDprEvents.push({type:'resize',dpr:devicePixelRatio}));window.pdfDprQuery=matchMedia('(resolution: 1dppx)');window.pdfDprQuery.addEventListener('change',event=>window.pdfDprEvents.push({type:'media',matches:event.matches,dpr:devicePixelRatio}));true");
    await send('Emulation.setDeviceMetricsOverride',{width:1800,height:2600,deviceScaleFactor:3,mobile:false});
    let after,dprOnly;const redraw=Date.now()+3000;
    while(Date.now()<redraw){after=await evaluate('window.pdfPreview.metadata()');if(after.dpr===3&&after.zoom===1&&after.paperState==='ready'&&after.canvas.ready&&after.pixelDensity.x>2.98&&after.pixelDensity.y>2.98)break;await delay(50);}
    if(!(after.pixelDensity.x>2.98&&after.pixelDensity.y>2.98&&after.zoom===1)){
      const diagnostic=await evaluate("({events:[...window.pdfDprEvents],dpr:devicePixelRatio,media1:matchMedia('(resolution: 1dppx)').matches,media3:matchMedia('(resolution: 3dppx)').matches,visualViewport:{scale:visualViewport.scale,width:visualViewport.width,height:visualViewport.height}})");
      dprOnly={passed:false,reason:'Headless CDP changed density queries without delivering resolution or resize events.',after,diagnostic};
      // Headless CDP can update devicePixelRatio without delivering resolution
      // or resize events. Change host width by one CSS pixel to deliver the
      // ordinary resize event; keep the same File, instance and SDK zoom.
      await send('Emulation.setDeviceMetricsOverride',{width:1801,height:2600,deviceScaleFactor:3,mobile:false});
      const resized=Date.now()+10000;
      while(Date.now()<resized){after=await evaluate('window.pdfPreview.metadata()');if(after.paperState==='ready'&&after.canvas.ready&&after.pixelDensity.x>2.98&&after.pixelDensity.y>2.98)break;await delay(50);}
      if(!(after.pixelDensity.x>2.98&&after.pixelDensity.y>2.98&&after.zoom===1)){
        report.densityTransition={passed:false,before,after,dprOnly,resizeEvents:await evaluate('window.pdfDprEvents')};await writeFile(join(artifacts,`${prefix}-inspection.json`),JSON.stringify(report,null,2));
        throw new Error('DPR redraw failed after host resize: '+JSON.stringify(report.densityTransition));
      }
    }
    if(after.canvas.width*after.canvas.height>16000000||Math.max(after.canvas.width,after.canvas.height)>8192)throw new Error('DPR redraw exceeded page pixel limits.');
    report.densityTransition={passed:true,method:'Same opened File and SDK instance; DPR 1 -> 3 via screen metrics, no navigation or setZoom between measurements.'+(dprOnly?' CDP delivered no density-only events; host width +1 CSS pixel delivered the resize event.':''),before,after,eventDrivenRedraw:{passed:true,trigger:dprOnly?'host width +1 CSS pixel / resize event':'resolution change event',canvasPixels:after.canvas.width*after.canvas.height},...dprOnly?{dprOnly,resizeEvents:await evaluate('window.pdfDprEvents')}:{}};
    console.log(JSON.stringify({densityTransition:{before:{dpr:before.dpr,zoom:before.zoom,canvas:before.canvas,pixelDensity:before.pixelDensity},after:{dpr:after.dpr,zoom:after.zoom,canvas:after.canvas,pixelDensity:after.pixelDensity}}}));
  }
  if(args.includes('--self-test')){
    await send('Page.navigate',{url:origin+'/tests/pdf-quality-browser.html'});
    const deadline=Date.now()+20000;while(!await evaluate('!!window.pdfPreview')){if(Date.now()>deadline)throw new Error('QA tests failed to start.');await delay(100);}
    await choose('#fixture-files',[file]);
    let checks;const ready=Date.now()+60000;while(Date.now()<ready){checks=await evaluate('window.officeBrowserTestResults');if(checks)break;await delay(100);}
    if(!checks)throw new Error('PDF quality checks timed out.');report.checks=checks;
    console.log(`PDF quality checks: ${checks.results.filter(check=>check.ok).length}/${checks.total}`);
    if(!checks.passed)throw new Error(JSON.stringify(checks.results.filter(check=>!check.ok)));
  }
  if(reference){await choose('#reference',[reference]);const image=await evaluate('window.pdfPreview.reference()');report.reference={width:image.width,height:image.height,crop:image.crop,image:`${prefix}-reference-paper.png`,rasterSample:await evaluate(`window.pdfPreview.rasterSample(${JSON.stringify(image.png)})`)};await writeFile(join(artifacts,report.reference.image),Buffer.from(image.png,'base64'));}
  if(option('--before')){const before=JSON.parse(await readFile(resolve(option('--before')),'utf8'));report.before={tag:before.tag,samples:before.samples};}
  const reportName=`${prefix}-inspection.json`;await writeFile(join(artifacts,reportName),JSON.stringify(report,null,2));
  const data=JSON.stringify(report).replaceAll('<','\\u003c');
  const html=`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>PDF 纸页质量检查</title><style>body{margin:24px;background:#edf1f5;color:#25334a;font:14px/1.7 system-ui}h1{font-size:22px}.row{display:flex;gap:20px;align-items:flex-start}.row figure{margin:0;flex:1;min-width:0}.row img{width:100%;height:auto;display:block}pre{background:white;padding:16px;white-space:pre-wrap}nav{display:flex;gap:8px;margin:16px 0}button{padding:6px 12px}</style><h1>PDF 纸页对照</h1><p>截图经过真实 File → Worker → SDK。参考图只裁掉应用背景，按纸页宽度等比例显示。截图包含不同缩放与字体光栅化差异，本页不据此声称逐像素一致。</p><nav id="samples"></nav><div class="row"><figure id="reference"></figure><figure id="before" hidden><figcaption>修复前</figcaption><img id="baseline"></figure><figure><figcaption id="caption"></figcaption><img id="preview"></figure></div><pre id="stats"></pre><script type="application/json" id="data">${data}</script><script>const report=JSON.parse(document.querySelector('#data').textContent);if(report.reference){const image=document.createElement('img');image.src=report.reference.image;image.alt='参考图纸页可见区域';document.querySelector('#reference').append(image);}function show(index){const item=report.samples[index];document.querySelector('#preview').src=item.image;document.querySelector('#caption').textContent='SDK · DPR '+item.dpr+' · '+item.zoom*100+'%';if(report.before){document.querySelector('#before').hidden=false;const baseline=report.before.samples.find(sample=>sample.dpr===item.dpr&&sample.zoom===item.zoom)||report.before.samples[0];document.querySelector('#baseline').src=baseline.image;}document.querySelector('#stats').textContent=JSON.stringify(item,null,2);}report.samples.forEach((item,index)=>{const button=document.createElement('button');button.textContent='DPR '+item.dpr+' / '+item.zoom*100+'%';button.onclick=()=>show(index);document.querySelector('#samples').append(button);});show(Math.max(0,report.samples.findIndex(item=>item.dpr===1&&item.zoom===1.5)));</script></html>`;
  await writeFile(join(artifacts,`${prefix}-gallery.html`),html);console.log(`Gallery: artifacts/${prefix}-gallery.html`);
  await call('Browser.close');
}finally{
  ws?.close();if(browser.exitCode===null){await Promise.race([new Promise(resolve=>browser.once('exit',resolve)),delay(1500)]);if(browser.exitCode===null)browser.kill();}
  await new Promise(resolve=>server.close(resolve));
  if(dirname(resolve(profile))===resolve(tmpdir())&&basename(profile).startsWith('office-pdf-preview-'))await rm(profile,{recursive:true,force:true,maxRetries:8,retryDelay:200}).catch(()=>{});
}
