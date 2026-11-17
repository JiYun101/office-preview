// Capture the actual SDK slide at reference resolution and measure pixel differences.
// Native File selection, the production Worker and renderer are used throughout.
import { spawn } from 'node:child_process';
import { access, copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { startServer } from './serve.mjs';

const root=resolve(dirname(fileURLToPath(import.meta.url)),'..');
const args=process.argv.slice(2);
const option=name=>{const at=args.indexOf(name);return at<0?null:args[at+1];};
const referencePath=option('--reference');
const tag=option('--tag')||'current';
const at=args.indexOf('--files'),files=[];
if(at>=0)for(const value of args.slice(at+1)){if(value.startsWith('--'))break;files.push(resolve(value));}
if(!referencePath||!files.length||!/^[a-z0-9-]+$/i.test(tag))throw new Error('Usage: node scripts/slide-visual-check.mjs --reference <image> --files <pptx> <ppt> [--tag baseline]');
const reference=resolve(referencePath);
await Promise.all([reference,...files].map(file=>access(file)));
const prefix=`slide-visual-${tag}`;
const candidates=[process.env.OFFICE_PREVIEW_BROWSER,'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe','C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge','/usr/bin/chromium','/usr/bin/google-chrome'].filter(Boolean);
let executable;
for(const candidate of candidates){try{await access(candidate);executable=candidate;break;}catch{}}
if(!executable)throw new Error('Chromium / Edge was not found. Set OFFICE_PREVIEW_BROWSER.');
const artifacts=join(root,'artifacts');await mkdir(artifacts,{recursive:true});
const profile=await mkdtemp(join(tmpdir(),'office-slide-visual-'));
const server=await startServer({port:0});
const origin=`http://127.0.0.1:${server.address().port}`;
const browser=spawn(executable,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--remote-debugging-port=0',`--user-data-dir=${profile}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay=ms=>new Promise(resolve=>setTimeout(resolve,ms));
let ws;
try{
  let address;
  const deadline=Date.now()+20000;
  while(Date.now()<deadline){try{address=(await readFile(join(profile,'DevToolsActivePort'),'utf8')).trim().split('\n');break;}catch{await delay(100);}}
  if(!address)throw new Error('Browser debug endpoint startup timed out.');
  ws=new WebSocket(`ws://127.0.0.1:${address[0]}${address[1]}`);
  await new Promise((resolve,reject)=>{ws.addEventListener('open',resolve,{once:true});ws.addEventListener('error',reject,{once:true});});
  let sequence=0;
  const pending=new Map(),errors=[];
  ws.addEventListener('message',event=>{
    const message=JSON.parse(event.data);
    if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.exception?.description||message.params.exceptionDetails.text);
    if(pending.has(message.id)){const task=pending.get(message.id);pending.delete(message.id);clearTimeout(task.timer);message.error?task.reject(new Error(message.error.message)):task.resolve(message.result);}
  });
  const call=(method,params={},sessionId)=>new Promise((resolve,reject)=>{
    const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP timeout: '+method));},30000);
    pending.set(id,{resolve,reject,timer});ws.send(JSON.stringify({id,method,params,...(sessionId?{sessionId}:{})}));
  });
  const target=await call('Target.createTarget',{url:'about:blank'});
  const {sessionId}=await call('Target.attachToTarget',{targetId:target.targetId,flatten:true});
  const send=(method,params)=>call(method,params,sessionId);
  await send('Page.enable');await send('Runtime.enable');
  const evaluate=async expression=>{const result=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(result.exceptionDetails)throw new Error(result.exceptionDetails.exception?.description||result.exceptionDetails.text);return result.result.value;};
  const waitFor=async expression=>{const deadline=Date.now()+20000;while(Date.now()<deadline){const result=await evaluate(expression);if(result)return result;await delay(100);}throw new Error('Page readiness timeout: '+errors.slice(-3).join('\n'));};
  await send('Page.navigate',{url:origin+'/tests/slide-visual.html'});
  await waitFor('!!window.slideVisual');
  const testStyle=option('--style-text');
  if(testStyle)await evaluate(`window.slideVisual.addStyle(${JSON.stringify(testStyle)})`);
  const choose=async(selector,paths)=>{const document=await send('DOM.getDocument');const node=await send('DOM.querySelector',{nodeId:document.root.nodeId,selector});await send('DOM.setFileInputFiles',{nodeId:node.nodeId,files:paths});};
  await choose('#reference',[reference]);await choose('#documents',files);
  const size=await evaluate('window.slideVisual.referenceSize()');
  if(args.includes('--color-check')) {
    const colors=await evaluate('window.slideVisual.colorInterpolation()');
    await writeFile(join(artifacts,`${prefix}-colors.json`),JSON.stringify(colors,null,2));
    console.log(JSON.stringify(colors.spaces.map(item=>({space:item.space,mae:item.mae,samples:item.samples.filter(sample=>[477,955,1433].includes(sample.x))})),null,2));
  }
  if(size.width>8192||size.height>8192||size.width*size.height>32000000)throw new Error('Reference dimensions exceed visual-check capacity.');
  await send('Emulation.setDeviceMetricsOverride',{width:size.width,height:size.height,deviceScaleFactor:1,mobile:false});
  const report={tag,referenceName:basename(reference),width:size.width,height:size.height,method:'Native File → Worker → SDK. Width-preserving scale; reference-height crop, no image stretching.',testStyle:testStyle||null,threshold:30,files:[]};
  const referenceName=`${prefix}-reference.png`;
  if(extname(reference).toLowerCase()==='.png')await copyFile(reference,join(artifacts,referenceName));
  else{
    const png=await evaluate("(async()=>{const f=document.querySelector('#reference').files[0];const u=URL.createObjectURL(f);const i=new Image();i.src=u;await i.decode();const c=document.createElement('canvas');c.width=i.naturalWidth;c.height=i.naturalHeight;c.getContext('2d').drawImage(i,0,0);URL.revokeObjectURL(u);return c.toDataURL('image/png').split(',')[1]})()");
    await writeFile(join(artifacts,referenceName),Buffer.from(png,'base64'));
  }
  report.referenceImage=referenceName;
  for(let index=0;index<files.length;index++){
    const prepared=await evaluate(`window.slideVisual.prepare(${index},${size.width},${size.height})`);
    const geometry=await evaluate('window.slideVisual.geometry()');
    const capture=await send('Page.captureScreenshot',{format:'png',clip:{x:prepared.bounds.x,y:prepared.bounds.y,width:size.width,height:size.height,scale:1},captureBeyondViewport:true});
    const extension=extname(files[index]).slice(1).replace(/[^a-z0-9]/gi,'').toLowerCase()||'document';
    const imageName=`${prefix}-${index+1}-${extension}.png`,diffName=`${prefix}-${index+1}-${extension}-diff.png`;
    await writeFile(join(artifacts,imageName),Buffer.from(capture.data,'base64'));
    const compared=await evaluate(`window.slideVisual.compare(${JSON.stringify(capture.data)})`);
    await writeFile(join(artifacts,diffName),Buffer.from(compared.diffPng,'base64'));
    delete compared.diffPng;
    const entry={...prepared,image:imageName,diffImage:diffName,...compared,geometry};
    report.files.push(entry);
    console.log(`${prepared.name}: ${size.width}×${size.height}; MAE=${compared.metrics.whole.maeRGB.toFixed(4)}; changed=${(compared.metrics.whole.changedPixelRatio*100).toFixed(3)}%; top-strip MAE=${compared.metrics.topStrip.maeRGB.toFixed(4)}`);
  }
  const reportName=`${prefix}-metrics.json`;
  await writeFile(join(artifacts,reportName),JSON.stringify(report,null,2));
  const json=JSON.stringify(report).replaceAll('<','\\u003c');
  const html=`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Slide visual comparison</title><style>body{margin:24px;font:14px/1.6 system-ui;background:#f4f6fa;color:#26354b}h1{font-size:22px}nav,header{display:flex;align-items:center;gap:16px;flex-wrap:wrap;margin:12px 0}button{padding:6px 12px;cursor:pointer}label{display:flex;align-items:center;gap:8px}.stage{position:relative;width:min(100%,${size.width}px);aspect-ratio:${size.width}/${size.height};background:#fff;overflow:hidden;border:1px solid #ccd4e0}.stage img{display:block;width:100%;height:100%;position:absolute;inset:0;object-fit:contain}.row{display:grid;grid-template-columns:1fr 1fr;gap:16px}.row img{display:block;width:100%}pre{white-space:pre-wrap;max-height:320px;overflow:auto;background:white;padding:16px;font-size:12px}</style><h1>原图与 SDK 预览 · 同尺寸视觉对照</h1><p>滑块控制原图覆盖比例；指标反映像素差异，包含字体光栅化差异，不能代替人工视觉验收。</p><nav id="files"></nav><header><label>原图透明度<input id="opacity" type="range" min="0" max="100" value="50"><span id="value">50%</span></label><a href="${reportName}">完整测量与 DOM 几何</a></header><div class="stage"><img id="preview"><img id="reference" src="${referenceName}" style="opacity:.5"></div><h2>并排与差异</h2><div class="row"><img src="${referenceName}" alt="原图"><img id="side" alt="SDK 预览"></div><img id="diff" style="max-width:100%;margin-top:16px" alt="绝对RGB差异放大3倍"><pre id="metrics"></pre><script type="application/json" id="data">${json}</script><script>const data=JSON.parse(document.querySelector('#data').textContent);function show(i){const item=data.files[i];for(const key of ['preview','side'])document.getElementById(key).src=item.image;document.querySelector('#diff').src=item.diffImage;document.querySelector('#metrics').textContent=JSON.stringify({name:item.name,metrics:item.metrics,topStripSamples:item.topStripSamples},null,2)}data.files.forEach((item,i)=>{const b=document.createElement('button');b.textContent=item.name;b.onclick=()=>show(i);document.querySelector('#files').append(b)});document.querySelector('#opacity').oninput=e=>{document.querySelector('#reference').style.opacity=e.target.value/100;document.querySelector('#value').textContent=e.target.value+'%'};show(0);</script></html>`;
  const htmlName=`${prefix}-comparison.html`;
  await writeFile(join(artifacts,htmlName),html);
  console.log('Comparison: artifacts/'+htmlName);console.log('Metrics: artifacts/'+reportName);
  await call('Browser.close');
}finally{
  ws?.close();await new Promise(resolve=>server.close(resolve));
  if(browser.exitCode===null){await Promise.race([new Promise(resolve=>browser.once('exit',resolve)),delay(1500)]);if(browser.exitCode===null)browser.kill();}
  if(dirname(resolve(profile))===resolve(tmpdir())&&basename(profile).startsWith('office-slide-visual-'))await rm(profile,{recursive:true,force:true,maxRetries:8,retryDelay:200}).catch(()=>{});
}
