// Read actual local Word files through File → Worker → SDK and capture paper pages.
import { spawn } from 'node:child_process';
import { access, copyFile, mkdir, mkdtemp, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { basename, dirname, extname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { startServer } from './serve.mjs';

const root=resolve(dirname(fileURLToPath(import.meta.url)),'..'),args=process.argv.slice(2);
const option=name=>{const index=args.indexOf(name);return index<0?undefined:args[index+1];};
const tag=option('--tag')||'current',reference=option('--reference'),files=[];
const index=args.indexOf('--files');
if(index>=0)for(const argument of args.slice(index+1)){if(argument.startsWith('--'))break;files.push(resolve(argument));}
if(!files.length||!/^[a-z0-9-]+$/i.test(tag))throw new Error('Usage: node scripts/word-preview-check.mjs --files <docx> <doc> [--reference <screenshot>] [--tag current]');
await Promise.all([...files,...(reference?[resolve(reference)]:[])].map(file=>access(file)));
const candidates=[process.env.OFFICE_PREVIEW_BROWSER,'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe','C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe','/Applications/Google Chrome.app/Contents/MacOS/Google Chrome','/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge','/usr/bin/chromium','/usr/bin/google-chrome'].filter(Boolean);
let executable;for(const path of candidates){try{await access(path);executable=path;break;}catch{}}
if(!executable)throw new Error('No Chromium browser found. Set OFFICE_PREVIEW_BROWSER.');
const artifacts=join(root,'artifacts'),prefix=`word-preview-${tag}`;await mkdir(artifacts,{recursive:true});
const profile=await mkdtemp(join(tmpdir(),'office-word-preview-')),server=await startServer({port:0});
const origin=`http://127.0.0.1:${server.address().port}`;
const browser=spawn(executable,['--headless=new','--disable-gpu','--no-first-run','--no-default-browser-check','--remote-debugging-port=0',`--user-data-dir=${profile}`,'about:blank'],{windowsHide:true,stdio:'ignore'});
const delay=ms=>new Promise(resolve=>setTimeout(resolve,ms));let ws,failed=false;
try{
  let address;const startup=Date.now()+20000;
  while(Date.now()<startup){try{address=(await readFile(join(profile,'DevToolsActivePort'),'utf8')).trim().split('\n');break;}catch{await delay(100);}}
  if(!address)throw new Error('Browser startup timed out.');
  ws=new WebSocket(`ws://127.0.0.1:${address[0]}${address[1]}`);
  await new Promise((resolve,reject)=>{ws.addEventListener('open',resolve,{once:true});ws.addEventListener('error',reject,{once:true});});
  let sequence=0;const pending=new Map(),errors=[];
  ws.addEventListener('message',event=>{const message=JSON.parse(event.data);if(message.method==='Runtime.exceptionThrown')errors.push(message.params.exceptionDetails.exception?.description||message.params.exceptionDetails.text);if(pending.has(message.id)){const task=pending.get(message.id);pending.delete(message.id);clearTimeout(task.timer);message.error?task.reject(new Error(message.error.message)):task.resolve(message.result);}});
  const call=(method,params={},sessionId)=>new Promise((resolve,reject)=>{const id=++sequence,timer=setTimeout(()=>{pending.delete(id);reject(new Error('CDP timeout: '+method));},60000);pending.set(id,{resolve,reject,timer});ws.send(JSON.stringify({id,method,params,...(sessionId?{sessionId}:{})}));});
  const target=await call('Target.createTarget',{url:'about:blank'}),{sessionId}=await call('Target.attachToTarget',{targetId:target.targetId,flatten:true});
  const send=(method,params)=>call(method,params,sessionId);
  await send('Page.enable');await send('Runtime.enable');
  const evaluate=async expression=>{const response=await send('Runtime.evaluate',{expression,returnByValue:true,awaitPromise:true});if(response.exceptionDetails)throw new Error(response.exceptionDetails.exception?.description||response.exceptionDetails.text);return response.result.value;};
  const deadline=Date.now()+30000;await send('Page.navigate',{url:origin+'/tests/word-browser.html?inspect=1'});
  while(!await evaluate('!!window.wordPreview')){if(Date.now()>deadline)throw new Error('Word page failed to load: '+errors.slice(-3).join('\n'));await delay(100);}
  await send('Emulation.setDeviceMetricsOverride',{width:1400,height:1700,deviceScaleFactor:1,mobile:false});
  const documentNode=await send('DOM.getDocument'),input=await send('DOM.querySelector',{nodeId:documentNode.root.nodeId,selector:'#fixture-files'});
  await send('DOM.setFileInputFiles',{nodeId:input.nodeId,files});
  const report={tag,method:'Native File → Worker → SDK; paper pages captured at 100% CSS zoom. No Word engine conversion.',files:[]};
  for(let index=0;index<files.length;index++){
    try{
      const metadata=await evaluate(`window.wordPreview.open(${index})`),screenshots=[];
      const captureIndices=[...new Set([0,1,2,metadata.firstImagePageIndex].filter(page=>Number.isInteger(page)&&page>=0&&page<metadata.pageCount))];
      for(const page of captureIndices){
        const bounds=await evaluate(`window.wordPreview.page(${page})`);
        if(!(bounds.width>0&&bounds.height>0&&bounds.width<=8192&&bounds.height<=8192))throw new Error('Paper page dimensions exceed capture bounds.');
        const capture=await send('Page.captureScreenshot',{format:'png',clip:{x:bounds.x,y:bounds.y,width:bounds.width,height:bounds.height,scale:1},captureBeyondViewport:true});
        const filename=`${prefix}-${index+1}-${metadata.format}-page-${page+1}.png`;
        await writeFile(join(artifacts,filename),Buffer.from(capture.data,'base64'));screenshots.push({image:filename,...bounds});
      }
      report.files.push({...metadata,screenshots});console.log(`${metadata.format.toUpperCase()}: ${metadata.pageCount} pages, ${metadata.tocLeaderCount} TOC leaders, ${metadata.imageCount} images (${metadata.sampledImages} sampled, ${metadata.failedImages} failed).`);
      if(metadata.failedImages)failed=true;
    }catch(error){failed=true;report.files.push({name:basename(files[index]),format:extname(files[index]).slice(1),error:error.message});console.error(`${basename(files[index])}: ${error.message}`);}
  }
  if(reference){const name=prefix+'-reference.png';await copyFile(resolve(reference),join(artifacts,name));report.referenceImage=name;}
  const reportName=prefix+'-inspection.json';await writeFile(join(artifacts,reportName),JSON.stringify(report,null,2));
  const data=JSON.stringify(report).replaceAll('<','\\u003c');
  const html=`<!doctype html><html lang="zh-CN"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Word 纸页预览检查</title><style>body{font:14px/1.7 system-ui;background:#edf1f5;color:#25334a;margin:24px}h1{font-size:22px}.reference{max-width:100%;display:block;margin-bottom:24px}.pages{display:flex;gap:18px;align-items:flex-start;overflow:auto;padding:12px}.pages img{width:min(48vw,794px);height:auto;box-shadow:0 2px 9px #2333}pre{background:white;padding:16px;white-space:pre-wrap}nav{display:flex;gap:12px}button{padding:8px 14px}</style><h1>Word 原件预览检查</h1><p>原图包含应用界面和不同缩放，以下只作人工排版对照，不声称逐像素一致。所有文档由本地自研 SDK 读取，未转换为参考截图。</p><div id="reference"></div><nav id="files"></nav><div id="pages" class="pages"></div><pre id="stats"></pre><script type="application/json" id="data">${data}</script><script>const report=JSON.parse(document.querySelector('#data').textContent);if(report.referenceImage){const image=document.createElement('img');image.className='reference';image.src=report.referenceImage;image.alt='用户提供的参考截图';document.querySelector('#reference').append(image);}function show(index){const item=report.files[index],pages=document.querySelector('#pages');pages.replaceChildren();for(const page of item.screenshots||[]){const image=document.createElement('img');image.src=page.image;image.alt='SDK纸页预览';pages.append(image);}document.querySelector('#stats').textContent=JSON.stringify(item,null,2);}report.files.forEach((item,index)=>{const button=document.createElement('button');button.textContent=item.name;button.onclick=()=>show(index);document.querySelector('#files').append(button);});show(0);</script></html>`;
  await writeFile(join(artifacts,prefix+'-gallery.html'),html);console.log(`Gallery: artifacts/${prefix}-gallery.html`);
  await call('Browser.close');
}finally{
  ws?.close();
  if(browser.exitCode===null){await Promise.race([new Promise(resolve=>browser.once('exit',resolve)),delay(1500)]);if(browser.exitCode===null)browser.kill();}
  await new Promise(resolve=>server.close(resolve));
  if(dirname(resolve(profile))===resolve(tmpdir())&&basename(profile).startsWith('office-word-preview-'))await rm(profile,{recursive:true,force:true,maxRetries:8,retryDelay:200}).catch(()=>{});
}
if(failed)process.exitCode=1;
