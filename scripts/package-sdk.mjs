import { mkdir, readdir, readFile, writeFile } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { makeZip } from './generate-fixtures.mjs';

const root = fileURLToPath(new URL('../', import.meta.url));
const entries = {};
async function collect(directory, prefix) {
  for (const item of await readdir(directory, { withFileTypes: true })) {
    const name = `${prefix}/${item.name}`;
    if (item.isDirectory()) await collect(join(directory, item.name), name);
    else if (item.isFile()) entries[name] = await readFile(join(directory, item.name));
    else throw new Error(`SDK 资源不能包含符号链接：${name}`);
  }
}
await collect(join(root, 'src'), 'src');
for (const name of ['embed.html', 'demo/embed.js']) entries[name] = await readFile(join(root, name));
entries['INTEGRATION.md'] = `# Office Preview SDK

此包只包含浏览器预览程序资源，不含开发样本、用户文件或转换服务。运行时无需 Node、npm、Microsoft Office 或 WPS。当前为自研开发版，格式和视觉兼容性仍有边界。

将整个目录放入宿主的同源静态资源目录，保留目录结构。容器需要明确高度。

\`\`\`js
import { OfficePreview } from './src/index.js';
const preview = new OfficePreview(document.querySelector('#preview'));
await preview.open(file); // File / Blob，或 ArrayBuffer + { name: '报告.pptx' }
// preview.setZoom(1.25);
// preview.fit();
// preview.destroy();
\`\`\`

默认只有内容和右下角缩放。PPT、PDF、Word 连续滚动；表格按上下/左右滚动访问，多工作表保留必要切换。兼容提示在返回模型和 load 事件的 warnings 中，由宿主选择显示。

若宿主自带缩放，构造时传 { showZoom: false }，再调用 setZoom / fit。

iframe 指向 embed.html?parentOrigin=<编码后的宿主精确来源>，等待 office-preview:ready 后发送 { type: 'office-preview:open', requestId, buffer, name }，buffer 用 ArrayBuffer 和 transfer list。核对消息的 source、origin、requestId；成功返回 office-preview:loaded，失败返回 office-preview:error。

跨来源宿主可发送 { type: 'office-preview:setZoom', requestId, zoom: 1.25 } 或 { type: 'office-preview:fit', requestId }，接收 office-preview:zoom；embed.html 的 showZoom=0 参数关闭内部控件。

桌面软件用现有 WebView2 / WKWebView 加载这套资源，再通过 window.officePreview.openBase64(base64, name) 或字节桥传文件。macOS 原生示例另见工程中的 examples/macos。此包不是可直接放入任意原生程序的 DLL。

请通过 HTTP / HTTPS 或宿主正确映射的 Web 资源来源使用，不能直接双击 file:// 页面运行模块与 Worker。宿主需允许同源脚本、Worker 和 blob 图片。原文件当前上限 32MiB。

容器可设置 CSS 属性 --office-preview-background 来配合宿主背景。解析和显示均在本地完成，文档不上传。
`;
const output = join(root, 'artifacts', 'OfficePreviewSDK.zip');
await mkdir(dirname(output), { recursive: true });
const archive = makeZip(entries);
await writeFile(output, archive);
console.log(JSON.stringify({ output, files: Object.keys(entries).length, bytes: archive.length }));
