# Office Preview SDK

本项目是纯自研预览内核的**开发验证版**，范围为 Word、Excel、PowerPoint、PDF 和 Outlook 单封邮件。同一套 JavaScript SDK 可嵌入网页，也可由 Windows WebView2 或 macOS WKWebView 承载。文档在本地工作线程中解析，不上传到服务器。

解析器、XML 树、ZIP / OLE 容器读取、文档模型、渲染与嵌入协议均为本项目编写。没有接入 WPS、Microsoft Office、LibreOffice、ONLYOFFICE、商业转换服务或第三方文档解析库，没有 npm 依赖；压缩解码、字符解码、图像显示和字体排版使用浏览器标准 API。仓库为私有开发项目，未设置开源授权。

**当前不具备“全套 Office 完美预览”能力。** 能读取格式与能准确还原页面是不同的目标，下表给出本版实际边界。所有样例都是自产开发样本，通过这些样本不代表通过真实用户文档兼容验收。

## 启动

开发演示和资源打包使用 Node.js 22 或更新版本，浏览器需支持原生 `DecompressionStream('deflate-raw')`；目标包括近期 Chrome / Edge / Safari，以及 WebView2 / WKWebView。无需运行 `npm install`。打包后的 macOS 原生宿主不需要 Node。

```powershell
npm run fixtures
node scripts/serve.mjs
```

仓库不提交 `fixtures/` 中的样例文件（包括内嵌样例字节的 `samples.js`），也不提交 `artifacts/` 中的截图、核验结果和打包产物。首次运行 `npm run fixtures` 会在本机生成自产开发样本；用户原件请仅保存在本机，不要加入版本控制。

打开 <http://127.0.0.1:4173/>。演示工作台提供本地文件选择、开发样本以及 CSV 编码 / 分隔符设置；实际 SDK 预览区使用极简界面，页面连续滚动，默认只保留浮动缩放。工作台外侧的文件选择和导航由演示宿主管理，不会随 SDK 一起嵌入。

开发静态服务仅监听本机，不接收文档上传。请通过 HTTP / HTTPS 使用，不能直接双击 HTML 以 `file://` 运行 ES modules 和 Worker。

## 内嵌预览界面

SDK 默认显示文档和右下角缩放控件，没有顶部文件信息栏、兼容提示横幅、底部状态栏、PPT 左侧页签或上一页 / 下一页按钮。Word、PPT 和 PDF 按纸页从上到下连续滚动；PPT / PDF 的远端页面按需绘制。电子表格通过纵向和横向滚动浏览，多个工作表时保留紧凑的工作表切换。

文件选择、文件名、加载进度和兼容提示由宿主管理。`open()` 返回的模型及 `load` 事件仍包含 `warnings`，iframe 的 `office-preview:loaded` 也会返回它们；去掉界面横幅不改变解析能力或兼容边界。打开失败时组件仍显示明确错误，并通知宿主。

若宿主已有缩放工具，可用 `new OfficePreview(container, { showZoom: false })` 只挂载内容，通过实例的 `setZoom()` / `fit()` 控制。iframe 可用 `showZoom=0` 关闭内部缩放；同源 iframe 和原生 WebView 可调用现有 JS 桥，跨来源 iframe 使用下面的缩放消息。

## 本版格式能力

| 格式 | 已实现 | 未实现 / 重要边界 |
| --- | --- | --- |
| CSV / TSV | 引号、引号转义、字段换行、空字段；逗号 / 分号 / 制表符；UTF-8 / BOM UTF-16 / GB18030 | 自动编码和分隔符识别可能需要手动纠正；所有字段按文字显示，不执行公式 |
| DOCX | 默认 / 段落 / 字符样式继承、中西文字体槽；目录编号、缩进、点引导制表位与缓存页码；行距 / 网格；独立纸页、保存分页与自动排版；行内及无环绕浮图的页面 / 边距 / 段落 / 单元格锚点；表格不等宽网格、纵横合并、行高、单边边框、内边距、底色和垂直对齐 | 整本精确分页、页眉页脚、脚注、复杂浮动环绕与图文组合对齐、表格样式条件继承、跨页表格拆分及固定行高溢出裁剪、公式与图表等仍不完整；目录页码不重新计算 |
| XLSX | 多工作表、共享 / inline 字符串、数字 / 布尔 / 错误值、基本样式、列宽、合并区域、公式缓存值 | 尚无公式计算、完整数字 / 日期格式、图表、图片、条件格式、筛选和打印布局；隐藏内容仍可能显示，并给出提示 |
| PPTX | 有效页序、文本与坐标、部分图片、基础形状、背景与线性渐变、实际圆角 / 线宽、字距 / 段落间距、组坐标及部分占位符位置 | 静态内容子集；尚无完整母版 / 主题还原、路径渐变、SmartArt、图表、动画、切换、音视频和完整字体效果 |
| DOC 97–2003 | FIB / CLX / PieceTable；STSH、CHPX/PAPX 字符和段落格式；字体 / 目录制表位 / 自动列表 / 缓存字段；SEPX 页面、行网格与显式分页；PICF / BSE / BLIP 常见行内图片 | 完整表格、页眉页脚、浮动环绕及原生分页缓存仍有限制；加密不支持；仅当声明数据完整时接受末尾未填充的数据扇区 |
| XLS 97–2003 | BIFF8、多工作表、共享字符串、缓存值、合并；FONT / XF / PALETTE 样式；字体、对齐、填充、边框、行高列宽；常见数字 / 日期 / 百分比格式 | 图表、绘图、条件格式、打印版式和公式计算尚未还原；复杂数字格式仍有限制；BIFF5 及更早版本不支持；隐藏表跳过并提示 |
| PPT 97–2003 | 当前编辑 / 持久对象索引、有效页序；真实尺寸、颜色、线性渐变、圆角、基础母版；字体 / 段落及文件内兼容数据中的字距 / 小数字号；常见图片；裁剪、单对象旋转与翻转 | 路径 / 特殊焦点渐变、旋转组、图表 / OLE、音视频和动画仍不完整；旧格式更早版本不支持；无可解析绘图时明确提示并降级为文字 |
| TXT | 文字与换行 | 无样式 |
| PDF | 自主读取页面、嵌入 TrueType / OpenType 字形、CID→GID 映射、粗斜体；基础路径、填充 / 描边透明度与文字绘制模式；Canvas 随缩放和屏幕 DPI 重绘，纵向连续滚动 | 基础静态预览；原始 Type1 / CFF、Type3、自定义 CID Encoding、文字裁剪、复杂透明合成、图案、表单和批注等尚不完整，不调用系统内置 PDF 查看器 |
| EML / MSG | 单封邮件的主题、收发件人、纯文本正文、附件清单；EML MIME 解码、MSG 常用 MAPI 属性 | HTML 转为文字，不执行 HTML、不加载外链；附件只列清单；PST / OST 邮箱存档尚未实现 |

RTF、老式 XML Office、宏 / 模板扩展名、WPS 私有格式、Visio、Project、Access、OneNote、PST / OST 邮箱存档等尚未接入。密码解密和损坏文件修复未实现。不会执行文档宏、公式脚本、外部关系和嵌入对象。

Office 预览需要字体在宿主环境可用，中文字体已提供 Windows / macOS 常见字体回退；PDF 优先读取文件内嵌字体。字体替换和浏览器字体度量仍会影响断行。Word 按文件纸张和边距显示独立页面，随容器自动排列并纵向滚动；优先保留保存的分页边界，缺失处自动排版，不能保证总页数与原生 Word / WPS 完全一致。旧 XLS 界面可解释常见数字 / 日期格式，原始值和公式缓存保持不变；XLSX 当前界面仍按原始值 / 缓存值显示。

PDF 读取真实的交叉引用与页树，支持传统 xref 表、xref 流、对象流和增量更新；可读取 ToUnicode 映射、基本文本 / 路径 / 裁剪、页面裁剪框 / 旋转、Form XObject、JPEG 与部分 8 位光栅图像。TrueType / OpenType 嵌入字形通过原始 GID 绘制，保留轮廓、度量和提示信息；PDF 字宽只推进后续文字位置，不拉伸嵌入字形。缺失、未支持或无法加载的字体使用系统回退并返回兼容提示。填充与描边分别保留透明度，支持常量 alpha；软蒙版、特殊混合模式、透明度组和图案等仍不完整。加密、未实现的内容流压缩或内联图像等会明确拒绝。不会执行 PDF JavaScript、外部动作或附件。

PDF 画布按当前缩放与实际屏幕 DPI 生成像素，切换 DPI 时也会重新绘制；保留已完成页面直到新画布绘制结束。每页仍受 1600 万像素 / 8192 边长上限约束，超高缩放时清晰度会受此限制。字体按预览实例注册，打开新文件或销毁组件时释放，不依赖宿主安装相同字体。

DOCX 浮图按文件的参考框、对齐方式及偏移定位，不参与行内图片的宽高和字号度量。表格保留源网格和单元格格式，行高区分 auto / atLeast / exact；用实际生效的内边距和边框计算行框，不使用报告名称或正文触发特殊分支。无环绕图片可覆盖表格边界和签署行；相对段落定位的浮图超过单元格可用高度时，将该格内容对齐到顶部，保持文字与浮图锚点一致。复杂图文组合的完整包围范围对齐和文字绕排仍不完整。

邮件是文件预览能力，不连接 Outlook 账户。HTML 正文降级为纯文本；附件仅显示元数据，不打开附件、远程图片或嵌入邮件。压缩 RTF、S/MIME 解密与签名验证尚未实现。

## 网页直接集成

可运行的[直接 SDK 示例](http://127.0.0.1:4173/examples/sdk.html)与[iframe 示例](http://127.0.0.1:4173/examples/iframe.html)都把文件选择放在宿主一侧。运行 `node scripts/package-sdk.mjs` 可生成 `artifacts/OfficePreviewSDK.zip`，解压后直接部署到宿主静态目录；包内只有预览程序和接入说明，不含开发样本、用户文档或 Node 运行时。

将 `src/` 作为静态资源部署，保留模块目录结构即可，无构建步骤，也不需要第三方前端框架。宿主只需提供一个有实际高度的 HTML 容器，并传入已有权限读取的 `File` 或字节。SDK 使用 Shadow DOM 隔离样式，每次解析在独立 Worker 中运行。宿主 CSP 需允许同源模块 / Worker，显示内部图片时需允许 `img-src blob:`。

```html
<div id="preview" style="height:700px"></div>
<script type="module">
  import { OfficePreview } from './src/index.js';

  const preview = new OfficePreview(document.querySelector('#preview'));
  preview.addEventListener('load', event => {
    console.log(event.detail.format, event.detail.warnings);
  });
  preview.addEventListener('error', event => console.error(event.detail.message));

  // file 来自 input.files[0] 或拖放事件。错误通过 Promise 拒绝返回。
  // await preview.open(file);
  // 也可传 ArrayBuffer / Uint8Array，但须给出带扩展名的文件名：
  // await preview.open(buffer, { name: '报告.docx' });
  // preview.setZoom(1.2);
  // preview.destroy();
</script>
```

容器可设置 CSS 自定义属性 `--office-preview-background` 配合宿主背景色；若已有全局缩放工具，可关闭内部控件，示例见上文。

`open()` 返回规范化文档模型；`getDocument()` 读取当前模型；`getCapabilities()` 返回当前支持声明。`setZoom(1.25)` 表示 125%，`fit()` 恢复适合容器宽度；`destroy()` 释放实例。连续打开时，之前的解析会取消并以 `AbortError` 拒绝。宿主应处理 `open()` 的拒绝，销毁实例后不能再次打开。

CSV 可传 `{ encoding: 'gb18030', delimiter: ',' }`；省略时自动识别。SDK 仅打开宿主提供的文件字节，不自动请求任意文件 URL；宿主已有权限 / 存储体系继续负责文件读取。

## iframe 集成

可运行示例：<http://127.0.0.1:4173/examples/iframe.html>。

```html
<iframe id="preview" title="文件预览" style="width:100%;height:700px;border:0"></iframe>
```

宿主脚本设置 iframe 的 `src`，监听就绪后把文件字节传入；单独设置 `src` 不会自动取得宿主的文件权限。示例保留完整的就绪、加载、失败和新文件覆盖旧请求的处理。

1. iframe 指向 `embed.html?parentOrigin=<编码后的宿主来源>`，例如宿主为 `https://app.example.com`，传入这个精确来源，不带路径和末尾斜线。
2. 宿主等待来自该 iframe、来源匹配的 `office-preview:ready` 消息。
3. 发送 `{ type: 'office-preview:open', requestId, buffer, name, options }`，其中 `buffer` 为 ArrayBuffer，建议使用 transfer list。
4. 预览器返回 `office-preview:loaded`（含兼容提示）或 `office-preview:error`。宿主同时检查 `event.source`、`event.origin` 和 `requestId`。

默认缩放由 iframe 内部提供。宿主需要自己的缩放工具时，将 `showZoom=0` 加入 iframe URL，在文件加载完成后发送下列消息：

```js
frame.contentWindow.postMessage({
  type: 'office-preview:setZoom', requestId: 'zoom-1', zoom: 1.25
}, viewerOrigin);

// 恢复适合容器宽度：
frame.contentWindow.postMessage({ type: 'office-preview:fit', requestId: 'fit-1' }, viewerOrigin);
```

组件返回 `{ type: 'office-preview:zoom', requestId, zoom }`，其中 `zoom` 是实际生效值（当前范围 0.2–3）。缩放消息只接受有限数字，使用与打开文件相同的来源和父窗口校验；协议错误返回 `office-preview:error`。同源 iframe / WebView 的 JS 桥也可调用 `window.officePreview.setZoom(1.25)` 和 `window.officePreview.fit()`。

跨来源嵌入时，部署服务器须允许宿主页面作为 `frame-ancestors`。当前组件不需要跨来源隔离、数据库、文档转换服务器或额外授权服务。

| 宿主 | 接入方式 | 宿主提供 |
| --- | --- | --- |
| 网页系统 | 直接导入 SDK，或嵌入 iframe | 容器高度、文件或字节、已有文件权限 |
| Windows 桌面软件 | WebView2 承载同一份 SDK / `embed.html` | 本地资源来源与原生文件读取 |
| macOS 桌面软件 | WKWebView 承载同一份 SDK / `embed.html` | 本地资源来源与原生文件读取 |

这是浏览器 / WebView 组件，不是向任意原生窗口直接放入一个 DLL 的控件。已有 WebView 的宿主可以复用现有容器；其他桌面程序需要先接入相应 WebView。

## 桌面 WebView 集成

Windows 使用宿主现有的 WebView2；macOS 提供 AppKit + WKWebView 源码示例。两者加载同一份 `embed.html` 与 `src/`；本地文件由原生宿主读取，不依赖 iframe 获得磁盘权限。

WebView 加载完成后，宿主可调用：

```js
await window.officePreview.openBase64(base64Bytes, '报告.docx');
// 或在 WebView 内用 JavaScript 取得 Uint8Array 后：
await window.officePreview.open(bytes, '报告.docx');
```

原生端拼接 JavaScript 参数时使用 JSON 序列化器。WebView2 可将资源目录映射到可信虚拟 HTTPS 来源，再加载 `https://office.local/embed.html`。Windows 原生窗口封装尚未提供。WKWebView 示例使用 `callAsyncJavaScript` 的命名参数传文件名和内容，并等待返回的 Promise。

## macOS 接入与打包

源码见 [examples/macos/README.md](examples/macos/README.md)，目标为 **macOS 14+、Swift 5.9+**。同一源码可在 Intel Mac 或 Apple Silicon Mac 按本机架构构建；当前开发环境是 Windows，尚未在 Mac 编译或运行，不能视为已完成真机兼容验收。

```sh
node scripts/package-macos.mjs
# 将 artifacts/macos/OfficePreviewMac 整个目录复制到 Mac
# 在该目录运行：
swift run
```

这是可移植的 SwiftPM 开发源码包，包含 SDK 资源、原生文件选择、来源检查和 JS 桥，不是已签名或公证的 `.app` / 安装包。为让 ES Modules 和 Worker 正常使用同源加载，示例内置一个仅绑定 `127.0.0.1` 随机端口的只读程序资源服务；无需 Node、远程文档服务或数据库。所选文档不会放入 HTTP 目录。

网页发出 `office-preview:ready` 后原生端才开放文件选择，加载与失败通过 `window.webkit.messageHandlers.officePreview` 回传。宿主核对主框架、来源和端口。字体栈已加入 Apple 系统字体与苹方；页面断行仍受可用字体影响。

共享代码的浏览器检查可在 Mac 的 Chrome / Edge 中运行；Safari / WKWebView 需要在 Mac 上单独执行相同样本验收。WebKit 的原生解压依据：[Compression Streams](https://docs.webkit.org/Deep%20Dive/Modules/CompressionStreams.html)。

## 开发与验证

```powershell
node scripts/generate-all-fixtures.mjs
node --test tests/*.test.mjs
node scripts/browser-check.mjs --screenshot
```

浏览器检查使用本机 Edge / Chrome 和标准 DevTools 协议，无自动化框架依赖。若未自动发现浏览器，可设置 `OFFICE_PREVIEW_BROWSER`。脚本创建并清理独立临时浏览器配置，不使用日常浏览器用户目录。

浏览器检查覆盖基础解析与真实 SDK / Worker、连续滚动、二维表格窗口、缩放 / PDF 清晰度、异步取消与资源释放、文本安全、跨来源 iframe / 原生 JS 桥，以及已有文字与样式场景。当前 **143 项 Node 检查、84 项浏览器检查通过**，包含自产嵌入字体的真实绘制、字形比例、文字模式、透明度合成、缩放像素密度、FontFace 释放，以及 DOCX 非等宽表格、行框、合并、浮图锚点与溢出对齐。测试不能取代所有真实文件的兼容验收，macOS 原生宿主仍未在真机编译运行。

项目包含 Node 格式读取测试与真实浏览器回归，浏览器检查另覆盖极简 iframe、纵向连续页面、缩放、宿主协议、错误恢复和按需绘制；执行结果以对应命令输出为准。界面简化只改变组件交互，不扩大格式支持声明，也不代表所有内容与原生 Office 完全一致。真实样例的原件、截图及具体核验数据仅保存在本机。

Word 原件核验命令：`node scripts/word-preview-check.mjs --files <DOCX路径> <DOC路径> --reference <参考截图> --tag current`。它通过真实本地 File 读取，输出前三张纸页、首个含图页、布局统计和人工对照页到 `artifacts/word-preview-*`。整本分页仍有差异，不能拿预览页数重写或校验原目录页码。图片按需加载，避免同时解码整本文档的图片。

PDF 原件核验命令：`node scripts/pdf-preview-check.mjs --file <PDF路径> --reference <参考截图> --tag current`。它通过真实 File → Worker → SDK，记录 100% / 150% / 200% 与 DPR 1 / 2 的实际纸页、画布像素、字体加载和色块元数据，输出截图及等比对照页到 `artifacts/pdf-preview-*`；只保存核验结果，不将用户原件或内嵌字体另存进源码和样本。浏览器光栅化和截图缩放仍可能产生差异，核验指标不代表任意 PDF 的兼容率。

表格报告核验命令：`node scripts/word-report-check.mjs --file <DOCX路径> --reference <参考截图> --tag current --self-test`。自产样本通过独立原生度量核对三种行高规则与浮图对齐；原生程序仅用于开发核验，SDK 的生产路径始终为自主解析和浏览器绘制。用户原件保持只读，核验图和元数据仅存于 `artifacts/word-report-*`，不进入 SDK 或 macOS 源码包。

新增原图视觉核验：`node scripts/slide-visual-check.mjs --reference <原图PNG> --files <PPTX路径> <PPT路径> --tag current`。实际 File → Worker → SDK 预览按参考宽度等比渲染，保持纵横比例，只裁去参考范围以外的底部，不拉伸图片。输出同尺寸 PNG、差异图、分区像素指标与交互叠图页，均保存在 `artifacts/slide-visual-*`，不进入源码包。

渐变色彩插值、部分字形的栅格化仍可能存在残差；像素指标不是兼容率，也不能代表其他页面或文档。参考图与用户原件均未放进开发样本或 macOS 包。

浏览器验证环境为 Windows 上的 Edge；Safari / WKWebView 与 macOS 原生窗口未编译或运行验证。PDF / 邮件 / 旧版样式浏览器检查通过真实本地文件选择传入 File，避免把本机代理的下载行为混入解析结果。演示的“开发样本”按钮使用只包含自产样本的静态字节资源，实际用户文件始终由本地选择 / 宿主传入。可用 `node scripts/browser-check.mjs --inspect-files <文件路径...>` 在本机查看任意已授权文件；只输出版式元数据和截图，不把原文件加入源码包。

首版容量限制：原文件 32MiB；ZIP 单项解压 32MiB，总计 128MiB，最多 10000 项；XML 最多 200000 个节点 / 属性，深度 128；工作表最多 50000 行、1024 列，最多 1000000 个数据单元格。表格界面每段显示 200 行、64 列，合并区域跨段裁剪显示。超过边界时明确报错。

PDF 最多 500 页，限制流解码和图像展开容量；EML 限制 MIME 嵌套深度与分段数。当前容量边界用于开发验证，不代表生产性能或任意真实文件的完整兼容保证。

## 内核结构和后续验收

链路为 `文件字节 → ZIP / CFB / CSV → 格式读取器 → 统一文档模型 → 渲染器 → SDK / iframe / WebView`。当前只有服务于这条链路的模块，没有数据库、转换服务和插件框架。

达到高还原目标，后续需要以真实样本逐项推进：Word 的样式、字体度量、分页与环绕；Excel 的数字格式、公式与图表；PowerPoint 的母版、形状、动画与媒体；再补充更早格式和其他 Office 产品。每项能力都应以样本文档的内容、外观和交互结果验收，不能仅以文件成功打开验收。

自研实现参考公开格式规范，没有复制第三方项目实现：

- [Microsoft MS-CFB 容器规范](https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/53989ce4-7b05-4f8d-829b-d08d6148375b)
- [Microsoft MS-DOC 正文读取](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-doc/01d5d8c4-cf9c-4ef9-80fd-439e763cfe01)
- [Microsoft MS-XLS Unicode 字符串](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/173d9f51-e5d3-43da-8de2-be7f22e119b9)
- [Microsoft MS-PPT 文档流](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-ppt/1fc22d56-28f9-4818-bd45-67c2bf721ccf)
- [Microsoft Open XML 文档](https://learn.microsoft.com/en-us/office/open-xml/)
- [Word 行高实现注记](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/5919e0bd-e6ce-477e-8d66-0e5282f5c506)
- [RFC 5322 邮件结构](https://www.rfc-editor.org/rfc/rfc5322.html)、[RFC 2045 MIME](https://www.rfc-editor.org/rfc/rfc2045.html)、[RFC 2047 头字段编码](https://www.rfc-editor.org/rfc/rfc2047.html)
- [Microsoft MS-OXMSG 属性流](https://learn.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxmsg/ed03f930-fdef-4cc8-b400-b69007d7f416)
