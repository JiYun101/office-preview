# macOS 原生嵌入示例

这是 AppKit + WKWebView 的 SwiftPM 源码工程，目标 **macOS 14+、Swift 5.9+**，无第三方 Swift 依赖。相同源码可在 Intel Mac 或 Apple Silicon Mac 上按本机架构编译。本次开发环境为 Windows，已完成资源打包验证与源码审查，**尚未在 macOS 编译或运行**。

## 打包与运行

先在 SDK 工程根目录运行一次：

```sh
node scripts/package-macos.mjs
```

把生成的 `artifacts/macos/OfficePreviewMac` 整个目录复制到 Mac。在该目录执行：

```sh
swift run
```

Mac 需要安装包含 Swift 工具链的 Xcode 或 Command Line Tools。打包完成后的程序运行不需要 Node、不下载 npm 依赖，也不调用外部文档服务。可使用 `swift build -c release` 编译本机 release 可执行文件；生成目录中的资源 bundle 需与可执行文件一起保留。这里交付的是开发源码示例，未生成签名、公证或通用架构 `.app`。

源码模板的 `WebAssets` 只有占位文件。请编译打包后的目录；直接运行未打包模板会明确提示缺少 SDK 资源。每次修改 SDK 后重新执行打包脚本，以更新本机 Web 资源。

## 文件与桥接

窗口收到可信主框架发出的 `office-preview:ready` 后才启用“打开本地文件”。此信号表示 SDK API 已注册；网页导航完成不等于 SDK 就绪。

`NSOpenPanel` 选择本地 DOC/DOCX、XLS/XLSX/CSV/TSV、PPT/PPTX、TXT、PDF、EML 或 MSG 文件。原生端先检查 32MiB 容量，再读取文件、编码为 base64，通过 `callAsyncJavaScript` 的命名参数调用 `window.officePreview.openBase64`，并等待返回的 Promise。文件名、内容不拼接到 JavaScript 源码。具体格式能力以 SDK 的兼容提示为准；此宿主不会补足内核尚未实现的排版、动画或邮件附件能力，也不支持 PST/OST 邮箱数据库。

WebKit 通过 `window.webkit.messageHandlers.officePreview` 向宿主发送 ready、loaded、error。宿主核对 WKWebView、主框架、URL、端口和 security origin。导航只允许本实例的 `http://127.0.0.1:<随机端口>/embed.html` 主框架，其他页面与子框架导航均拒绝。PDF 由 SDK 自己解析并绘制，无需系统 PDF 查看器。子框架不能发送文件桥接消息。

## 本机资源服务器

Network.framework 的 `NWListener` 仅绑定 `127.0.0.1` 随机端口。只读服务器提供打包的 `embed.html`、`demo/embed.js` 和 `src/**` 程序资源，拒绝非 GET/HEAD、路径越界和其他资源路径，不提供上传接口，也不会将选中的文档放入 HTTP 目录。该地址让 ES Modules 与 Worker 使用正常同源加载，无需依赖 `file://` 或自定义 URL scheme 的实现差异。

程序使用非持久 WebKit 数据存储。`Info.plist` 经 SwiftPM 链接设置嵌入可执行文件，只为 `127.0.0.1` 配置 ATS HTTP 例外，没有关闭全部 ATS。若集成到已有 `.app`，应把这一精确的例外合入宿主自己的 Info.plist，并保留当前来源检查。

本示例没有启用 App Sandbox。若嵌入启用沙盒的宿主，需要由该宿主配置 `com.apple.security.network.client`、`com.apple.security.network.server` 以及 `com.apple.security.files.user-selected.read-only`；文件读取使用用户选择 URL 的 security scope。无需给预览器整个磁盘的读取权限。SDK 资源与文件选取流程都应在实际宿主中再次验证。

## Apple / WebKit 一手依据

- [WKWebView.callAsyncJavaScript](https://developer.apple.com/documentation/webkit/wkwebview/callasyncjavascript(_:arguments:in:in:completionhandler:))：命名参数与异步返回值。
- [WKScriptMessage.frameInfo](https://developer.apple.com/documentation/webkit/wkscriptmessage/frameinfo) 和 [WKFrameInfo](https://developer.apple.com/documentation/webkit/wkframeinfo)：发送框架与安全来源。
- [NWParameters.requiredLocalEndpoint](https://developer.apple.com/documentation/network/nwparameters/requiredlocalendpoint)：监听器的本机 IP 与端口。
- [ATS 的 IP 地址例外](https://developer.apple.com/documentation/bundleresources/information-property-list/nsapptransportsecurity/nsexceptiondomains)：macOS 14 起可按 IP 地址设置例外。
- [WebKit Compression Streams](https://docs.webkit.org/Deep%20Dive/Modules/CompressionStreams.html)：包含 `deflate-raw` 的原生解压能力。
