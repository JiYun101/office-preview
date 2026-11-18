import AppKit
import WebKit
import UniformTypeIdentifiers

@main
@MainActor
enum OfficePreviewMac {
    static func main() {
        let application = NSApplication.shared
        let delegate = PreviewApplication()
        application.delegate = delegate
        application.setActivationPolicy(.regular)
        application.run()
        withExtendedLifetime(delegate) {}
    }
}

@MainActor
final class PreviewApplication: NSObject, NSApplicationDelegate, WKNavigationDelegate, WKScriptMessageHandler {
    private var window: NSWindow!
    private var webView: WKWebView!
    private var server: AssetServer?
    private var pageURL: URL?
    private var ready = false
    private let openButton = NSButton(title: "打开本地文件…", target: nil, action: nil)
    private let status = NSTextField(labelWithString: "正在启动本地预览器…")

    func applicationDidFinishLaunching(_ notification: Notification) {
        let configuration = WKWebViewConfiguration()
        configuration.websiteDataStore = .nonPersistent()
        configuration.userContentController.add(self, name: "officePreview")
        webView = WKWebView(frame: .zero, configuration: configuration)
        webView.navigationDelegate = self
        openButton.target = self
        openButton.action = #selector(openFile)
        openButton.isEnabled = false
        let header = NSStackView(views: [openButton, status])
        header.orientation = .horizontal
        header.spacing = 16
        status.setContentHuggingPriority(.defaultLow, for: .horizontal)
        let content = NSView()
        for view in [header, webView!] { view.translatesAutoresizingMaskIntoConstraints = false; content.addSubview(view) }
        NSLayoutConstraint.activate([
            header.topAnchor.constraint(equalTo: content.topAnchor, constant: 12),
            header.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 16),
            header.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -16),
            webView.topAnchor.constraint(equalTo: header.bottomAnchor, constant: 12),
            webView.leadingAnchor.constraint(equalTo: content.leadingAnchor),
            webView.trailingAnchor.constraint(equalTo: content.trailingAnchor),
            webView.bottomAnchor.constraint(equalTo: content.bottomAnchor)
        ])
        window = NSWindow(contentRect: NSRect(x: 0, y: 0, width: 1180, height: 800),
                          styleMask: [.titled, .closable, .miniaturizable, .resizable], backing: .buffered, defer: false)
        window.title = "Office Preview · macOS 开发示例"
        window.contentView = content
        window.center()
        window.makeKeyAndOrderFront(nil)
        NSApplication.shared.activate(ignoringOtherApps: true)
        let menu = NSMenu()
        let applicationItem = NSMenuItem()
        let applicationMenu = NSMenu()
        applicationMenu.addItem(withTitle: "退出 Office Preview", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")
        applicationItem.submenu = applicationMenu
        menu.addItem(applicationItem)
        NSApplication.shared.mainMenu = menu
        do {
            guard let assets = Bundle.module.url(forResource: "WebAssets", withExtension: nil),
                  FileManager.default.fileExists(atPath: assets.appendingPathComponent("embed.html").path) else {
                throw NSError(domain: "OfficePreviewMac", code: 1,
                              userInfo: [NSLocalizedDescriptionKey: "缺少 WebAssets。请先运行 SDK 的 package-macos.mjs，再编译生成目录。"])
            }
            let server = AssetServer(root: assets)
            self.server = server
            try server.start { [weak self] result in
                guard let self else { return }
                switch result {
                case .success(let origin):
                    self.pageURL = origin.appendingPathComponent("embed.html")
                    self.webView.load(URLRequest(url: self.pageURL!))
                case .failure(let error): self.showError(error.localizedDescription)
                }
            }
        } catch { showError(error.localizedDescription) }
    }

    private func trusted(_ url: URL?) -> Bool {
        guard let url, let pageURL else { return false }
        return url.scheme == "http" && url.host == "127.0.0.1" && url.port == pageURL.port &&
            url.path == "/embed.html" && url.user == nil && url.password == nil
    }

    func userContentController(_ userContentController: WKUserContentController, didReceive message: WKScriptMessage) {
        guard message.name == "officePreview", message.webView === webView,
              message.frameInfo.isMainFrame, trusted(message.frameInfo.request.url),
              let port = pageURL?.port, message.frameInfo.securityOrigin.host == "127.0.0.1",
              message.frameInfo.securityOrigin.port == port,
              message.frameInfo.securityOrigin.protocol == "http",
              let body = message.body as? [String: Any], let type = body["type"] as? String else { return }
        switch type {
        case "office-preview:ready":
            ready = true
            openButton.isEnabled = true
            status.stringValue = "准备就绪 · 文件在本机解析"
        case "office-preview:loaded":
            let count = (body["warnings"] as? [String])?.count ?? 0
            status.stringValue = "\(body["name"] as? String ?? "文档") · 兼容提示 \(count) 项"
        case "office-preview:error": status.stringValue = "预览失败：\(body["error"] as? String ?? "未知错误")"
        default: break
        }
    }

    func webView(_ webView: WKWebView, decidePolicyFor navigationAction: WKNavigationAction,
                 decisionHandler: @escaping (WKNavigationActionPolicy) -> Void) {
        if navigationAction.targetFrame?.isMainFrame == true, trusted(navigationAction.request.url) {
            ready = false
            openButton.isEnabled = false
            decisionHandler(.allow)
            return
        }
        decisionHandler(.cancel)
    }

    func webView(_ webView: WKWebView, didFailProvisionalNavigation navigation: WKNavigation!, withError error: Error) {
        showError(error.localizedDescription)
    }

    @objc private func openFile() {
        guard ready, openButton.isEnabled, trusted(webView.url) else { return }
        let panel = NSOpenPanel()
        panel.canChooseFiles = true
        panel.canChooseDirectories = false
        panel.allowsMultipleSelection = false
        panel.allowedContentTypes = ["doc", "docx", "docm", "dotx", "dotm", "xls", "xlsx", "xlsm", "xltx", "xltm",
                                     "ppt", "pptx", "pptm", "ppsx", "ppsm", "potx", "potm", "csv", "tsv", "txt", "pdf", "eml", "msg"]
            .compactMap { UTType(filenameExtension: $0, conformingTo: .data) }
        panel.beginSheetModal(for: window) { [weak self] response in
            guard let self, response == .OK, let url = panel.url, self.ready, self.trusted(self.webView.url) else { return }
            let scoped = url.startAccessingSecurityScopedResource()
            defer { if scoped { url.stopAccessingSecurityScopedResource() } }
            do {
                let limit = 32 * 1024 * 1024
                if (try url.resourceValues(forKeys: [.fileSizeKey]).fileSize ?? 0) > limit {
                    throw NSError(domain: "OfficePreviewMac", code: 2, userInfo: [NSLocalizedDescriptionKey: "文件超过首版 32MB 限制。"])
                }
                let data = try Data(contentsOf: url, options: .mappedIfSafe)
                guard data.count <= limit else {
                    throw NSError(domain: "OfficePreviewMac", code: 2, userInfo: [NSLocalizedDescriptionKey: "文件超过首版 32MB 限制。"])
                }
                self.openButton.isEnabled = false
                self.status.stringValue = "正在解析 \(url.lastPathComponent)…"
                self.webView.callAsyncJavaScript(
                    "return await window.officePreview.openBase64(base64, name)",
                    arguments: ["base64": data.base64EncodedString(), "name": url.lastPathComponent],
                    in: nil, in: .page
                ) { [weak self] result in
                    guard let self else { return }
                    self.openButton.isEnabled = self.ready
                    if case .failure(let error) = result { self.status.stringValue = "预览失败：\(error.localizedDescription)" }
                }
            } catch { self.status.stringValue = "打开失败：\(error.localizedDescription)" }
        }
    }

    private func showError(_ message: String) {
        ready = false
        openButton.isEnabled = false
        status.stringValue = message
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool { true }
    func applicationWillTerminate(_ notification: Notification) {
        webView?.configuration.userContentController.removeScriptMessageHandler(forName: "officePreview")
        server?.stop()
    }
}
