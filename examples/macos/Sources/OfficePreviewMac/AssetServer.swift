import Foundation
import Network

// The web worker and ES modules need a normal same-origin HTTP URL. This server
// serves bundled program assets only; user-selected document bytes never enter it.
final class AssetServer {
    private let root: URL
    private let queue = DispatchQueue(label: "OfficePreviewMac.assets")
    private var listener: NWListener?

    init(root: URL) { self.root = root.standardizedFileURL.resolvingSymlinksInPath() }

    func start(_ completion: @escaping (Result<URL, Error>) -> Void) throws {
        let parameters = NWParameters.tcp
        parameters.requiredLocalEndpoint = .hostPort(host: "127.0.0.1", port: .any)
        let listener = try NWListener(using: parameters)
        self.listener = listener
        listener.newConnectionHandler = { [weak self] connection in
            guard let self else { connection.cancel(); return }
            connection.start(queue: self.queue)
            self.receive(connection, buffered: Data())
        }
        listener.stateUpdateHandler = { [weak listener] state in
            switch state {
            case .ready:
                guard let port = listener?.port,
                      let url = URL(string: "http://127.0.0.1:\(port.rawValue)/") else { return }
                DispatchQueue.main.async { completion(.success(url)) }
            case .failed(let error):
                listener?.cancel()
                DispatchQueue.main.async { completion(.failure(error)) }
            default: break
            }
        }
        listener.start(queue: queue)
    }

    func stop() { listener?.cancel(); listener = nil }

    private func receive(_ connection: NWConnection, buffered: Data) {
        connection.receive(minimumIncompleteLength: 1, maximumLength: 8192) { [weak self] chunk, _, complete, error in
            guard let self, error == nil else { connection.cancel(); return }
            var bytes = buffered
            if let chunk { bytes.append(chunk) }
            if bytes.count > 16384 { self.respond(connection, status: "431 Request Header Fields Too Large"); return }
            if let boundary = bytes.range(of: Data("\r\n\r\n".utf8)) {
                self.handle(connection, header: bytes[..<boundary.lowerBound])
            } else if complete { connection.cancel() }
            else { self.receive(connection, buffered: bytes) }
        }
    }

    private func handle(_ connection: NWConnection, header: Data.SubSequence) {
        guard let text = String(data: Data(header), encoding: .utf8) else {
            respond(connection, status: "400 Bad Request"); return
        }
        let lines = text.components(separatedBy: "\r\n")
        let request = (lines.first ?? "").split(separator: " ", omittingEmptySubsequences: false)
        guard request.count == 3, request[2] == "HTTP/1.1" || request[2] == "HTTP/1.0" else {
            respond(connection, status: "400 Bad Request"); return
        }
        let method = String(request[0])
        guard method == "GET" || method == "HEAD" else {
            respond(connection, status: "405 Method Not Allowed"); return
        }
        guard let port = listener?.port,
              lines.dropFirst().contains(where: { line in
                  let field = line.split(separator: ":", maxSplits: 1).map(String.init)
                  return field.count == 2 && field[0].lowercased() == "host" &&
                      field[1].trimmingCharacters(in: .whitespaces) == "127.0.0.1:\(port.rawValue)"
              }) else { respond(connection, status: "400 Bad Request"); return }
        let target = String(request[1]).split(separator: "?", maxSplits: 1, omittingEmptySubsequences: false)[0]
        guard let path = String(target).removingPercentEncoding, path.hasPrefix("/"),
              !path.contains("\\"), !path.contains("\0"),
              !path.components(separatedBy: "/").contains("..") else {
            respond(connection, status: "403 Forbidden"); return
        }
        guard path == "/embed.html" || path == "/demo/embed.js" || path.hasPrefix("/src/") else {
            respond(connection, status: "404 Not Found"); return
        }
        let file = root.appendingPathComponent(String(path.dropFirst()))
            .standardizedFileURL.resolvingSymlinksInPath()
        guard file.path.hasPrefix(root.path + "/"), let bytes = try? Data(contentsOf: file) else {
            respond(connection, status: "404 Not Found"); return
        }
        let mime = ["html": "text/html; charset=utf-8", "js": "text/javascript; charset=utf-8",
                    "css": "text/css; charset=utf-8", "json": "application/json; charset=utf-8",
                    "wasm": "application/wasm", "svg": "image/svg+xml", "png": "image/png"]
        respond(connection, status: "200 OK", bytes: bytes,
                mime: mime[file.pathExtension] ?? "application/octet-stream", head: method == "HEAD")
    }

    private func respond(_ connection: NWConnection, status: String, bytes: Data = Data(),
                         mime: String = "text/plain; charset=utf-8", head: Bool = false) {
        let header = "HTTP/1.1 \(status)\r\nContent-Type: \(mime)\r\nContent-Length: \(bytes.count)\r\nCache-Control: no-store\r\nX-Content-Type-Options: nosniff\r\nConnection: close\r\n\r\n"
        var response = Data(header.utf8)
        if !head { response.append(bytes) }
        connection.send(content: response, completion: .contentProcessed { _ in connection.cancel() })
    }
}
