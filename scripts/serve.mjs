import { createServer } from 'node:http';
import { readFile, stat } from 'node:fs/promises';
import { dirname, extname, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const types = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript; charset=utf-8', '.mjs': 'text/javascript; charset=utf-8', '.css': 'text/css; charset=utf-8', '.json': 'application/json; charset=utf-8', '.svg': 'image/svg+xml', '.png': 'image/png', '.csv': 'text/csv; charset=utf-8', '.md': 'text/plain; charset=utf-8' };
Object.assign(types, { '.doc': 'application/msword', '.xls': 'application/vnd.ms-excel', '.ppt': 'application/vnd.ms-powerpoint', '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' });
Object.assign(types, { '.pdf': 'application/pdf', '.eml': 'message/rfc822', '.msg': 'application/vnd.ms-outlook' });

export async function startServer({ port = 4173, host = '127.0.0.1' } = {}) {
  const server = createServer(async (request, response) => {
    try {
      if (!['GET', 'HEAD'].includes(request.method)) { response.writeHead(405).end(); return; }
      const url = new URL(request.url, 'http://localhost');
      const pathname = decodeURIComponent(url.pathname);
      const relative = pathname === '/' ? 'index.html' : pathname.slice(1);
      if (relative.split(/[\\/]/).some(part => part.startsWith('.'))) { response.writeHead(404).end(); return; }
      const target = resolve(root, relative);
      if (!target.startsWith(root + sep)) { response.writeHead(403).end(); return; }
      if (!(await stat(target)).isFile()) { response.writeHead(404).end(); return; }
      const data = await readFile(target);
      response.writeHead(200, { 'Content-Type': types[extname(target)] || 'application/octet-stream', 'Content-Length': data.length, 'Cache-Control': 'no-store', 'X-Content-Type-Options': 'nosniff' });
      response.end(request.method === 'HEAD' ? undefined : data);
    } catch (error) { response.writeHead(error.code === 'ENOENT' || error.code === 'ENOTDIR' ? 404 : 400).end('File unavailable'); }
  });
  await new Promise((accept, reject) => { server.once('error', reject); server.listen(port, host, accept); });
  return server;
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  const index = process.argv.indexOf('--port');
  const port = index < 0 ? 4173 : Number(process.argv[index + 1]);
  const server = await startServer({ port });
  console.log(`Office 预览开发演示：http://127.0.0.1:${server.address().port}`);
  console.log('静态服务仅用于开发；文件由浏览器本地读取，不接收上传。');
}
