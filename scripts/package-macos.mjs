import { lstat, mkdir, readdir, readFile, unlink, writeFile } from 'node:fs/promises';
import { dirname, extname, isAbsolute, join, parse, relative, resolve, sep } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const markerName = '.office-preview-macos-package.json';
const packageKind = 'office-preview-macos-source';
const assetRoot = 'Sources/OfficePreviewMac/WebAssets';
const assetExtensions = new Set(['.js', '.mjs', '.css', '.json', '.wasm', '.svg', '.png', '.woff', '.woff2']);

async function info(path) {
  try { return await lstat(path); }
  catch (error) { if (error.code === 'ENOENT') return null; throw error; }
}

function validRelative(path) {
  return typeof path === 'string' && path.length && !isAbsolute(path) && !/[\\\0:]/.test(path) &&
    path.split('/').every(part => part && part !== '.' && part !== '..');
}

async function filesUnder(directory, prefix = '') {
  const result = [];
  for (const entry of await readdir(directory, { withFileTypes: true })) {
    const name = prefix ? `${prefix}/${entry.name}` : entry.name;
    if (entry.isSymbolicLink()) throw new Error(`拒绝跟随资源符号链接：${name}`);
    if (entry.isDirectory()) result.push(...await filesUnder(join(directory, entry.name), name));
    else if (entry.isFile()) result.push(name);
  }
  return result.sort();
}

async function safeTarget(output, name) {
  if (!validRelative(name)) throw new Error(`无效的打包路径：${name}`);
  let target = output;
  const parts = name.split('/');
  for (let index = 0; index < parts.length; index++) {
    target = join(target, parts[index]);
    const current = await info(target);
    if (current?.isSymbolicLink()) throw new Error(`拒绝覆盖符号链接：${name}`);
    if (current && (index < parts.length - 1 ? !current.isDirectory() : !current.isFile())) throw new Error(`打包路径类型冲突：${name}`);
  }
  if (!target.startsWith(output + sep)) throw new Error('打包目标超出输出目录。');
  return target;
}

export async function prepareMacPackage(outputDir = join(projectRoot, 'artifacts', 'macos', 'OfficePreviewMac')) {
  const output = resolve(outputDir);
  if (output === parse(output).root || output === projectRoot || projectRoot.startsWith(output + sep)) throw new Error('输出目录不能覆盖工程或磁盘根目录。');
  const existing = await info(output);
  if (existing && (!existing.isDirectory() || existing.isSymbolicLink())) throw new Error('输出目标必须是普通目录。');
  let previous = { files: [] };
  const marker = existing ? await info(join(output, markerName)) : null;
  if (marker) {
    if (!marker.isFile() || marker.isSymbolicLink()) throw new Error('打包标记必须是普通文件。');
    previous = JSON.parse(await readFile(join(output, markerName), 'utf8'));
    if (previous.kind !== packageKind || previous.version !== 1 || !Array.isArray(previous.files) || !previous.files.every(validRelative)) throw new Error('输出目录不是受支持的 macOS 源码包。');
  } else if (existing && (await readdir(output)).length) {
    throw new Error('输出目录非空且未经本脚本管理，拒绝覆盖已有文件。');
  }
  const files = new Map();
  const templateFiles = ['Package.swift', 'Info.plist', 'README.md', 'Sources/OfficePreviewMac/App.swift', 'Sources/OfficePreviewMac/AssetServer.swift'];
  for (const name of templateFiles) files.set(name, await readFile(join(projectRoot, 'examples', 'macos', ...name.split('/'))));
  files.set(`${assetRoot}/embed.html`, await readFile(join(projectRoot, 'embed.html')));
  files.set(`${assetRoot}/demo/embed.js`, await readFile(join(projectRoot, 'demo', 'embed.js')));
  // Discover all current SDK modules, including newly added PDF/mail parsers.
  for (const name of await filesUnder(join(projectRoot, 'src'))) {
    if (assetExtensions.has(extname(name).toLowerCase())) files.set(`${assetRoot}/src/${name}`, await readFile(join(projectRoot, 'src', ...name.split('/'))));
  }
  const owned = new Set(previous.files);
  const targets = new Map();
  for (const name of new Set([...previous.files, ...files.keys()])) {
    const target = await safeTarget(output, name);
    if (await info(target) && !owned.has(name)) throw new Error(`拒绝覆盖未受管理的文件：${name}`);
    targets.set(name, target);
  }
  const resources = join(output, ...assetRoot.split('/'));
  if (await info(resources)) {
    for (const name of await filesUnder(resources)) {
      if (!owned.has(`${assetRoot}/${name}`)) throw new Error(`WebAssets 中存在未受管理的文件，拒绝将其纳入资源包：${name}`);
    }
  }
  await mkdir(output, { recursive: true });
  for (const [name, bytes] of files) {
    const target = targets.get(name);
    await mkdir(dirname(target), { recursive: true });
    await writeFile(target, bytes);
  }
  // Remove only obsolete files named by our previous manifest; never remove a tree.
  for (const name of previous.files) if (!files.has(name) && await info(targets.get(name))) await unlink(targets.get(name));
  const manifest = { kind: packageKind, version: 1, files: [...files.keys()].sort() };
  await writeFile(join(output, markerName), JSON.stringify(manifest, null, 2) + '\n');
  return { outputDir: output, files: files.size, resources: manifest.files.filter(name => name.startsWith(assetRoot + '/')).length };
}

if (process.argv[1] && pathToFileURL(resolve(process.argv[1])).href === import.meta.url) {
  if (process.argv.length > 3) throw new Error('用法：node scripts/package-macos.mjs [输出目录]');
  const result = await prepareMacPackage(process.argv[2]);
  console.log(JSON.stringify(result, null, 2));
  console.log('已生成可复制到 Mac 的 SwiftPM 源码包；当前平台未执行 macOS 编译或运行验证。');
}
