import test from 'node:test';
import assert from 'node:assert/strict';
import { lstat, mkdir, mkdtemp, readdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { dirname, join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { prepareMacPackage } from '../scripts/package-macos.mjs';

const projectRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const resourcePrefix = 'Sources/OfficePreviewMac/WebAssets/';

async function temporary(t) {
  const directory = await mkdtemp(join(tmpdir(), 'office-preview-macos-test-'));
  t.after(async () => {
    assert.ok(resolve(directory).startsWith(resolve(tmpdir()) + sep), 'cleanup must remain inside the named temporary directory');
    await rm(directory, { recursive: true, force: true });
  });
  return join(directory, 'OfficePreviewMac');
}

async function tree(directory, prefix = '') {
  const files = [];
  for (const item of await readdir(directory, { withFileTypes: true })) {
    const name = prefix ? `${prefix}/${item.name}` : item.name;
    if (item.isDirectory()) files.push(...await tree(join(directory, item.name), name));
    else if (item.isFile()) files.push(name);
  }
  return files.sort();
}

test('macOS source package contains SwiftPM host and a complete local module graph', async t => {
  const output = await temporary(t);
  const result = await prepareMacPackage(output);
  assert.equal(result.outputDir, resolve(output));
  const manifest = JSON.parse(await readFile(join(output, '.office-preview-macos-package.json'), 'utf8'));
  const packageSource = await readFile(join(output, 'Package.swift'), 'utf8');
  assert.match(packageSource, /\.macOS\(\.v14\)/);
  assert.match(packageSource, /\.copy\("WebAssets"\)/);
  assert.match(packageSource, /__info_plist/);
  assert.ok(!packageSource.includes('.package('), 'SwiftPM host must not download third-party packages');
  const assets = join(output, 'Sources', 'OfficePreviewMac', 'WebAssets');
  assert.match(await readFile(join(assets, 'embed.html'), 'utf8'), /demo\/embed\.js/);
  const sourceModules = (await tree(join(projectRoot, 'src'))).filter(name => /\.js$/.test(name));
  for (const name of sourceModules) {
    const packaged = `src/${name}`;
    assert.ok(manifest.files.includes(resourcePrefix + packaged), `missing SDK module ${packaged}`);
    const text = await readFile(join(assets, ...packaged.split('/')), 'utf8');
    for (const match of text.matchAll(/(?:from\s*|import\s*|new URL\(\s*)['"](\.[^'"]+)['"]/g)) {
      const target = resolve(dirname(join(assets, ...packaged.split('/'))), match[1]);
      assert.ok(target.startsWith(assets + sep), `${packaged} imports outside WebAssets`);
      assert.ok((await lstat(target)).isFile(), `${packaged} has an unresolved local reference ${match[1]}`);
    }
  }
  assert.ok(manifest.files.every(name => !/(?:^|\/)(?:fixtures|tests|node_modules)(?:\/|$)/.test(name)));
  assert.ok(manifest.files.filter(name => name.startsWith(resourcePrefix)).every(name => !/\.(?:docx?|xlsx?|pptx?|csv|tsv|pdf|eml|msg)$/i.test(name)), 'user documents must not become bundled WebAssets');
  assert.ok(!manifest.files.some(name => name.includes('PLACEHOLDER')), 'portable source package must contain real SDK assets');
});

test('macOS packager refuses an unmanaged nonempty directory without changing its files', async t => {
  const output = await temporary(t);
  await mkdir(output, { recursive: true });
  await writeFile(join(output, 'README.md'), 'Existing user content');
  await assert.rejects(prepareMacPackage(output), /非空|覆盖/);
  assert.equal(await readFile(join(output, 'README.md'), 'utf8'), 'Existing user content');
  assert.deepEqual(await readdir(output), ['README.md']);
});

test('macOS packager safely refreshes owned files and rejects stray WebAssets documents', async t => {
  const output = await temporary(t);
  const first = await prepareMacPackage(output);
  const second = await prepareMacPackage(output);
  assert.equal(second.files, first.files);
  const userDocument = join(output, 'Sources', 'OfficePreviewMac', 'WebAssets', 'private-document.pdf');
  await writeFile(userDocument, 'Private content must stay untouched');
  await assert.rejects(prepareMacPackage(output), /未受管理/);
  assert.equal(await readFile(userDocument, 'utf8'), 'Private content must stay untouched');
  const manifest = JSON.parse(await readFile(join(output, '.office-preview-macos-package.json'), 'utf8'));
  assert.ok(!manifest.files.some(name => name.endsWith('private-document.pdf')));
});
