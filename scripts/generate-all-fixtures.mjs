import { readFile, writeFile } from 'node:fs/promises';
import { generateFixtures } from './generate-fixtures.mjs';
import { generateEmailFixtures } from './email-fixture.mjs';
import { buildPdfFixture } from './pdf-fixture.mjs';
import { generateStyledFixtures } from './legacy-style-fixtures.mjs';
import { generateWordFixtures } from './word-fixtures.mjs';

await generateFixtures();
await generateEmailFixtures();
await writeFile(new URL('../fixtures/sample.pdf', import.meta.url), await buildPdfFixture());
await generateStyledFixtures();
await generateWordFixtures();
// Bundle only our synthetic demo documents. Real user files never enter this asset.
const samples = {};
for (const format of ['csv', 'docx', 'xlsx', 'pptx', 'doc', 'xls', 'ppt', 'pdf', 'eml', 'msg']) {
  const name = `${['xls', 'ppt', 'docx'].includes(format) ? 'styled' : 'sample'}.${format}`;
  samples[format] = { name, base64: (await readFile(new URL(`../fixtures/${name}`, import.meta.url))).toString('base64') };
}
await writeFile(new URL('../fixtures/samples.js', import.meta.url), '// Generated synthetic demo files; regenerate with npm run fixtures.\nexport const samples = ' + JSON.stringify(samples) + ';\n');
console.log('已生成 Word / Excel / PowerPoint / CSV / PDF / EML / MSG 开发样本。');
