import test from 'node:test';
import assert from 'node:assert/strict';
import { buildWordPageAreasDoc } from '../scripts/word-fixtures.mjs';
import { parseDocument } from '../src/parse.js';

const textOf = blocks => blocks.flatMap(block => block.runs).map(run => run.text).join('');

test('DOC page areas read across pieces after footnotes and preserve section variants and cached fields', async () => {
  const model = await parseDocument(buildWordPageAreasDoc(), { name: 'page-areas.doc' });
  const [first, next] = model.sections;
  assert.equal(model.evenAndOddHeaders, true);
  assert.equal(first.titlePage, true);
  assert.equal(first.pageNumberStart, 4);
  assert.equal(first.headerDistance, 32);
  assert.equal(first.footerDistance, 24);
  assert.equal(textOf(first.headers.first), 'First header');
  assert.equal(textOf(first.headers.default), 'Odd header');
  assert.equal(textOf(first.headers.even), 'Even header');
  assert.equal(textOf(first.footers.first), 'First footer');
  assert.equal(textOf(first.footers.default), '5');
  assert.equal(textOf(first.footers.even), 'Even footer');
  assert.equal(next.titlePage, false);
  assert.equal(next.pageNumberStart, 8);
  assert.equal(textOf(next.headers.default), 'Next odd header');
  assert.equal(textOf(next.headers.even), 'Even header');
  assert.equal(textOf(next.footers.default), '5');
  assert.equal(textOf(model.blocks), 'First body\fSecond body\fThird bodyNext section\fLast body');
  assert.ok(model.warnings.some(message => message.includes('文件保存的结果')));
});

test('an invalid DOC header index leaves main-story text readable', async () => {
  const model = await parseDocument(buildWordPageAreasDoc({ invalidIndex: true }), { name: 'page-areas.doc' });
  assert.equal(model.sections[0].headers, undefined);
  assert.equal(textOf(model.blocks), 'First body\fSecond body\fThird bodyNext section\fLast body');
  assert.ok(model.warnings.some(message => message.includes('页眉页脚索引无效')));
});
