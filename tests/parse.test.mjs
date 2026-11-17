import test from 'node:test';
import assert from 'node:assert/strict';
import { parseDocument } from '../src/parse.js';
import { readZip } from '../src/containers/zip.js';
import { buildXlsx, makeZip } from '../scripts/generate-fixtures.mjs';

async function workbookWithMerges(ranges) {
  const parts = await readZip(buildXlsx());
  const name = 'xl/worksheets/sheet1.xml';
  const original = new TextDecoder().decode(parts.get(name));
  const merges = `<mergeCells count="${ranges.length}">${ranges.map(ref => `<mergeCell ref="${ref}"/>`).join('')}</mergeCells>`;
  const modified = original.replace(/<mergeCells\b[\s\S]*?<\/mergeCells>/, merges);
  assert.notEqual(modified, original, 'development fixture merge replacement failed');
  parts.set(name, new TextEncoder().encode(modified));
  return makeZip(Object.fromEntries(parts));
}

test('parseDocument rejects overlapping XLSX merges beyond the first two visible segments', async () => {
  const bytes = await workbookWithMerges(['A401:B402', 'B402:C403']);
  await assert.rejects(parseDocument(bytes, { name: 'overlapping.xlsx' }), /重叠.*合并|合并.*重叠/);
});

test('parseDocument accepts adjacent, non-overlapping XLSX merges beyond row 400', async () => {
  const bytes = await workbookWithMerges(['A401:A402', 'B401:B402']);
  const model = await parseDocument(bytes, { name: 'adjacent.xlsx' });
  assert.equal(model.kind, 'spreadsheet');
  assert.deepEqual(model.sheets[0].merges, [
    { startRow: 400, startCol: 0, endRow: 401, endCol: 0 },
    { startRow: 400, startCol: 1, endRow: 401, endCol: 1 },
  ]);
});
