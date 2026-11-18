import { binary, warn } from './legacy-binary.js';

// MS-DOC 2.4.3, 2.9.313, 2.9.317, and 2.9.321.
const px = 15;
const sides = ['top', 'left', 'bottom', 'right'];
const verticalAlignments = ['top', 'center', 'bottom'];
const signed16 = value => value << 16 >> 16;

function readBorder(data, start, palette, table = false) {
  const type = data.u8(start + 1);
  if (data.u32(start) === 0xffffffff || !type && !table) return undefined;
  if (!type) return { style: 'none', width: 0 };
  return {
    style: ({ 3: 'double', 6: 'dotted', 7: 'dashed', 8: 'dashed', 9: 'dashed' })[type] || 'solid',
    width: Math.max(type === 3 ? 18 : 2, data.u8(start)) / 6,
    color: palette[data.u8(start + 2)] || '#000000',
  };
}

function readDefinition(bytes, palette) {
  if (!bytes.length) return null;
  const data = binary(bytes), count = data.u8(0), cellsStart = 1 + (count + 1) * 2;
  if (!count || count > 63 || bytes.length < cellsStart || (bytes.length - cellsStart) % 20) return null;
  const edges = Array.from({ length: count + 1 }, (_, index) => signed16(data.u16(1 + index * 2)));
  if (edges.some((edge, index) => index > 0 && edge <= edges[index - 1])) return null;
  const cells = Array.from({ length: count }, (_, index) => {
    const offset = cellsStart + index * 20;
    if (offset >= bytes.length) return {};
    const flags = data.u16(offset), borders = {};
    for (const [side, key] of sides.entries()) {
      const border = readBorder(data, offset + 4 + side * 4, palette);
      if (border) borders[key] = border;
    }
    return { horizontalMerge: flags & 3, verticalMerge: (flags >>> 5) & 3, verticalAlign: verticalAlignments[(flags >>> 7) & 3], borders };
  });
  return { edges, cells };
}

// Word can encode a row with a sequence of cell inserts and width changes.
function readInsertedDefinition(prls) {
  const cells = [];
  let left = 0, gap = 0;
  for (const { id, bytes } of prls) {
    const data = binary(bytes), start = bytes[0], end = bytes[1];
    if (id === 0x9601) left = signed16(data.u16(0));
    else if (id === 0x9602) gap = signed16(data.u16(0));
    else if (id === 0x7621) {
      if (start > cells.length || !end || cells.length + end > 63) return null;
      const width = data.u16(2);
      cells.splice(start, 0, ...Array.from({ length: end }, () => ({ width })));
    } else if ([0x5622, 0x7623, 0x5624, 0x5625].includes(id)) {
      if (start >= end || end > cells.length) return null;
      if (id === 0x5622) cells.splice(start, end - start);
      else for (let index = start; index < end; index++) {
        if (id === 0x7623) cells[index].width = data.u16(2);
        else cells[index].horizontalMerge = id === 0x5625 ? 0 : index === start ? 2 : 1;
      }
    } else if (id === 0xd62b && bytes.length === 2 && cells[start]) cells[start].verticalMerge = end;
  }
  if (!cells.length || cells.some(cell => !cell.width)) return null;
  const edges = [left - gap];
  for (const { width } of cells) edges.push(edges.at(-1) + width);
  return edges.at(-1) - edges[0] <= 31680 ? { edges, cells } : null;
}

export function readDocTableParagraph(prls, palette) {
  let inTable = false, depth, rowEnd = false, cellEnd = false;
  for (const { id, bytes } of prls) {
    if (id === 0x2416) inTable = bytes[0] === 1;
    else if (id === 0x2417 || id === 0x244c) rowEnd = bytes[0] === 1;
    else if (id === 0x244b) cellEnd = bytes[0] === 1;
    else if (id === 0x6649) depth = binary(bytes).u32(0) | 0;
    else if (id === 0x664a) depth = (depth ?? Number(inTable)) + (binary(bytes).u32(0) | 0);
  }
  depth ??= Number(inTable);
  if (depth <= 0) return null;
  if (depth !== 1 || !rowEnd) return { depth, rowEnd, cellEnd };
  const definition = prls.findLast(prl => prl.id === 0xd608);
  const height = prls.findLast(prl => prl.id === 0x9407);
  const alignment = prls.findLast(prl => prl.id === 0x548a || prl.id === 0x5400);
  const outerBorders = prls.findLast(prl => prl.id === 0xd605 && prl.bytes.length === 24);
  const borders = outerBorders ? Object.fromEntries([...sides, 'insideH', 'insideV'].map((side, index) => [side, readBorder(binary(outerBorders.bytes), index * 4, palette, true)])) : undefined;
  return {
    depth, rowEnd, cellEnd,
    definition: definition ? readDefinition(definition.bytes, palette) : readInsertedDefinition(prls),
    height: height ? signed16(binary(height.bytes).u16(0)) : 0,
    align: alignment ? ['left', 'center', 'right'][binary(alignment.bytes).u16(0)] : undefined,
    borders,
  };
}

function tableBlock(entries, warnings) {
  const rows = [];
  let cells = [], paragraphs = [];
  for (const { block, mark, table } of entries) {
    if (table.depth !== 1) return null;
    if (mark === '\x07' && table.rowEnd) {
      if (!cells.length || paragraphs.length || block.runs.length) return null;
      rows.push({ cells, ...table });
      cells = [];
      continue;
    }
    paragraphs.push(block);
    if (mark === '\x07') { cells.push(paragraphs); paragraphs = []; }
    if (cells.length > 63) return null;
  }
  if (!rows.length || cells.length || paragraphs.length) return null;

  const first = rows[0], edges = first.definition?.edges, start = edges?.[0], end = edges?.at(-1);
  const measured = rows.every(({ cells, definition }) => definition && definition.cells.length === cells.length && definition.edges[0] === start && definition.edges.at(-1) === end);
  const distinct = new Set();
  if (measured) for (const row of rows) {
    for (const edge of row.definition.edges) distinct.add(edge);
    if (distinct.size > 1025) { distinct.clear(); break; }
  }
  const boundaries = new Map(Array.from(distinct).sort((a, b) => a - b).map((edge, column) => [edge, column]));
  if (!boundaries.size) warn(warnings, '部分 DOC 表格缺少一致的列边界，按基本表格显示。');
  const columnWidths = [];
  let previous;
  for (const edge of boundaries.keys()) {
    if (previous !== undefined) columnWidths.push((edge - previous) / px);
    previous = edge;
  }
  let activeMerges = new Map();
  const grid = rows.map(row => {
    const rendered = [], continuing = new Map(), definition = boundaries.size ? row.definition : null;
    for (let index = 0; index < row.cells.length; index++) {
      const plan = definition?.cells[index] || {}, blocks = row.cells[index];
      const start = definition ? boundaries.get(definition.edges[index]) : index;
      let end = definition ? boundaries.get(definition.edges[index + 1]) : index + 1;
      let borders = plan.borders;
      if (plan.horizontalMerge >= 2) {
        while (definition.cells[index + 1]?.horizontalMerge === 1) {
          index++;
          end = boundaries.get(definition.edges[index + 1]);
          const right = definition.cells[index].borders?.right;
          if (right) borders = { ...borders, right };
        }
      }
      const key = `${start}:${end}`, anchor = activeMerges.get(key);
      if (plan.verticalMerge === 1 && anchor && blocks.every(block => !block.runs.length)) {
        anchor.rowSpan = (anchor.rowSpan || 1) + 1;
        if (borders?.bottom) anchor.borders = { ...anchor.borders, bottom: borders.bottom };
        continuing.set(key, anchor);
        continue;
      }
      if (plan.verticalMerge === 1 || plan.horizontalMerge === 1) warn(warnings, '部分 DOC 表格合并标记缺少有效起始单元格，按独立单元格显示。');
      const cell = { blocks, column: start, ...(end - start > 1 ? { colSpan: end - start } : {}), ...(plan.verticalAlign ? { verticalAlign: plan.verticalAlign } : {}), ...(borders ? { borders } : {}) };
      if (plan.verticalMerge === 3) continuing.set(key, cell);
      rendered.push(cell);
    }
    activeMerges = continuing;
    return rendered;
  });
  return {
    type: 'table', rows: grid, columnWidths, align: first.align, sectionIndex: entries[0].block.sectionIndex,
    indent: boundaries.size ? start / px : 0,
    ...(first.borders ? { borders: first.borders } : {}),
    cellMargins: { top: 0, bottom: 0, left: 7.2, right: 7.2 },
    rowProperties: rows.map(({ height }) => height ? { height: Math.abs(height) / px, heightRule: height < 0 ? 'exact' : 'atLeast' } : {}),
  };
}

export function assembleDocTables(entries, warnings) {
  const blocks = [];
  let pending = [];
  function flush() {
    if (!pending.length) return;
    const table = tableBlock(pending, warnings);
    if (table) blocks.push(table);
    else {
      warn(warnings, '部分 DOC 表格含嵌套或不完整的行列标记，按正文顺序显示。');
      for (const { block, mark, table } of pending) if (block.runs.length || mark === '\r' && !table.rowEnd) blocks.push(block);
    }
    pending = [];
  }
  for (const entry of entries) {
    if (!entry.table?.depth || pending.length && pending[0].block.sectionIndex !== entry.block.sectionIndex) flush();
    if (entry.table?.depth) pending.push(entry);
    else blocks.push(entry.block);
  }
  flush();
  return blocks;
}
