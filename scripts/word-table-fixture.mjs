// Self-authored OOXML table sample. No user document data or images are used.
// String entries work with both Node and the browser's tiny test ZIP writer.
export function buildWordTableFixtureEntries({ percentage = false } = {}) {
  const w = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
  const r = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const p = text => `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:line="300" w:lineRule="exact"/></w:pPr><w:r><w:rPr><w:sz w:val="21"/></w:rPr><w:t>${text}</w:t></w:r></w:p>`;
  const cell = (text, properties = '') => `<w:tc><w:tcPr>${properties}<w:vAlign w:val="center"/></w:tcPr>${p(text)}</w:tc>`;
  const row = (height, rule, cells) => `<w:tr><w:trPr><w:trHeight w:val="${height}" w:hRule="${rule}"/></w:trPr>${cells}</w:tr>`;
  const border = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(side => `<w:${side} w:val="single" w:sz="6" w:color="000000"/>`).join('');
  const properties = `<w:tblW w:w="${percentage ? 5000 : 9000}" w:type="${percentage ? 'pct' : 'dxa'}"/><w:tblLayout w:type="fixed"/><w:tblBorders>${border}</w:tblBorders><w:tblCellMar><w:top w:w="30" w:type="dxa"/><w:left w:w="60" w:type="dxa"/><w:bottom w:w="30" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar>`;
  const body = row(600, 'exact', cell('A', '<w:gridSpan w:val="2"/><w:shd w:val="clear" w:fill="DDEEFF"/>') + cell('B', '<w:vMerge w:val="restart"/>') + cell('C')) +
    row(900, 'atLeast', cell('D', '<w:gridSpan w:val="2"/>') + cell('', '<w:vMerge/>') + cell('E')) +
    row(450, 'atLeast', ['F', 'G', 'H', 'I'].map(text => cell(text)).join(''));
  const table = `<w:tbl><w:tblPr>${properties}</w:tblPr><w:tblGrid>${[900, 1800, 2700, 3600].map(width => `<w:gridCol w:w="${width}"/>`).join('')}</w:tblGrid>${body}</w:tbl>`;
  return {
    '[Content_Types].xml': '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': `<Relationships xmlns="${r}"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
    'word/_rels/document.xml.rels': `<Relationships xmlns="${r}"/>`,
    'word/document.xml': `<w:document xmlns:w="${w}"><w:body>${table}<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1200" w:right="1200" w:bottom="1200" w:left="1200"/></w:sectPr></w:body></w:document>`,
  };
}
