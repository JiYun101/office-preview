import test from 'node:test';
import assert from 'node:assert/strict';
import { parseXml, elements, children, first, attr } from '../src/xml.js';

test('XML reader works without DOMParser and preserves mixed text, CDATA and entities', () => {
  assert.equal(typeof globalThis.DOMParser, 'undefined');
  const doc = parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<!-- outside -->
    <r xmlns="urn:root" title='&quot;&apos;&amp;&lt;&gt;'>A &amp; B<child><![CDATA[<tag>&raw;]]></child>&#20013;&#x1F600;<?inspect value?></r>`);
  assert.equal(doc.documentElement.localName, 'r');
  assert.equal(doc.documentElement.namespaceURI, 'urn:root');
  assert.equal(doc.documentElement.textContent, 'A & B<tag>&raw;中😀');
  assert.equal(attr(doc.documentElement, 'title'), `"'&<>`);
  assert.equal(first(doc, 'child').textContent, '<tag>&raw;');
  assert.equal(doc.getElementsByTagNameNS('*', 'r').length, 1);
  assert.equal(doc.documentElement.getElementsByTagNameNS('*', 'r').length, 0);
  assert.equal(children(doc.documentElement, 'child').length, 1);
});

test('XML namespace scopes distinguish id/r:id and expanded attribute names', () => {
  const doc = parseXml('<p:r xmlns:p="urn:one" xmlns:r="urn:relationships" id="256" r:id="slideA"><p:c xmlns:p="urn:two" xmlns="urn:default" plain="x"><item/></p:c><p:c/></p:r>');
  const root = doc.documentElement;
  assert.deepEqual(root.attributes.filter(item => item.localName === 'id').map(item => [item.name, item.namespaceURI, item.value]), [
    ['id', '', '256'], ['r:id', 'urn:relationships', 'slideA'],
  ]);
  assert.equal(root.children[0].namespaceURI, 'urn:two');
  assert.equal(root.children[1].namespaceURI, 'urn:one');
  assert.equal(root.children[0].attributes.find(item => item.name === 'plain').namespaceURI, '');
  assert.equal(elements(doc, 'item')[0].namespaceURI, 'urn:default');
  assert.equal(doc.getElementsByTagNameNS('urn:one', 'c').length, 1);
  assert.equal(doc.getElementsByTagNameNS('urn:two', 'c').length, 1);
  assert.equal(parseXml('<r xmlns="urn:one"><c xmlns=""/></r>').documentElement.children[0].namespaceURI, '');
});

test('XML byte input decodes UTF-8, UTF-16LE and UTF-16BE', () => {
  const source = '<r>中文😀</r>';
  assert.equal(parseXml(new TextEncoder().encode(source)).documentElement.textContent, '中文😀');
  assert.equal(parseXml(Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(source, 'utf16le')])).documentElement.textContent, '中文😀');
  const bigEndian = Buffer.from(source, 'utf16le');
  bigEndian.swap16();
  assert.equal(parseXml(Buffer.concat([Buffer.from([0xfe, 0xff]), bigEndian])).documentElement.textContent, '中文😀');
});

test('XML text and attribute newline normalization matches XML 1.0', () => {
  const doc = parseXml('<r a="a\r\nb\tc&#10;d">a\r\nb\rc</r>');
  assert.equal(attr(doc.documentElement, 'a'), 'a b c\nd');
  assert.equal(doc.documentElement.textContent, 'a\nb\nc');
});

test('XML rejects malformed structure, invalid entities and DTDs', () => {
  const malformed = [
    '', '<r>', '<r><c></r></c>', '<r/><extra/>', 'text<r/>', '<r/>text', '<r a=x/>', '<r a="x/>', '<r a="<"/>',
    '<r a="x"a2="y"/>', '<r a="x" a="y"/>', '<r>&custom;</r>', '<r>&amp</r>', '<r>&#0;</r>', '<r>&#xD800;</r>',
    '<r>&#x110000;</r>', '<r>]]></r>', '<r><!-- broken -- comment --></r>', '<r><!--x---></r>', '<r><![CDATA[x</r>',
    '<r/> <?xml version="1.0"?>', '<?xml version="1.1"?><r/>', '<?XML version="1.0"?><r/>', '<r xmlns:p=""><p:c/></r>',
    '<p:r/>', '<r p:a="1"/>', '<p:1 xmlns:p="urn:x"/>', '<r xmlns:p="urn:x" p:1="x"/>', '<r xmlns:a="urn:x" xmlns:b="urn:x" a:id="1" b:id="2"/>',
    '<r xmlns:xml="urn:wrong"/>', '<r xmlns="http://www.w3.org/XML/1998/namespace"/>', '<r xmlns:xmlns="urn:x"/>',
    '<!DOCTYPE r SYSTEM "https://example.invalid/a"><r/>', '<!DOCTYPE r [<!ENTITY x "abc">]><r>&x;</r>', '<r>\u0001</r>',
    '<r>\ud800</r>', '<r/ / >', '<r><c/></r', '<r a="1"',
  ];
  for (const source of malformed) assert.throws(() => parseXml(source), undefined, source);
});

test('XML depth and node limits fail explicitly and bounded valid trees remain searchable', () => {
  const bounded = parseXml('<a>'.repeat(128) + 'leaf' + '</a>'.repeat(128));
  assert.equal(bounded.documentElement.textContent, 'leaf');
  assert.equal(elements(bounded, 'a').length, 128);
  assert.throws(() => parseXml('<a>'.repeat(129) + '</a>'.repeat(129)), /128/);
  assert.throws(() => parseXml('<r>' + '<c/>'.repeat(200000) + '</r>'), /200000/);
});
