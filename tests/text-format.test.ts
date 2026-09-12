import { test } from 'node:test';
import assert from 'node:assert/strict';
import { textStyle, highlightColors } from '../src/ui/text-format';
import { Workspace } from '../src/engine/workspace';
import { spans } from '../src/engine/document';
import { fixture, relationships, relation } from './helpers';
import { W } from '../src/package/xml';

test('display cascade respects explicit off, none, and automatic color without changing font metrics', () => {
  const inherited = { b: 'on', i: '1', u: 'single', color: '123456', highlight: 'yellow' };
  const styled = textStyle({ inherited, direct: { rFonts: 'Arial', sz: '72' } });
  assert.match(styled, /font-weight:700;font-style:italic/);
  assert.match(styled, /text-decoration-line:underline/);
  assert.match(styled, /color:#123456;background-color:#ffff00/);
  assert.doesNotMatch(styled, /font-family|font-size/);
  const cleared = textStyle({
    inherited,
    direct: { b: '0', i: 'false', u: 'none', color: 'auto', highlight: 'none' },
  });
  assert.match(cleared, /font-weight:400;font-style:normal;text-decoration-line:none/);
  assert.match(cleared, /color:inherit;background-color:transparent/);
});

test('highlight uses only the OOXML palette and never treats shading or arbitrary strings as CSS', () => {
  assert.equal(Object.keys(highlightColors).length, 16);
  assert.match(
    textStyle({ inherited: {}, direct: { highlight: 'darkYellow' } }),
    /background-color:#808000/,
  );
  for (const value of ['rebeccapurple', '#ff0000', 'red;position:fixed', 'constructor']) {
    const style = textStyle({
      inherited: {},
      direct: { highlight: value, color: value, u: value, shd: 'yellow' },
    });
    assert.match(style, /text-decoration-line:none/);
    assert.match(style, /color:inherit;background-color:transparent/);
  }
});

test('bold and italic toggle through style inheritance while direct formatting is absolute', () => {
  const bytes = fixture(
    '<w:p><w:pPr><w:pStyle w:val="Derived"/></w:pPr><w:r><w:t>Style</w:t></w:r><w:r><w:rPr><w:b/><w:i w:val="0"/></w:rPr><w:t>Direct</w:t></w:r></w:p>',
    {
      'word/_rels/document.xml.rels': relationships(relation('styles', 'styles', 'styles.xml')),
      'word/styles.xml': `<w:styles xmlns:w="${W}"><w:style w:styleId="Base"><w:rPr><w:b/><w:i/></w:rPr></w:style><w:style w:styleId="Derived"><w:basedOn w:val="Base"/><w:rPr><w:b/><w:i w:val="false"/></w:rPr></w:style></w:styles>`,
    },
  );
  const w = new Workspace(),
    id = w.open('styles.docx', bytes);
  const [style, direct] = spans(w.document(id).model);
  assert.match(textStyle(style!), /font-weight:400;font-style:italic/);
  assert.match(textStyle(direct!), /font-weight:700;font-style:normal/);
  assert.deepEqual(w.package(id).save(), bytes);
});
