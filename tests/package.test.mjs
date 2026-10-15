import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { createRequire } from 'node:module';
import { posix } from 'node:path';
import vm from 'node:vm';
import { unzipSync, strFromU8 } from 'fflate';
import { DOMParser } from '@xmldom/xmldom';
import { htmlToDocx, asBlob, DOCX_MIME_TYPE } from '../dist/index.js';
import { asBuffer } from '../dist/node.js';

const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_REL = 'http://schemas.openxmlformats.org/package/2006/relationships';
const CONTENT_TYPES = 'http://schemas.openxmlformats.org/package/2006/content-types';
const DC = 'http://purl.org/dc/elements/1.1/';
const HTML = '<!doctype html><html><head><meta charset="utf-8"><style>h1 { color: navy; }</style></head><body><h1>你好，Word 🌏</h1><p>A &amp; B</p></body></html>';

function equalBytes(actual, expected, message = 'DOCX bytes must match') {
  assert.equal(actual.byteLength, expected.byteLength, `${message}: length`);
  const digest = bytes => createHash('sha256').update(bytes).digest('hex');
  assert.equal(digest(actual), digest(expected), `${message}: SHA-256`);
}

function readXml(parts, name) {
  assert.ok(parts[name], `Missing package part: ${name}`);
  const failures = [];
  const xml = new DOMParser({
    onError(level, message) { failures.push(`${level}: ${message}`); },
  }).parseFromString(strFromU8(parts[name]), 'application/xml');
  assert.deepEqual(failures, [], `${name} must be well-formed XML`);
  assert.ok(xml.documentElement, `${name} needs a root element`);
  return xml;
}

function one(parent, namespace, name) {
  const matches = parent.getElementsByTagNameNS(namespace, name);
  assert.equal(matches.length, 1, `Expected one ${name}`);
  return matches.item(0);
}

function all(parent, namespace, name) {
  return Array.from(parent.getElementsByTagNameNS(namespace, name));
}

function packageFor(html = HTML, options) {
  return unzipSync(htmlToDocx(html, options));
}

function sectionFor(options) {
  const document = readXml(packageFor(HTML, options), 'word/document.xml');
  const section = one(document, WORD, 'sectPr');
  return {
    size: one(section, WORD, 'pgSz'),
    margins: one(section, WORD, 'pgMar'),
  };
}

function wordAttributes(element, expected) {
  for (const [name, value] of Object.entries(expected)) {
    assert.equal(element.getAttributeNS(WORD, name), String(value), `w:${name}`);
  }
}

test('DOCX contains well-formed, correctly linked OPC and WordprocessingML parts', () => {
  const bytes = htmlToDocx(HTML);
  assert.ok(bytes instanceof Uint8Array);
  assert.equal(bytes[0], 0x50);
  assert.equal(bytes[1], 0x4b);
  const parts = unzipSync(bytes);
  for (const name of [
    '[Content_Types].xml', '_rels/.rels', 'word/document.xml',
    'word/_rels/document.xml.rels', 'word/afchunk.mht', 'docProps/core.xml',
  ]) assert.ok(parts[name], name);
  for (const name of Object.keys(parts).filter(name => /\.(?:xml|rels)$/.test(name))) {
    readXml(parts, name);
  }

  const contentTypes = readXml(parts, '[Content_Types].xml');
  assert.equal(contentTypes.documentElement.namespaceURI, CONTENT_TYPES);
  const defaults = new Map(all(contentTypes, CONTENT_TYPES, 'Default')
    .map(part => [part.getAttribute('Extension'), part.getAttribute('ContentType')]));
  assert.equal(defaults.get('rels'), 'application/vnd.openxmlformats-package.relationships+xml');
  const overrides = new Map(all(contentTypes, CONTENT_TYPES, 'Override')
    .map(part => [part.getAttribute('PartName'), part.getAttribute('ContentType')]));
  assert.equal(overrides.get('/word/document.xml'),
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml');
  assert.equal(overrides.get('/word/afchunk.mht'), 'message/rfc822');
  assert.equal(overrides.get('/docProps/core.xml'), 'application/vnd.openxmlformats-package.core-properties+xml');
  for (const name of overrides.keys()) {
    assert.ok(name.startsWith('/'), `Override must use an absolute part name: ${name}`);
    assert.ok(parts[name.slice(1)], `Content type points to missing part: ${name}`);
  }
  for (const name of Object.keys(parts).filter(name => !name.endsWith('/') && name !== '[Content_Types].xml')) {
    assert.ok(overrides.has(`/${name}`) || defaults.has(name.split('.').at(-1)),
      `Part has no declared content type: ${name}`);
  }

  const roots = readXml(parts, '_rels/.rels');
  assert.equal(roots.documentElement.namespaceURI, PACKAGE_REL);
  const rootRelationships = all(roots, PACKAGE_REL, 'Relationship');
  const main = rootRelationships.find(rel => rel.getAttribute('Type') === `${OFFICE_REL}/officeDocument`);
  assert.ok(main, 'Package must link the main document');
  assert.equal(main.getAttribute('Target'), 'word/document.xml');
  assert.ok(rootRelationships.some(rel =>
    rel.getAttribute('Type') === `${PACKAGE_REL}/metadata/core-properties` &&
    rel.getAttribute('Target') === 'docProps/core.xml'));

  const document = readXml(parts, main.getAttribute('Target'));
  assert.equal(document.documentElement.namespaceURI, WORD);
  const body = one(document, WORD, 'body');
  const bodyElements = Array.from(body.childNodes).filter(node => node.nodeType === 1);
  assert.equal(bodyElements.at(-1).localName, 'sectPr', 'Section properties must be the final body element');
  const altChunk = one(body, WORD, 'altChunk');
  const id = altChunk.getAttributeNS(OFFICE_REL, 'id');
  assert.ok(id, 'altChunk must reference a relationship in the office relationship namespace');

  const relationships = readXml(parts, 'word/_rels/document.xml.rels');
  assert.equal(relationships.documentElement.namespaceURI, PACKAGE_REL);
  const documentRelationships = all(relationships, PACKAGE_REL, 'Relationship');
  const htmlRelationship = documentRelationships.find(rel => rel.getAttribute('Id') === id);
  assert.ok(htmlRelationship, 'altChunk relationship ID must resolve');
  assert.equal(htmlRelationship.getAttribute('Type'), `${OFFICE_REL}/aFChunk`,
    'The aFChunk relationship type is case sensitive');
  assert.notEqual(htmlRelationship.getAttribute('TargetMode'), 'External');
  assert.equal(posix.normalize(posix.join('word', htmlRelationship.getAttribute('Target'))), 'word/afchunk.mht');

  for (const [base, rels] of [['', rootRelationships], ['word', documentRelationships]]) {
    const ids = rels.map(rel => rel.getAttribute('Id'));
    assert.equal(new Set(ids).size, ids.length, 'Relationship IDs must be unique within a part');
    for (const rel of rels) {
      if (rel.getAttribute('TargetMode') === 'External') continue;
      const target = rel.getAttribute('Target');
      const name = target.startsWith('/') ? target.slice(1) : posix.normalize(posix.join(base, target));
      assert.ok(parts[name], `Relationship points to missing part: ${name}`);
    }
  }
  assert.match(strFromU8(parts['word/afchunk.mht']), /Content-Type: multipart\/related/i);
});

test('defaults use portrait Letter paper and one-inch content margins', () => {
  const { size, margins } = sectionFor();
  wordAttributes(size, { w: 12240, h: 15840, orient: 'portrait' });
  wordAttributes(margins, {
    top: 1440, right: 1440, bottom: 1440, left: 1440,
    header: 720, footer: 720, gutter: 0,
  });
});

test('A4 landscape and partial margins are represented in twips', () => {
  const { size, margins } = sectionFor({
    pageSize: 'A4', orientation: 'landscape', margins: { top: 720, left: 0, gutter: 100 },
  });
  wordAttributes(size, { w: 16838, h: 11906, orient: 'landscape' });
  wordAttributes(margins, {
    top: 720, right: 1440, bottom: 1440, left: 0,
    header: 720, footer: 720, gutter: 100,
  });
});

test('custom page dimensions respect orientation and allow explicit zero margins', () => {
  const options = {
    pageSize: { width: 9000, height: 6000 },
    margins: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 },
  };
  wordAttributes(sectionFor(options).size, { w: 6000, h: 9000, orient: 'portrait' });
  const landscape = sectionFor({ ...options, orientation: 'landscape' });
  wordAttributes(landscape.size, { w: 9000, h: 6000, orient: 'landscape' });
  wordAttributes(landscape.margins, options.margins);
});

test('invalid page, margin, compression, property, and image-policy options fail early', () => {
  const invalid = [
    [null, TypeError], [[], TypeError],
    [{ orientation: 'sideways' }, TypeError],
    [{ pageSize: 'A3' }, TypeError],
    [{ pageSize: { width: 0, height: 5000 } }, RangeError],
    [{ pageSize: { width: 5000.5, height: 6000 } }, RangeError],
    [{ pageSize: { width: Infinity, height: 6000 } }, RangeError],
    [{ pageSize: { width: 2147483648, height: 6000 } }, RangeError],
    [{ margins: null }, TypeError],
    [{ margins: { left: -1 } }, RangeError],
    [{ margins: { top: NaN } }, RangeError],
    [{ margins: { bottom: 0.5 } }, RangeError],
    [{ margins: { sideways: 1 } }, TypeError],
    [{ margins: { left: 6120, right: 6120 } }, RangeError],
    [{ margins: { top: 7920, bottom: 7920 } }, RangeError],
    [{ margins: { gutter: 10000 } }, RangeError],
    [{ compression: -1 }, RangeError],
    [{ compression: 10 }, RangeError],
    [{ compression: 1.5 }, RangeError],
    [{ compression: '6' }, RangeError],
    [{ externalImages: 'fetch' }, TypeError],
    [{ properties: null }, TypeError],
    [{ properties: { title: 42 } }, TypeError],
  ];
  for (const [options, error] of invalid) {
    assert.throws(() => htmlToDocx(HTML, options), error, JSON.stringify(options));
  }
});

test('non-string and empty HTML inputs are rejected', () => {
  for (const html of [undefined, null, 42, {}, '', ' \n\t ']) {
    assert.throws(() => htmlToDocx(html), TypeError);
  }
});

test('metadata preserves Unicode and XML-sensitive text without creating elements', () => {
  const properties = {
    title: '标题 <tag> & "quoted" \'single\' 🌏',
    creator: '张三 & 李四',
    description: '第一行\n第二行\t≤ 10 > 2',
  };
  const parts = packageFor(HTML, { properties });
  const core = readXml(parts, 'docProps/core.xml');
  assert.equal(core.documentElement.namespaceURI, 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties');
  for (const [name, value] of Object.entries(properties)) {
    const element = one(core, DC, name);
    assert.equal(element.textContent, value);
    assert.equal(Array.from(element.childNodes).filter(node => node.nodeType === 1).length, 0);
  }
  assert.equal(core.getElementsByTagName('tag').length, 0);
});

test('XML-forbidden metadata characters are rejected instead of corrupting the document', () => {
  for (const title of ['bad\u0000text', 'bad\u000btext', '\ud800', '\ufffe', '\uffff']) {
    assert.throws(() => htmlToDocx(HTML, { properties: { title } }), TypeError);
  }
});

test('same input produces deterministic ZIP bytes at each supported compression extreme', () => {
  for (const compression of [0, 6, 9]) {
    const options = { compression, properties: { title: '稳定输出' } };
    const first = htmlToDocx(HTML, options);
    const second = htmlToDocx(HTML, options);
    equalBytes(first, second);
    assert.deepEqual(Object.keys(unzipSync(first)).sort(), Object.keys(unzipSync(second)).sort());
  }
});

test('Blob and Node Buffer APIs contain exactly the same DOCX bytes', async () => {
  const options = { pageSize: 'A4', properties: { creator: '测试者' } };
  const bytes = htmlToDocx(HTML, options);
  const blob = asBlob(HTML, options);
  const buffer = asBuffer(HTML, options);
  assert.equal(DOCX_MIME_TYPE, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  assert.ok(blob instanceof Blob);
  assert.equal(blob.type, DOCX_MIME_TYPE);
  assert.equal(blob.size, bytes.length);
  assert.ok(Buffer.isBuffer(buffer));
  equalBytes(new Uint8Array(await blob.arrayBuffer()), bytes);
  equalBytes(new Uint8Array(buffer), bytes);
});

test('CommonJS package exports and the Node subpath are usable through require', async () => {
  const require = createRequire(import.meta.url);
  const common = require('html-docx-kit');
  const node = require('html-docx-kit/node');
  const options = { orientation: 'landscape', properties: { title: 'CommonJS' } };
  const expected = htmlToDocx(HTML, options);
  assert.equal(common.DOCX_MIME_TYPE, DOCX_MIME_TYPE);
  equalBytes(common.htmlToDocx(HTML, options), expected);
  equalBytes(new Uint8Array(await common.asBlob(HTML, options).arrayBuffer()), expected);
  assert.ok(Buffer.isBuffer(node.asBuffer(HTML, options)));
  equalBytes(new Uint8Array(node.asBuffer(HTML, options)), expected);
});

test('browser IIFE works without Node globals and produces matching Blob and ZIP bytes', async () => {
  const source = await readFile(new URL('../dist/html-docx-kit.global.js', import.meta.url), 'utf8');
  const context = vm.createContext({ Blob, atob });
  vm.runInContext(source, context, { filename: 'html-docx-kit.global.js', timeout: 10000 });
  assert.equal(vm.runInContext('typeof Buffer', context), 'undefined');
  assert.equal(vm.runInContext('typeof require', context), 'undefined');
  assert.equal(vm.runInContext('typeof process', context), 'undefined');
  assert.equal(typeof context.htmlDocx.htmlToDocx, 'function');
  const options = { pageSize: 'A4', properties: { title: '浏览器 🌏' } };
  const bytes = Uint8Array.from(context.htmlDocx.htmlToDocx(HTML, options));
  equalBytes(bytes, htmlToDocx(HTML, options), 'Browser IIFE and Node');
  readXml(unzipSync(bytes), 'word/document.xml');
  const blob = context.htmlDocx.asBlob(HTML, options);
  assert.equal(blob.type, DOCX_MIME_TYPE);
  equalBytes(new Uint8Array(await blob.arrayBuffer()), bytes);
});

test('browser byte output is available even when Blob is not provided', async () => {
  const source = await readFile(new URL('../dist/html-docx-kit.global.js', import.meta.url), 'utf8');
  const context = vm.createContext({ atob });
  vm.runInContext(source, context, { filename: 'html-docx-kit.global.js', timeout: 10000 });
  equalBytes(Uint8Array.from(context.htmlDocx.htmlToDocx(HTML)), htmlToDocx(HTML));
  assert.throws(() => context.htmlDocx.asBlob(HTML), /Blob is unavailable/);
});
