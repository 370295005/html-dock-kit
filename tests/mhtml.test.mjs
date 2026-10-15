import test from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync } from 'fflate';
import { htmlToDocx } from '../dist/index.js';

// Inspect the public API's actual archive and decode the MIME bodies with
// Node's independent base64 implementation, rather than library internals.
function readMhtml(html, options) {
  const archive = unzipSync(htmlToDocx(html, options));
  assert.ok(archive['word/afchunk.mht'], 'DOCX must contain its MHTML altChunk');
  const mhtml = Buffer.from(archive['word/afchunk.mht']).toString('ascii');
  assert.doesNotMatch(mhtml, /(?<!\r)\n|\r(?!\n)/, 'MIME must use CRLF throughout');
  assert.match(mhtml, /^MIME-Version: 1\.0\r\n/);
  const boundary = /boundary="([^"]+)"/.exec(mhtml)?.[1];
  assert.ok(boundary);
  const chunks = mhtml.split(`--${boundary}`);
  assert.equal(chunks.at(-1), '--\r\n', 'MIME closing delimiter must be present');
  const parts = chunks.slice(1, -1).map(chunk => {
    assert.ok(chunk.startsWith('\r\n'));
    const separator = chunk.indexOf('\r\n\r\n');
    assert.ok(separator > 0);
    const headers = chunk.slice(2, separator);
    const body = chunk.slice(separator + 4).trim();
    assert.match(headers, /Content-Transfer-Encoding: base64/);
    for (const line of body.split('\r\n')) {
      assert.ok(line.length <= 76, 'base64 MIME lines must fit 76 columns');
      assert.match(line, /^[A-Za-z0-9+/=]*$/);
    }
    const encoded = body.replaceAll('\r\n', '');
    const bytes = Buffer.from(encoded, 'base64');
    assert.equal(bytes.toString('base64'), encoded, 'MIME body must be canonical base64');
    return {
      contentType: /(?:^|\r\n)Content-Type: ([^\r]+)/.exec(headers)?.[1],
      location: /(?:^|\r\n)Content-Location: ([^\r]+)/.exec(headers)?.[1],
      bytes,
    };
  });
  assert.equal(parts[0]?.contentType, 'text/html; charset="utf-8"');
  return { mhtml, html: parts[0].bytes.toString('utf8'), images: parts.slice(1) };
}

test('fragments become full UTF-8 HTML and MIME preserves Unicode and long content', () => {
  const text = '你好，世界 🌍 café e\u0301 — العربية'.repeat(200);
  const result = readMhtml(`<h1>${text}</h1><p>&amp; &lt;文字&gt;</p>`);
  assert.match(result.html, /^<!DOCTYPE html><html><head><meta charset="utf-8"><\/head><body>/);
  assert.ok(result.html.includes(`<h1>${text}</h1>`));
  assert.ok(result.html.includes('<p>&amp; &lt;文字&gt;</p>'));
  assert.ok(result.html.endsWith('</body></html>'));
  assert.equal(result.images.length, 0);
});

test('all conflicting charset metadata is replaced with one first UTF-8 declaration', () => {
  const result = readMhtml(`<!DOCTYPE html><HTML><HEAD>
    <TITLE>中文</TITLE><META CHARSET=GB18030>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <meta name="description" content="keep this">
    </HEAD><BODY><meta charset="utf-16"><p>你好</p></BODY></HTML>`);
  assert.match(result.html, /<head><meta charset="utf-8">/);
  assert.equal((result.html.match(/<meta charset=/g) ?? []).length, 1);
  assert.doesNotMatch(result.html, /GB18030|iso-8859-1|utf-16|http-equiv/);
  assert.ok(result.html.includes('<meta name="description" content="keep this">'));
  assert.ok(result.html.includes('<title>中文</title>'));
});

test('actual img attributes handle case, quoting, entities, percent escapes and deduplication', () => {
  const result = readMhtml(`<IMG SRC='DATA:IMAGE/PNG;BASE64,/w=='>
    <img src=data:image/png,%FF>
    <img src="data&#58;image/png;base64,/&#119;==">
    <img src="data:image/png;base64,%2Fw%3D%3D">`);
  assert.equal(result.images.length, 1);
  assert.equal(result.images[0].contentType, 'image/png');
  assert.deepEqual(result.images[0].bytes, Buffer.from([255]));
  assert.equal(result.images[0].location, 'file:///C:/html-docx-kit/image1.png');
  assert.equal((result.html.match(/src="file:\/\/\/C:\/html-docx-kit\/image1\.png"/g) ?? []).length, 4);
  assert.doesNotMatch(result.html, /data:image/i);
});

test('percent-encoded binary payloads preserve non-UTF-8 bytes and MIME types deduplicate separately', () => {
  const result = readMhtml('<img src="data:image/png,%00%FF%80%FEABC"><img src="data:image/gif,%00%ff%80%feABC">');
  assert.equal(result.images.length, 2);
  assert.equal(result.images[0].location, 'file:///C:/html-docx-kit/image1.png');
  assert.equal(result.images[1].location, 'file:///C:/html-docx-kit/image2.gif');
  for (const image of result.images) assert.deepEqual(image.bytes, Buffer.from([0, 255, 128, 254, 65, 66, 67]));
});

test('MIME parameters and UTF-8 SVG data survive and equivalent parameters deduplicate', () => {
  const svg = '<svg xmlns="http://www.w3.org/2000/svg"><text>你好 🌟</text></svg>';
  const encoded = encodeURIComponent(svg);
  const result = readMhtml(`<img src="data:image/svg+xml;CHARSET=UTF-8;name=test%3Basset.svg,${encoded}">
    <img src="data:IMAGE/SVG+XML;name=test%3basset.svg;charset=utf-8,${encoded}">`);
  assert.equal(result.images.length, 1);
  assert.equal(result.images[0].contentType, 'image/svg+xml; charset=utf-8; name="test;asset.svg"');
  assert.equal(result.images[0].bytes.toString('utf8'), svg);
  assert.equal(result.images[0].location, 'file:///C:/html-docx-kit/image1.svg');
});

test('base64 whitespace is accepted and arbitrary byte values roundtrip', () => {
  const bytes = Buffer.from(Array.from({ length: 256 }, (_, index) => index));
  const encoded = bytes.toString('base64').replace(/.{40}/g, '$&\n \t');
  const result = readMhtml(`<img src="data:image/png;base64,${encoded}">`);
  assert.deepEqual(result.images[0].bytes, bytes);
});

test('scripts, styles, comments and attribute text are preserved without false image matches', () => {
  const script = `const markup = "<img src='https://example.com/script.png'>";`;
  const style = `.x::before { content: "<img src='https://example.com/style.png'>"; }`;
  const result = readMhtml(`<style>${style}</style><script>${script}</script>
    <!-- <img src="https://example.com/comment.png"> -->
    <p data-example="&lt;img src='https://example.com/attr.png'&gt;">example</p>
    <textarea><img src="https://example.com/textarea.png"></textarea>
    <template><img src="data:image/png;base64,AQ=="></template>`);
  assert.ok(result.html.includes(`<script>${script}</script>`));
  assert.ok(result.html.includes(`<style>${style}</style>`));
  assert.ok(result.html.includes('<!-- <img src="https://example.com/comment.png"> -->'));
  assert.ok(result.html.includes('https://example.com/attr.png'));
  assert.ok(result.html.includes('https://example.com/textarea.png'));
  assert.equal(result.images.length, 1);
  assert.ok(result.html.includes('<template><img src="file:///C:/html-docx-kit/image1.png"></template>'));
});

test('fixed MIME delimiter remains safe when its literal appears in HTML', () => {
  const result = readMhtml('<p>----html-docx-kit-mhtml-boundary</p>');
  assert.ok(result.html.includes('----html-docx-kit-mhtml-boundary'));
  assert.equal(result.images.length, 0);
});

test('external image sources reject by default and preserve mode never fetches them', () => {
  for (const source of ['https://example.invalid/image.png', '/images/example.png', './image.png', 'file:///tmp/image.png', 'blob:example']) {
    assert.throws(() => htmlToDocx(`<img src="${source}">`), /External img src.*image\/\* data URL.*externalImages.*never fetched/);
    const result = readMhtml(`<img src="${source}">`, { externalImages: 'preserve' });
    assert.ok(result.html.includes(`src="${source}"`));
    assert.equal(result.images.length, 0);
  }
});

test('missing and empty src remain unchanged, nonempty srcset has an actionable error', () => {
  const result = readMhtml('<img alt="empty"><img src=""><img src="   "><img srcset="">');
  assert.ok(result.html.includes('<img alt="empty"><img src=""><img src="   "><img srcset="">'));
  assert.equal(result.images.length, 0);
  for (const options of [undefined, { externalImages: 'preserve' }]) {
    assert.throws(() => htmlToDocx('<img src="data:image/png;base64,AQ==" srcset="https://example.invalid/a.png 2x">', options), /srcset is unsupported.*single image/);
    assert.throws(() => htmlToDocx('<img srcset="https://example.invalid/a.png 2x">', options), /srcset is unsupported/);
  }
});

test('invalid data URLs, malformed base64 and unsafe MIME parameters reject without exposing payloads', () => {
  const invalidSources = [
    'data:image/png;base64',
    'data:text/plain;base64,QQ==',
    'data:;base64,QQ==',
    'data:image/;base64,QQ==',
    'data:image/png,',
    'data:image/png,%',
    'data:image/png,%GG',
    'data:image/png;base64,',
    'data:image/png;base64,TQ',
    'data:image/png;base64,TQ=',
    'data:image/png;base64,TQ===',
    'data:image/png;base64,/w==AAAA',
    'data:image/png;base64,/x==',
    'data:image/png;base64,AAB=',
    'data:image/png;base64,@@@@',
    'data:image/png;base64,%FF',
    'data:image/png;base64;charset=utf-8,QQ==',
    'data:image/png;charset=utf-8;charset=UTF-8,ABC',
    'data:image/png;name=,ABC',
    'data:image/png;name=%0d%0aContent-Type%3Aevil,ABC',
    'data:image/png;name=%GG,ABC',
  ];
  for (const source of invalidSources) {
    assert.throws(() => htmlToDocx(`<img src="${source}">`), error => {
      assert.match(error.message, /^Invalid image data URL:/);
      assert.ok(!error.message.includes(source));
      return true;
    }, source);
  }
});
