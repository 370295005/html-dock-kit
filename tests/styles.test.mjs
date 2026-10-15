import test from 'node:test';
import assert from 'node:assert/strict';
import { unzipSync } from 'fflate';
import { parse } from 'parse5';
import { htmlToDocx } from '../dist/index.js';

// Assert against the public DOCX archive, decoding the MIME body independently
// with Node rather than importing the library's CSS or MHTML helpers.
function readHtml(source, options) {
  const archive = unzipSync(htmlToDocx(source, options));
  assert.ok(archive['word/afchunk.mht']);
  const mhtml = Buffer.from(archive['word/afchunk.mht']).toString('ascii');
  const boundary = /boundary="([^"]+)"/.exec(mhtml)?.[1];
  assert.ok(boundary);
  const firstPart = mhtml.split(`--${boundary}`)[1];
  const separator = firstPart.indexOf('\r\n\r\n');
  assert.ok(separator > 0);
  assert.match(firstPart.slice(0, separator), /Content-Type: text\/html; charset="utf-8"/);
  const html = Buffer.from(firstPart.slice(separator + 4).trim(), 'base64').toString('utf8');
  return { html, document: parse(html) };
}

function* nodes(root) {
  yield root;
  for (const child of root.childNodes ?? []) yield* nodes(child);
  if (root.content) yield* nodes(root.content);
}

function attribute(node, name) {
  return node.attrs?.find(attribute => attribute.name === name)?.value;
}

function byId(document, id) {
  const matches = [...nodes(document)].filter(node => attribute(node, 'id') === id);
  assert.equal(matches.length, 1, `Expected one element with id="${id}"`);
  return matches[0];
}

function declarations(node) {
  // These fixtures use simple declarations with no embedded semicolons.
  const result = new Map();
  for (const declaration of (attribute(node, 'style') ?? '').split(';')) {
    const separator = declaration.indexOf(':');
    if (separator < 0) continue;
    const name = declaration.slice(0, separator).trim().toLowerCase();
    const rawValue = declaration.slice(separator + 1).trim();
    const important = /\s*!important\s*$/i.test(rawValue);
    const value = rawValue.replace(/\s*!important\s*$/i, '').trim();
    result.set(name, { value, important });
  }
  return result;
}

function expectProperty(node, name, value, important = false) {
  assert.deepEqual(declarations(node).get(name), { value, important },
    `id="${attribute(node, 'id')}" ${name}`);
}

function text(node) {
  return [...nodes(node)].filter(node => node.nodeName === '#text').map(node => node.value).join('');
}

test('default inlining resolves class-table/small-table cell padding in stylesheet order', () => {
  const source = `<style>
    .class-table { border-collapse:collapse; font-size:9pt; }
    .class-table th, .class-table td { border:1pt solid #94a3b8; padding:6pt; }
  </style><style>
    .small-table { font-size:8pt; }
    .small-table td, .small-table th { padding:4pt; }
  </style>
  <table class="class-table"><tr><td id="normal">普通单元格</td></tr></table>
  <table class="class-table small-table" id="small-table"><tr><th id="small-header">标题</th><td id="small-cell">紧凑内容</td></tr></table>`;
  for (const options of [undefined, { inlineStyles: true }]) {
    const { document } = readHtml(source, options);
    expectProperty(byId(document, 'normal'), 'padding', '6pt');
    expectProperty(byId(document, 'small-header'), 'padding', '4pt');
    expectProperty(byId(document, 'small-cell'), 'padding', '4pt');
    expectProperty(byId(document, 'small-cell'), 'border', '1pt solid #94a3b8');
    expectProperty(byId(document, 'small-table'), 'font-size', '8pt');
  }
});

test('class descendant and child selectors match only their intended document scope', () => {
  const { document } = readHtml(`<style>
    .report .label { color:#2563eb; }
    .report > table td { padding:5pt; }
    .report table td .emphasis { font-weight:bold; }
  </style>
  <section class="report">
    <p class="label" id="inside-label">内部标签</p>
    <table><tr><td id="inside-cell"><span class="emphasis" id="inside-emphasis">重点</span></td></tr></table>
  </section>
  <p class="label" id="outside-label">外部标签</p>
  <table><tr><td id="outside-cell"><span class="emphasis" id="outside-emphasis">普通</span></td></tr></table>`);
  expectProperty(byId(document, 'inside-label'), 'color', '#2563eb');
  expectProperty(byId(document, 'inside-cell'), 'padding', '5pt');
  expectProperty(byId(document, 'inside-emphasis'), 'font-weight', 'bold');
  for (const id of ['outside-label', 'outside-cell', 'outside-emphasis']) {
    assert.equal(attribute(byId(document, id), 'style'), undefined, id);
  }
});

test('matched static class rules are copied inline without leaving duplicate rules for Word', () => {
  const page = '@page { size:Letter; margin:1in; }';
  const css = `/* 仅保留尚需 Word 处理的 CSS */\n.cell { padding:4pt; color:#123456; }\n${page}`;
  const { html, document } = readHtml(`<style id="original-style" type="text/css">${css}</style><p class="cell" id="cell">内容</p>`);
  const style = byId(document, 'original-style');
  assert.equal(style.tagName, 'style');
  assert.equal(attribute(style, 'type'), 'text/css');
  assert.doesNotMatch(text(style), /\.cell\s*\{/);
  assert.ok(text(style).includes(page), '@page must remain exactly as authored');
  assert.ok(html.includes(page));
  expectProperty(byId(document, 'cell'), 'padding', '4pt');
  expectProperty(byId(document, 'cell'), 'color', '#123456');
  assert.equal(attribute(byId(document, 'cell'), 'class'), 'cell');
});

test('ordinary inline declarations override ordinary stylesheet declarations', () => {
  const { document } = readHtml(`<style>.cell { color:#2563eb; padding:4pt; }</style>
    <p class="cell" id="cell" style="color:#166534; padding:2pt;">内联优先</p>`);
  expectProperty(byId(document, 'cell'), 'color', '#166534');
  expectProperty(byId(document, 'cell'), 'padding', '2pt');
});

test('stylesheet !important overrides ordinary inline declarations and remains important', () => {
  const { document } = readHtml(`<style>.cell { color:#be123c !important; padding:4pt !important; }</style>
    <p class="cell" id="cell" style="color:#166534; padding:2pt;">重要规则</p>`);
  expectProperty(byId(document, 'cell'), 'color', '#be123c', true);
  expectProperty(byId(document, 'cell'), 'padding', '4pt', true);
});

test('inline !important wins over stylesheet !important, including an ID selector', () => {
  const { document } = readHtml(`<style>#cell.cell { color:#be123c !important; padding:4pt !important; }</style>
    <p class="cell" id="cell" style="color:#166534 !important; padding:2pt !important;">内联重要规则</p>`);
  expectProperty(byId(document, 'cell'), 'color', '#166534', true);
  expectProperty(byId(document, 'cell'), 'padding', '2pt', true);
});

test('pre-existing important style attributes survive unrelated CSS inlining', () => {
  const { document } = readHtml(`<style>.cell { color:#334155; }</style>
    <p class="cell" id="cell" style="padding:3pt !important; margin-left:7pt !important; border-left:2pt solid #123456 !important;">保持声明</p>`);
  const cell = byId(document, 'cell');
  expectProperty(cell, 'padding', '3pt', true);
  expectProperty(cell, 'margin-left', '7pt', true);
  expectProperty(cell, 'border-left', '2pt solid #123456', true);
  expectProperty(cell, 'color', '#334155');
});

test('pseudo-element content is retained as CSS without inserting generated elements', () => {
  const css = `.marker::before { content:"BEFORE"; color:red; } .marker::after { content:"AFTER"; } .marker { color:#123456; }`;
  const { document } = readHtml(`<style id="css">${css}</style><p class="marker" id="marker">只有原文</p>`);
  const marker = byId(document, 'marker');
  assert.equal(text(marker), '只有原文');
  assert.equal(marker.childNodes.filter(node => node.tagName).length, 0);
  const preservedCss = text(byId(document, 'css'));
  assert.match(preservedCss, /\.marker::before\s*\{/);
  assert.match(preservedCss, /content:\s*["']BEFORE["']/);
  assert.match(preservedCss, /\.marker::after\s*\{/);
  assert.match(preservedCss, /content:\s*["']AFTER["']/);
  assert.doesNotMatch(preservedCss, /\.marker\s*\{/,
    'The matched ordinary rule must not be reinterpreted by Word');
  expectProperty(marker, 'color', '#123456');
});

test('global body and paragraph rules remain available to Word after class rules are inlined', () => {
  const { document } = readHtml(`<!doctype html><html><head><style id="css">
    body { font-family:Arial; color:#334155; }
    p { margin:0 0 8pt; }
    .cell { padding:4pt; }
  </style></head><body id="body"><p class="cell" id="cell">正文</p></body></html>`);
  const preservedCss = text(byId(document, 'css'));
  assert.match(preservedCss, /body\s*\{/);
  assert.match(preservedCss, /p\s*\{/);
  assert.doesNotMatch(preservedCss, /\.cell\s*\{/);
  expectProperty(byId(document, 'body'), 'font-family', 'Arial');
  expectProperty(byId(document, 'body'), 'color', '#334155');
  expectProperty(byId(document, 'cell'), 'margin', '0 0 8pt');
  expectProperty(byId(document, 'cell'), 'padding', '4pt');
});

test('data-embed and non-CSS style elements retain their source and are skipped', () => {
  const embeddedCss = '.live { padding:70pt !important; color:#ff0000 !important; }';
  const nonCss = '.live { padding:80pt !important; color:#be123c !important; }';
  const { document } = readHtml(`<style>.live { padding:4pt; }</style>
    <style id="embedded" data-embed>${embeddedCss}</style>
    <style id="non-css" type="text/plain">${nonCss}</style>
    <p class="live" id="live">真实内容</p>`);
  const live = byId(document, 'live');
  expectProperty(live, 'padding', '4pt');
  assert.equal(declarations(live).has('color'), false);
  const embedded = byId(document, 'embedded');
  assert.equal(attribute(embedded, 'data-embed'), '');
  assert.equal(text(embedded), embeddedCss);
  const other = byId(document, 'non-css');
  assert.equal(attribute(other, 'type'), 'text/plain');
  assert.equal(text(other), nonCss);
});

for (const [name, block] of [
  ['media', '@media screen and (min-width:600px) {\n  /* keep media */\n  .live { color:#ff0000; padding:50pt !important; }\n}'],
  ['container', '@container report (min-width:10px) {\n  .live { color:#ff0000; padding:60pt !important; }\n}'],
  ['layer', '@layer utilities {\n  .live { color:#ff0000; padding:70pt !important; }\n}'],
]) {
  test(`@${name} blocks remain byte-for-byte and their nested rules are not inlined`, () => {
    const { document } = readHtml(`<style id="css">${block}\n.live { padding:4pt; }</style>
      <p class="live" id="live">真实内容</p>`);
    const preservedCss = text(byId(document, 'css'));
    assert.ok(preservedCss.includes(block), `@${name} block must remain exactly as authored`);
    // Only the rule inside the preserved at-block remains; its sibling static
    // class rule has already been copied to the element and removed.
    assert.equal((preservedCss.match(/\.live\s*\{/g) ?? []).length, 1);
    const live = byId(document, 'live');
    expectProperty(live, 'padding', '4pt');
    assert.equal(declarations(live).has('color'), false);
  });
}

test('inlineStyles:false preserves class-only markup and original inline attributes', () => {
  const css = '.class-table td { padding:4pt; color:#2563eb; }';
  const { document } = readHtml(`<style id="css">${css}</style>
    <table class="class-table"><tr><td id="class-only">仅 class</td><td id="inline" style="padding:2pt; color:#166534;">已有内联</td></tr></table>`, { inlineStyles: false });
  assert.equal(attribute(byId(document, 'class-only'), 'style'), undefined);
  assert.equal(attribute(byId(document, 'inline'), 'style'), 'padding:2pt; color:#166534;');
  assert.equal(text(byId(document, 'css')), css);
  const table = [...nodes(document)].find(node => node.tagName === 'table');
  assert.equal(attribute(table, 'class'), 'class-table');
});

test('inlineStyles rejects non-boolean values instead of silently coercing them', () => {
  for (const inlineStyles of [null, 0, 1, NaN, 'true', 'false', {}, []]) {
    assert.throws(() => htmlToDocx('<p>内容</p>', { inlineStyles }), TypeError);
  }
});

test('CSS-shaped text in templates, scripts, comments and textareas is inert', () => {
  const script = `const sample = "<style>.live { padding:80pt !important; color:#ff0000 !important; }</style>";`;
  const comment = '<style>.live { padding:90pt !important; color:#ff0000 !important; }</style>';
  const { html, document } = readHtml(`<style>.live { padding:4pt; }</style>
    <template id="template"><style>.live { padding:70pt !important; color:#ff0000 !important; }</style></template>
    <script>${script}</script>
    <!-- ${comment} -->
    <textarea><style>.live { padding:100pt !important; color:#ff0000 !important; }</style></textarea>
    <p class="live" id="live">真实内容</p>`);
  const live = byId(document, 'live');
  expectProperty(live, 'padding', '4pt');
  assert.equal(declarations(live).has('color'), false);
  assert.ok(html.includes(`<script>${script}</script>`));
  assert.ok(html.includes(`<!-- ${comment} -->`));
  assert.ok([...nodes(byId(document, 'template'))].some(node => node.tagName === 'style'));
});

test('multiple classes and multiple tables retain independent cascade results', () => {
  const { document } = readHtml(`<style>
    .class-table td { padding:6pt; color:#111111; }
    .class-table.compact td { padding:4pt; }
    .class-table.alert td { background-color:#ffedd5; }
    .class-table .special { color:#123456; }
    .outside .special { color:#654321; }
  </style>
  <table class="class-table compact alert"><tr><td class="special" id="compact">紧凑提示</td></tr></table>
  <table class="class-table"><tr><td id="normal">普通表格</td></tr></table>
  <table class="compact alert"><tr><td class="special" id="unscoped">非目标表格</td></tr></table>
  <section class="outside"><p class="special" id="outside">外部内容</p></section>`);
  const compact = byId(document, 'compact');
  expectProperty(compact, 'padding', '4pt');
  expectProperty(compact, 'background-color', '#ffedd5');
  expectProperty(compact, 'color', '#123456');
  const normal = byId(document, 'normal');
  expectProperty(normal, 'padding', '6pt');
  expectProperty(normal, 'color', '#111111');
  assert.equal(declarations(normal).has('background-color'), false);
  assert.equal(attribute(byId(document, 'unscoped'), 'style'), undefined);
  expectProperty(byId(document, 'outside'), 'color', '#654321');
});

test('inlining leaves HTML table presentation and image dimensions unchanged', () => {
  const { document } = readHtml(`<style>
    .table { width:100%; background-color:#ffedd5; }
    .table td { text-align:center; padding:4pt; }
    .image { width:132px; height:44px; }
  </style><table class="table" id="table" cellspacing="0" cellpadding="0"><tr><td id="cell">内容</td></tr></table>
    <img class="image" id="image" src="data:image/png;base64,AQ==" width="132" height="44">`);
  const table = byId(document, 'table');
  assert.equal(attribute(table, 'cellspacing'), '0');
  assert.equal(attribute(table, 'cellpadding'), '0');
  assert.equal(attribute(table, 'width'), undefined);
  assert.equal(attribute(table, 'bgcolor'), undefined);
  assert.equal(attribute(byId(document, 'cell'), 'align'), undefined);
  const image = byId(document, 'image');
  assert.equal(attribute(image, 'width'), '132');
  assert.equal(attribute(image, 'height'), '44');
  expectProperty(image, 'width', '132px');
  expectProperty(image, 'height', '44px');
});
