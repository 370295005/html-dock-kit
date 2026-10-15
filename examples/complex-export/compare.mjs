import { readFile, writeFile, mkdir } from 'node:fs/promises';
import vm from 'node:vm';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const directory = path.dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const value = (flag, fallback) => {
  const index = args.indexOf(flag);
  if (index < 0) return fallback;
  if (!args[index + 1]) throw new Error(`Missing value after ${flag}`);
  return path.resolve(args[index + 1]);
};
const reference = value('--reference');
if (!reference) {
  throw new Error('Pass --reference /path/to/html-docx-js/dist/html-docx.js (the original standalone browser bundle).');
}
const fixtures = value('--fixtures', directory);
const output = value('--outdir', path.join(directory, 'results'));
const newBundle = path.resolve(directory, '../../dist/html-docx-kit.global.js');

async function browserApi(filename) {
  // Both packages run in separate browser-like contexts with identical globals.
  const context = vm.createContext({ Blob, TextEncoder, TextDecoder, atob });
  context.window = context;
  context.self = context;
  vm.runInContext(await readFile(filename, 'utf8'), context, {
    filename: path.basename(filename), timeout: 10000,
  });
  if (typeof context.htmlDocx?.asBlob !== 'function') throw new Error(`No asBlob API in ${filename}`);
  return context.htmlDocx;
}

const [modern, legacy] = await Promise.all([browserApi(newBundle), browserApi(reference)]);
const modernVersion = JSON.parse(await readFile(path.resolve(directory, '../../package.json'), 'utf8')).version;
await mkdir(output, { recursive: true });
const manifest = [];
// Use equal Letter/portrait/one-inch settings to avoid page-size differences.
const options = {
  orientation: 'portrait',
  margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720, gutter: 0 },
};
for (const source of ['complex', 'boundary']) {
  const html = await readFile(path.join(fixtures, `${source}.html`), 'utf8');
  for (const [name, api] of [['new', modern], ['old', legacy]]) {
    const blob = api.asBlob(html, options);
    if (!(blob instanceof Blob)) throw new Error(`${name} did not return a Blob`);
    const filename = `${source}-${name}.docx`;
    const bytes = new Uint8Array(await blob.arrayBuffer());
    await writeFile(path.join(output, filename), bytes);
    manifest.push({ source: `${source}.html`, library: name === 'new' ? `html-docx-kit@${modernVersion}` : 'html-docx-js@0.3.1',
      filename, bytes: bytes.length, page: 'Letter portrait, 1440-twip content margins' });
  }
}
await writeFile(path.join(output, 'export-manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
console.log(JSON.stringify({ output, exports: manifest }, null, 2));
