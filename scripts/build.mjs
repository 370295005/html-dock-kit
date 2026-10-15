import { build } from 'esbuild';
import { execFileSync } from 'node:child_process';
import { readdir, readFile, writeFile, rm } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { join } from 'node:path';

const root = fileURLToPath(new URL('../', import.meta.url));
process.chdir(root);
await rm('dist', { recursive: true, force: true });
const common = {
  bundle: true, target: 'es2020', legalComments: 'eof', sourcemap: true,
  logLevel: 'info', metafile: true,
  // Juice's synchronous client only needs HTML parsing, not Cheerio's Node
  // network/stream entry. Imported Buffer in the Node API is unaffected.
  alias: { cheerio: 'cheerio/slim' },
  define: { Buffer: 'undefined', process: 'undefined' },
};
const builds = await Promise.all([
  build({ ...common, entryPoints: ['src/index.ts'], outfile: 'dist/index.js', format: 'esm', platform: 'browser' }),
  build({ ...common, entryPoints: ['src/index.ts'], outfile: 'dist/index.cjs', format: 'cjs', platform: 'browser' }),
  build({ ...common, entryPoints: ['src/node.ts'], outfile: 'dist/node.js', format: 'esm', platform: 'node' }),
  build({ ...common, entryPoints: ['src/node.ts'], outfile: 'dist/node.cjs', format: 'cjs', platform: 'node' }),
  build({ ...common, sourcemap: false, minify: true, entryPoints: ['src/index.ts'], outfile: 'dist/html-docx-kit.global.js', format: 'iife', platform: 'browser', globalName: 'htmlDocx' }),
]);
execFileSync(process.execPath, ['node_modules/typescript/bin/tsc', '-p', 'tsconfig.json'], { stdio: 'inherit' });
// Pair declarations with each module format for NodeNext consumers.
for (const name of await readdir('dist')) {
  if (name.endsWith('.d.ts')) {
    const declaration = await readFile(join('dist', name), 'utf8');
    await writeFile(join('dist', name.replace(/\.d\.ts$/, '.d.cts')),
      declaration.replace(/(from\s+['"][^'"]+)\.js(['"])/g, '$1.cjs$2'));
  }
}
const notices = [];
// Include every package that actually contributed code, including nested
// versions. Dev-only resource downloaders are not part of the shipped bundle.
const directories = new Set();
for (const result of builds) {
  for (const input of Object.keys(result.metafile.inputs)) {
    const absolute = `${root}${input}`;
    const index = absolute.lastIndexOf('/node_modules/');
    if (index < 0) continue;
    const prefix = absolute.slice(0, index + '/node_modules/'.length);
    const segments = absolute.slice(prefix.length).split('/');
    const name = segments[0].startsWith('@') ? segments.slice(0, 2).join('/') : segments[0];
    directories.add(`${prefix}${name}`);
  }
}
const packages = new Map();
for (const directory of directories) {
  const pkg = JSON.parse(await readFile(`${directory}/package.json`, 'utf8'));
  let license;
  const names = await readdir(directory);
  for (const filename of names.filter(name => /^licen[cs]e(?:\.(?:txt|md))?$/i.test(name))) {
    try { license = await readFile(`${directory}/${filename}`, 'utf8'); break; } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  if (!license) {
    try { license = await readFile(`scripts/licenses/${pkg.name}.txt`, 'utf8'); } catch (error) {
      if (error.code !== 'ENOENT') throw error;
    }
  }
  if (!license) throw new Error(`Missing bundled dependency license: ${pkg.name}`);
  packages.set(`${pkg.name}@${pkg.version}`, license);
}
for (const [name, license] of [...packages].sort(([a], [b]) => a.localeCompare(b))) {
  notices.push(`${name}\n${'='.repeat(60)}\n${license.trim()}\n`);
}
await writeFile('THIRD_PARTY_NOTICES.txt', notices.join('\n'));
