import { readFile, writeFile } from 'node:fs/promises';
import { asBuffer } from '../dist/node.js';

const source = new URL('./sample.html', import.meta.url);
const target = new URL('./sample.docx', import.meta.url);
const html = await readFile(source, 'utf8');
await writeFile(target, asBuffer(html, {
  pageSize: 'A4',
  margins: { top: 1000, bottom: 1000 },
  properties: { title: 'HTML 导出示例', creator: 'html-docx-kit' },
}));
console.log(`Created ${target.pathname}`);
