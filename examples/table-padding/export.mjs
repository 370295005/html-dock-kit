import { readFile, writeFile } from 'node:fs/promises';
import { asBuffer } from '../../dist/node.js';

const html = await readFile(new URL('./padding.html', import.meta.url), 'utf8');
const options = { pageSize: 'A4', properties: { title: '单元格内边距与垂直对齐' } };
await writeFile(new URL('./padding-fixed.docx', import.meta.url), asBuffer(html, options));
await writeFile(new URL('./padding-without-inlining.docx', import.meta.url),
  asBuffer(html, { ...options, inlineStyles: false }));
console.log('Generated padding-fixed.docx and padding-without-inlining.docx');
