import { zipSync } from 'fflate';
import { createMhtml } from './mhtml.js';
import { resolveOptions } from './options.js';
import { createParts } from './package.js';
import { inlineStyles } from './styles.js';
import type { HtmlDocxOptions } from './types.js';

export type { HtmlDocxOptions, Margins, PageSize, DocumentProperties } from './types.js';

export const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

/** Create a DOCX package synchronously. Word imports the embedded HTML when opened. */
export function htmlToDocx(html: string, options: HtmlDocxOptions = {}): Uint8Array {
  if (typeof html !== 'string') throw new TypeError('html must be a string.');
  if (html.trim().length === 0) throw new TypeError('html must not be empty.');
  const resolved = resolveOptions(options);
  const styledHtml = resolved.inlineStyles ? inlineStyles(html) : html;
  const mhtml = createMhtml(styledHtml, resolved.externalImages);
  return zipSync(createParts(mhtml, resolved), {
    level: resolved.compression,
    // ZIP stores local wall-clock values. A fixed local date gives stable bytes across time zones.
    mtime: new Date(2000, 0, 1, 0, 0, 0),
  });
}

/** Return a Blob in browsers and modern Node, consistently in both environments. */
export function asBlob(html: string, options: HtmlDocxOptions = {}): Blob {
  if (typeof Blob === 'undefined') throw new Error('Blob is unavailable; use htmlToDocx() to get a Uint8Array.');
  const bytes = htmlToDocx(html, options);
  const buffer = new ArrayBuffer(bytes.byteLength);
  new Uint8Array(buffer).set(bytes);
  return new Blob([buffer], { type: DOCX_MIME_TYPE });
}

export default { htmlToDocx, asBlob };
