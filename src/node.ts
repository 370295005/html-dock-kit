import { Buffer } from 'node:buffer';
import { htmlToDocx } from './index.js';
import type { HtmlDocxOptions } from './types.js';

export * from './index.js';

/** Node-only entry point; the main/browser entry never imports Node built-ins. */
export function asBuffer(html: string, options: HtmlDocxOptions = {}): Buffer {
  return Buffer.from(htmlToDocx(html, options));
}
