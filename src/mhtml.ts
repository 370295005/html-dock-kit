import { encodeUtf8 } from './utf8.js';
import { defaultTreeAdapter, html as htmlNamespace, parse, serialize } from 'parse5';
import type { DefaultTreeAdapterTypes } from 'parse5';

type Node = DefaultTreeAdapterTypes.Node;
type Element = DefaultTreeAdapterTypes.Element;

interface EmbeddedImage {
  contentType: string;
  location: string;
  base64: string;
}

// A hyphen cannot occur in a base64 body, so this constant cannot collide with
// either the HTML or an image attachment. All MIME bodies are base64 encoded.
const BOUNDARY = '----html-docx-kit-mhtml-boundary';
const DOCUMENT_LOCATION = 'file:///C:/html-docx-kit/document.html';
const BASE64_ALPHABET = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
const MIME_TOKEN = /^[!#$%&'*+.^_`|~0-9A-Za-z-]+$/;
const MIME_EXTENSIONS: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/bmp': 'bmp',
  'image/x-ms-bmp': 'bmp',
  'image/tiff': 'tif',
  'image/svg+xml': 'svg',
  'image/webp': 'webp',
  'image/x-icon': 'ico',
  'image/vnd.microsoft.icon': 'ico',
  'image/avif': 'avif',
  'image/heic': 'heic',
  'image/heif': 'heif',
};

function isElement(node: Node): node is Element {
  return 'tagName' in node;
}

function allElements(root: Node): Element[] {
  const elements: Element[] = [];
  const pending: Node[] = [root];
  while (pending.length > 0) {
    const node = pending.pop();
    if (!node) continue;
    if (isElement(node)) elements.push(node);
    if ('childNodes' in node) {
      for (let index = node.childNodes.length - 1; index >= 0; index -= 1) {
        pending.push(node.childNodes[index]!);
      }
    }
    if (isElement(node) && node.tagName === 'template') {
      pending.push(defaultTreeAdapter.getTemplateContent(node as DefaultTreeAdapterTypes.Template));
    }
  }
  return elements;
}

function getAttribute(element: Element, name: string): string | undefined {
  return element.attrs.find(attribute => attribute.name === name && !attribute.namespace)?.value;
}

function normalizeCharset(document: DefaultTreeAdapterTypes.Document, elements: Element[]): void {
  for (const element of elements) {
    if (element.namespaceURI !== htmlNamespace.NS.HTML || element.tagName !== 'meta') continue;
    const charset = getAttribute(element, 'charset');
    const httpEquiv = getAttribute(element, 'http-equiv');
    if (charset !== undefined || httpEquiv?.trim().toLowerCase() === 'content-type') {
      defaultTreeAdapter.detachNode(element);
    }
  }
  const html = document.childNodes.find(node => isElement(node) && node.tagName === 'html');
  const head = html && isElement(html)
    ? html.childNodes.find(node => isElement(node) && node.tagName === 'head')
    : undefined;
  if (!head || !isElement(head)) throw new Error('Unable to normalize the HTML document head.');
  const meta = defaultTreeAdapter.createElement('meta', htmlNamespace.NS.HTML, [
    { name: 'charset', value: 'utf-8' },
  ]);
  const firstChild = head.childNodes[0];
  if (firstChild) defaultTreeAdapter.insertBefore(head, meta, firstChild);
  else defaultTreeAdapter.appendChild(head, meta);

  if (!document.childNodes.some(node => node.nodeName === '#documentType')) {
    defaultTreeAdapter.setDocumentType(document, 'html', '', '');
    const doctype = document.childNodes.pop();
    if (doctype) document.childNodes.unshift(doctype);
  }
}

function dataUrlError(detail: string): Error {
  return new Error(`Invalid image data URL: ${detail}`);
}

function percentDecode(payload: string): Uint8Array {
  // Decode percent escapes as bytes, rather than decodeURIComponent, which
  // incorrectly requires arbitrary binary image data to be valid UTF-8.
  const input = encodeUtf8(payload);
  const output = new Uint8Array(input.length);
  let length = 0;
  for (let index = 0; index < input.length; index += 1) {
    const byte = input[index]!;
    if (byte !== 0x25) {
      output[length++] = byte;
      continue;
    }
    const digits = payloadHex(input[index + 1], input[index + 2]);
    if (digits === undefined) throw dataUrlError('malformed percent escape.');
    output[length++] = digits;
    index += 2;
  }
  return output.subarray(0, length);
}

function payloadHex(first: number | undefined, second: number | undefined): number | undefined {
  const hexValue = (byte: number | undefined): number | undefined => {
    if (byte === undefined) return undefined;
    if (byte >= 0x30 && byte <= 0x39) return byte - 0x30;
    if (byte >= 0x41 && byte <= 0x46) return byte - 0x41 + 10;
    if (byte >= 0x61 && byte <= 0x66) return byte - 0x61 + 10;
    return undefined;
  };
  const high = hexValue(first);
  const low = hexValue(second);
  return high === undefined || low === undefined ? undefined : (high << 4) | low;
}

function decodeBase64(input: Uint8Array): Uint8Array {
  let encoded = '';
  for (const byte of input) {
    // ASCII whitespace is commonly present in exported base64 image sources.
    if (byte === 9 || byte === 10 || byte === 12 || byte === 13 || byte === 32) continue;
    encoded += String.fromCharCode(byte);
  }
  if (encoded.length === 0 || encoded.length % 4 !== 0 || !/^[A-Za-z0-9+/]+={0,2}$/.test(encoded)) {
    throw dataUrlError('base64 must have a valid alphabet and RFC 4648 padding.');
  }
  const padding = encoded.endsWith('==') ? 2 : encoded.endsWith('=') ? 1 : 0;
  const decoded = new Uint8Array((encoded.length / 4) * 3 - padding);
  let offset = 0;
  for (let index = 0; index < encoded.length; index += 4) {
    const first = BASE64_ALPHABET.indexOf(encoded[index]!);
    const second = BASE64_ALPHABET.indexOf(encoded[index + 1]!);
    const third = encoded[index + 2] === '=' ? 0 : BASE64_ALPHABET.indexOf(encoded[index + 2]!);
    const fourth = encoded[index + 3] === '=' ? 0 : BASE64_ALPHABET.indexOf(encoded[index + 3]!);
    if (index === encoded.length - 4 && ((padding === 2 && (second & 15) !== 0) || (padding === 1 && (third & 3) !== 0))) {
      throw dataUrlError('base64 has nonzero padding bits.');
    }
    decoded[offset++] = (first << 2) | (second >> 4);
    if (offset < decoded.length) decoded[offset++] = (second << 4) | (third >> 2);
    if (offset < decoded.length) decoded[offset++] = (third << 6) | fourth;
  }
  return decoded;
}

function encodeBase64(bytes: Uint8Array): string {
  const chunks: string[] = [];
  let chunk = '';
  for (let index = 0; index < bytes.length; index += 3) {
    const first = bytes[index]!;
    const second = bytes[index + 1];
    const third = bytes[index + 2];
    chunk += BASE64_ALPHABET[first >> 2]!;
    chunk += BASE64_ALPHABET[((first & 3) << 4) | ((second ?? 0) >> 4)]!;
    chunk += second === undefined ? '=' : BASE64_ALPHABET[((second & 15) << 2) | ((third ?? 0) >> 6)]!;
    chunk += third === undefined ? '=' : BASE64_ALPHABET[third & 63]!;
    if (chunk.length >= 16_384) {
      chunks.push(chunk);
      chunk = '';
    }
  }
  chunks.push(chunk);
  return chunks.join('');
}

function parseImageDataUrl(source: string): { contentType: string; mimeType: string; bytes: Uint8Array } {
  const comma = source.indexOf(',');
  if (comma < 0) throw dataUrlError('missing comma before the image payload.');
  const metadata = source.slice(5, comma).split(';');
  const mimeType = metadata.shift()?.toLowerCase() ?? '';
  const subtype = mimeType.slice('image/'.length);
  if (!mimeType.startsWith('image/') || !MIME_TOKEN.test(subtype)) {
    throw dataUrlError('the media type must be image/*, for example image/png.');
  }
  const base64 = metadata[metadata.length - 1]?.toLowerCase() === 'base64';
  if (base64) metadata.pop();
  const parameters = new Map<string, string>();
  for (const parameter of metadata) {
    const separator = parameter.indexOf('=');
    const name = parameter.slice(0, separator).toLowerCase();
    if (separator < 1 || !MIME_TOKEN.test(name) || parameters.has(name)) {
      throw dataUrlError('malformed or duplicate media type parameter; ;base64 must come last.');
    }
    let value: string;
    try {
      value = decodeURIComponent(parameter.slice(separator + 1));
    } catch {
      throw dataUrlError('malformed percent escape in a media type parameter.');
    }
    if (value.length === 0 || /[^\x20-\x7e]/.test(value)) {
      throw dataUrlError('media type parameters must have nonempty printable ASCII values.');
    }
    if (value.startsWith('"') && value.endsWith('"')) {
      value = value.slice(1, -1).replace(/\\([\x20-\x7e])/g, '$1');
    }
    if (name === 'charset') value = value.toLowerCase();
    parameters.set(name, MIME_TOKEN.test(value) ? value : `"${value.replace(/["\\]/g, '\\$&')}"`);
  }
  const contentType = mimeType + [...parameters].sort(([left], [right]) => left < right ? -1 : left > right ? 1 : 0)
    .map(([name, value]) => `; ${name}=${value}`).join('');
  const payload = percentDecode(source.slice(comma + 1));
  const bytes = base64 ? decodeBase64(payload) : payload;
  if (bytes.length === 0) throw dataUrlError('the image payload is empty.');
  return { contentType, mimeType, bytes };
}

function foldBase64(encoded: string): string {
  const lines: string[] = [];
  for (let index = 0; index < encoded.length; index += 76) lines.push(encoded.slice(index, index + 76));
  return lines.join('\r\n');
}

function mimePart(contentType: string, location: string, base64: string): string {
  return [
    `--${BOUNDARY}`,
    `Content-Type: ${contentType}`,
    'Content-Transfer-Encoding: base64',
    `Content-Location: ${location}`,
    '',
    foldBase64(base64),
    '',
  ].join('\r\n');
}

/** Package normalized HTML and inline images as a portable MHTML altChunk. */
export function createMhtml(html: string, externalImages: 'error' | 'preserve' = 'error'): string {
  const document = parse(html);
  const elements = allElements(document);
  normalizeCharset(document, elements);
  const images: EmbeddedImage[] = [];
  const embeddedByContent = new Map<string, EmbeddedImage>();
  for (const element of elements) {
    if (element.namespaceURI !== htmlNamespace.NS.HTML || element.tagName !== 'img') continue;
    if (getAttribute(element, 'srcset')?.trim()) {
      throw new Error('Image srcset is unsupported. Choose a single image and use an image/* data URL in img src.');
    }
    const sourceAttribute = element.attrs.find(attribute => attribute.name === 'src' && !attribute.namespace);
    if (!sourceAttribute) continue;
    const source = sourceAttribute.value.replace(/^[\t\n\f\r ]+|[\t\n\f\r ]+$/g, '');
    if (!source) continue;
    if (!/^data:/i.test(source)) {
      if (externalImages === 'preserve') continue;
      throw new Error('External img src cannot be embedded. Convert it to an image/* data URL, or set externalImages: "preserve". Images are never fetched.');
    }
    const { contentType, mimeType, bytes } = parseImageDataUrl(source);
    const base64 = encodeBase64(bytes);
    const key = `${contentType}\n${base64}`;
    let image = embeddedByContent.get(key);
    if (!image) {
      const extension = MIME_EXTENSIONS[mimeType] ?? 'img';
      image = {
        contentType,
        location: `file:///C:/html-docx-kit/image${images.length + 1}.${extension}`,
        base64,
      };
      embeddedByContent.set(key, image);
      images.push(image);
    }
    sourceAttribute.value = image.location;
  }
  const parts = [
    'MIME-Version: 1.0',
    `Content-Type: multipart/related; type="text/html"; boundary="${BOUNDARY}"`,
    '',
    mimePart('text/html; charset="utf-8"', DOCUMENT_LOCATION, encodeBase64(encodeUtf8(serialize(document)))),
    ...images.map(image => mimePart(image.contentType, image.location, image.base64)),
    `--${BOUNDARY}--`,
    '',
  ];
  return parts.join('\r\n');
}
