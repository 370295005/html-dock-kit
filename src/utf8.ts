/** Native UTF-8 when available, with a correct portable fallback for browsers/VMs. */
export function encodeUtf8(value: string): Uint8Array {
  if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(value);
  const bytes = new Uint8Array(value.length * 3);
  let offset = 0;
  for (const character of value) {
    let code = character.codePointAt(0)!;
    // Match TextEncoder: lone UTF-16 surrogates become the replacement character.
    if (code >= 0xd800 && code <= 0xdfff) code = 0xfffd;
    if (code <= 0x7f) bytes[offset++] = code;
    else if (code <= 0x7ff) {
      bytes[offset++] = 0xc0 | (code >> 6);
      bytes[offset++] = 0x80 | (code & 0x3f);
    } else if (code <= 0xffff) {
      bytes[offset++] = 0xe0 | (code >> 12);
      bytes[offset++] = 0x80 | ((code >> 6) & 0x3f);
      bytes[offset++] = 0x80 | (code & 0x3f);
    } else {
      bytes[offset++] = 0xf0 | (code >> 18);
      bytes[offset++] = 0x80 | ((code >> 12) & 0x3f);
      bytes[offset++] = 0x80 | ((code >> 6) & 0x3f);
      bytes[offset++] = 0x80 | (code & 0x3f);
    }
  }
  return bytes.subarray(0, offset);
}
