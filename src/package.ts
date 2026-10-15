import { encodeUtf8 } from './utf8.js';
import type { ResolvedOptions } from './options.js';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function escapeXml(value: string): string {
  for (const character of value) {
    const code = character.codePointAt(0)!;
    if ((code < 32 && code !== 9 && code !== 10 && code !== 13) ||
        (code >= 0xd800 && code <= 0xdfff) || code === 0xfffe || code === 0xffff) {
      throw new TypeError('Document properties contain a character not allowed in XML 1.0.');
    }
  }
  return value.replace(/[&<>"']/g, character => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;',
  })[character]!);
}

export function createParts(mhtml: string, options: ResolvedOptions): Record<string, Uint8Array> {
  const { width, height, orientation, margins: m, properties: p } = options;
  const parts: Record<string, string> = {
    '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/afchunk.mht" ContentType="message/rfc822"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/></Types>`,
    '_rels/.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDocument" Type="${REL}/officeDocument" Target="word/document.xml"/><Relationship Id="rIdCore" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/></Relationships>`,
    'word/document.xml': `${XML}<w:document xmlns:w="${WORD}" xmlns:r="${REL}"><w:body><w:altChunk r:id="rIdHtml"/><w:sectPr><w:pgSz w:w="${width}" w:h="${height}" w:orient="${orientation}"/><w:pgMar w:top="${m.top}" w:right="${m.right}" w:bottom="${m.bottom}" w:left="${m.left}" w:header="${m.header}" w:footer="${m.footer}" w:gutter="${m.gutter}"/></w:sectPr></w:body></w:document>`,
    'word/_rels/document.xml.rels': `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHtml" Type="${REL}/aFChunk" Target="afchunk.mht" TargetMode="Internal"/></Relationships>`,
    'word/afchunk.mht': mhtml,
    'docProps/core.xml': `${XML}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>${escapeXml(p.title)}</dc:title><dc:creator>${escapeXml(p.creator)}</dc:creator><dc:description>${escapeXml(p.description)}</dc:description></cp:coreProperties>`,
  };
  return Object.fromEntries(Object.entries(parts).map(([name, content]) => [name, encodeUtf8(content)]));
}
