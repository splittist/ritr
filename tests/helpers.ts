import { strToU8, zipSync } from 'fflate';
import { W } from '../src/package/xml';

export const relNamespace = 'http://schemas.openxmlformats.org/package/2006/relationships';
export const officeRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
export const relationships = (body: string) =>
  `<Relationships xmlns="${relNamespace}">${body}</Relationships>`;
export const relation = (id: string, type: string, target: string, external = false) =>
  `<Relationship Id="${id}" Type="${officeRel}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`;
export const paragraph = (text: string) => `<w:p><w:r><w:t>${text}</w:t></w:r></w:p>`;
export function fixture(
  body = paragraph('Hello world'),
  extras: Record<string, string | Uint8Array> = {},
): Uint8Array {
  const parts: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
    '_rels/.rels': relationships(relation('rId1', 'officeDocument', 'word/document.xml')),
    'word/document.xml': `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="${W}" xmlns:x="urn:unknown" xmlns:r="${officeRel}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
    ...extras,
  };
  return zipSync(
    Object.fromEntries(
      Object.entries(parts).map(([name, content]) => [
        name,
        typeof content === 'string' ? strToU8(content) : content,
      ]),
    ),
  );
}
