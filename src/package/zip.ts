import { unzipSync, zipSync } from 'fflate';

const MAX_ARCHIVE = 64 * 1024 * 1024;
const MAX_EXPANDED = 128 * 1024 * 1024;
const MAX_PART = 32 * 1024 * 1024;
const decoder = new TextDecoder('utf-8', { fatal: true });

const crcTable = Uint32Array.from({ length: 256 }, (_, n) => {
  for (let bit = 0; bit < 8; bit++) n = n & 1 ? 0xedb88320 ^ (n >>> 1) : n >>> 1;
  return n >>> 0;
});
function crc32(bytes: Uint8Array): number {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = crcTable[(crc ^ byte) & 255]! ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

/** Check central directory BEFORE inflation, including resource bounds and duplicates. */
export function readZip(bytes: Uint8Array): Map<string, Uint8Array> {
  if (bytes.length > MAX_ARCHIVE) throw new Error('ZIP exceeds the 64 MiB input limit');
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  for (; eocd >= Math.max(0, bytes.length - 65557); eocd--) {
    if (
      view.getUint32(eocd, true) === 0x06054b50 &&
      eocd + 22 + view.getUint16(eocd + 20, true) === bytes.length
    )
      break;
  }
  if (eocd < 0) throw new Error('Missing ZIP directory');
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const directoryEnd = offset + view.getUint32(eocd + 12, true);
  if (
    view.getUint16(eocd + 4, true) ||
    view.getUint16(eocd + 6, true) ||
    count !== view.getUint16(eocd + 8, true) ||
    count > 10000 ||
    directoryEnd !== eocd
  ) {
    throw new Error('Unsupported ZIP64, multi-disk, or oversized directory');
  }
  const expected = new Map<string, { size: number; crc: number }>();
  let total = 0;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > directoryEnd || view.getUint32(offset, true) !== 0x02014b50)
      throw new Error('Invalid ZIP entry');
    const flags = view.getUint16(offset + 8, true);
    const method = view.getUint16(offset + 10, true);
    const size = view.getUint32(offset + 24, true);
    const nameLength = view.getUint16(offset + 28, true);
    const next =
      offset +
      46 +
      nameLength +
      view.getUint16(offset + 30, true) +
      view.getUint16(offset + 32, true);
    if (next > directoryEnd) throw new Error('Truncated ZIP entry');
    const name = decoder.decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (
      !name ||
      name.startsWith('/') ||
      name.includes('\\') ||
      name.includes(':') ||
      name.includes('\0') ||
      name.split('/').some((s) => s === '..' || s === '.') ||
      expected.has(name)
    )
      throw new Error(`Unsafe or duplicate ZIP entry: ${name}`);
    total += size;
    if (flags & 1 || ![0, 8].includes(method) || size > MAX_PART || total > MAX_EXPANDED)
      throw new Error(`Unsupported or oversized ZIP part: ${name}`);
    const local = view.getUint32(offset + 42, true);
    if (local + 30 > offset || view.getUint32(local, true) !== 0x04034b50)
      throw new Error(`Invalid ZIP local header: ${name}`);
    const localNameLength = view.getUint16(local + 26, true);
    const dataStart = local + 30 + localNameLength + view.getUint16(local + 28, true);
    if (
      dataStart + view.getUint32(offset + 20, true) > view.getUint32(eocd + 16, true) ||
      decoder.decode(bytes.subarray(local + 30, local + 30 + localNameLength)) !== name ||
      view.getUint16(local + 8, true) !== method
    )
      throw new Error(`Inconsistent ZIP headers: ${name}`);
    expected.set(name, { size, crc: view.getUint32(offset + 16, true) });
    offset = next;
  }
  if (offset !== directoryEnd) throw new Error('ZIP directory size mismatch');
  const entries = unzipSync(bytes, {
    filter: (entry) => {
      const part = expected.get(entry.name);
      if (!part || part.size !== entry.originalSize) throw new Error('ZIP size mismatch');
      return true;
    },
  });
  const parts = new Map<string, Uint8Array>();
  for (const [name, spec] of expected) {
    const content = entries[name];
    if (!content || content.length !== spec.size || crc32(content) !== spec.crc)
      throw new Error(`ZIP integrity check failed: ${name}`);
    parts.set(name, content);
  }
  return parts;
}

export function writeZip(parts: ReadonlyMap<string, Uint8Array>): Uint8Array {
  return zipSync(Object.fromEntries(parts), { level: 6 });
}
export const equalBytes = (a: Uint8Array, b: Uint8Array) =>
  a.length === b.length && a.every((v, i) => v === b[i]);
