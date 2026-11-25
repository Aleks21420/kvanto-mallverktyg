// Gör en arbetsbok (.xlsx) till en Excel-mall (.xltx). Skillnaden är en enda
// rad i filen [Content_Types].xml, som talar om att arbetsboken är en mall.
// En xlsx-fil är en zip-fil: övriga delar kopieras som de är, utan att packas upp.

const WORKBOOK = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
const TEMPLATE = "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
const MACRO_WORKBOOK = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";

interface Entry {
  name: Uint8Array;
  method: number;
  time: number;
  date: number;
  crc: number;
  data: Uint8Array;
  size: number;
}

/** Läser zip-filens innehållsförteckning (i slutet av filen). */
function readZip(bytes: Uint8Array): Entry[] {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50) {
      end = i;
      break;
    }
  }
  if (end < 0) throw new Error("Filen är inte en Excel-arbetsbok.");
  const count = view.getUint16(end + 10, true);
  let at = view.getUint32(end + 16, true);
  const entries: Entry[] = [];
  for (let i = 0; i < count; i++) {
    if (view.getUint32(at, true) !== 0x02014b50) throw new Error("Arbetsboken går inte att läsa.");
    const method = view.getUint16(at + 10, true);
    const time = view.getUint16(at + 12, true);
    const date = view.getUint16(at + 14, true);
    const crc = view.getUint32(at + 16, true);
    const packed = view.getUint32(at + 20, true);
    const size = view.getUint32(at + 24, true);
    const nameLength = view.getUint16(at + 28, true);
    const extraLength = view.getUint16(at + 30, true);
    const commentLength = view.getUint16(at + 32, true);
    const offset = view.getUint32(at + 42, true);
    const name = bytes.slice(at + 46, at + 46 + nameLength);
    // Datat börjar efter den lokala rubriken, vars längd kan skilja sig från förteckningens.
    const localName = view.getUint16(offset + 26, true);
    const localExtra = view.getUint16(offset + 28, true);
    const start = offset + 30 + localName + localExtra;
    entries.push({ name, method, time, date, crc, data: bytes.slice(start, start + packed), size });
    at += 46 + nameLength + extraLength + commentLength;
  }
  return entries;
}

function writeZip(entries: Entry[]): Uint8Array {
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;
  for (const e of entries) {
    const local = new DataView(new ArrayBuffer(30));
    local.setUint32(0, 0x04034b50, true);
    local.setUint16(4, 20, true);
    local.setUint16(6, 0x0800, true);
    local.setUint16(8, e.method, true);
    local.setUint16(10, e.time, true);
    local.setUint16(12, e.date, true);
    local.setUint32(14, e.crc, true);
    local.setUint32(18, e.data.length, true);
    local.setUint32(22, e.size, true);
    local.setUint16(26, e.name.length, true);
    parts.push(new Uint8Array(local.buffer), e.name, e.data);

    const entry = new DataView(new ArrayBuffer(46));
    entry.setUint32(0, 0x02014b50, true);
    entry.setUint16(4, 20, true);
    entry.setUint16(6, 20, true);
    entry.setUint16(8, 0x0800, true);
    entry.setUint16(10, e.method, true);
    entry.setUint16(12, e.time, true);
    entry.setUint16(14, e.date, true);
    entry.setUint32(16, e.crc, true);
    entry.setUint32(20, e.data.length, true);
    entry.setUint32(24, e.size, true);
    entry.setUint16(28, e.name.length, true);
    entry.setUint32(42, offset, true);
    central.push(new Uint8Array(entry.buffer), e.name);
    offset += 30 + e.name.length + e.data.length;
  }
  const size = central.reduce((s, p) => s + p.length, 0);
  const end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true);
  end.setUint16(8, entries.length, true);
  end.setUint16(10, entries.length, true);
  end.setUint32(12, size, true);
  end.setUint32(16, offset, true);
  const all = [...parts, ...central, new Uint8Array(end.buffer)];
  const out = new Uint8Array(all.reduce((s, p) => s + p.length, 0));
  let at = 0;
  for (const p of all) {
    out.set(p, at);
    at += p.length;
  }
  return out;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[n] = c >>> 0;
  }
  return table;
})();

export function crc32(data: Uint8Array): number {
  let crc = 0xffffffff;
  for (const b of data) crc = CRC_TABLE[(crc ^ b) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

async function inflate(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as BlobPart])
    .stream()
    .pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/** Arbetsboken som mall. Arbetsböcker med makron går inte (då behövs .xltm). */
export async function toTemplate(xlsx: Uint8Array): Promise<Uint8Array> {
  const entries = readZip(xlsx);
  const decoder = new TextDecoder();
  const types = entries.find((e) => decoder.decode(e.name) === "[Content_Types].xml");
  if (!types) throw new Error("Filen är inte en Excel-arbetsbok.");
  const raw = types.method === 8 ? await inflate(types.data) : types.data;
  const xml = decoder.decode(raw);
  if (xml.includes(MACRO_WORKBOOK)) {
    throw new Error("Arbetsboken innehåller makron och kan inte sparas som en vanlig mall.");
  }
  if (!xml.includes(WORKBOOK) && !xml.includes(TEMPLATE)) {
    throw new Error("Filen är inte en Excel-arbetsbok.");
  }
  const data = new TextEncoder().encode(xml.replace(WORKBOOK, TEMPLATE));
  Object.assign(types, { method: 0, data, size: data.length, crc: crc32(data) });
  return writeZip(entries);
}
