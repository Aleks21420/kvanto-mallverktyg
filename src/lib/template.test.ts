import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { inflateRawSync } from "node:zlib";
import { toTemplate } from "./template.ts";

/** Filerna i en zip-fil, uppackade (Node:s egen uppackning, för att kontrollera). */
function unzip(bytes: Uint8Array): Map<string, string> {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = bytes.length - 22;
  while (view.getUint32(end, true) !== 0x06054b50) end--;
  let at = view.getUint32(end + 16, true);
  const files = new Map<string, string>();
  for (let i = 0; i < view.getUint16(end + 10, true); i++) {
    const method = view.getUint16(at + 10, true);
    const packed = view.getUint32(at + 20, true);
    const n = view.getUint16(at + 28, true);
    const skip = n + view.getUint16(at + 30, true) + view.getUint16(at + 32, true);
    const name = new TextDecoder().decode(bytes.slice(at + 46, at + 46 + n));
    const offset = view.getUint32(at + 42, true);
    const start =
      offset + 30 + view.getUint16(offset + 26, true) + view.getUint16(offset + 28, true);
    const data = bytes.slice(start, start + packed);
    files.set(name, new TextDecoder().decode(method === 8 ? inflateRawSync(data) : data));
    at += 46 + skip;
  }
  return files;
}

test("arbetsboken blir en mall, med allt annat orört", async () => {
  const xlsx = new Uint8Array(
    readFileSync(new URL("../../test/fixtures/arbetspapper.xlsx", import.meta.url)),
  );
  const before = unzip(xlsx);
  const after = unzip(await toTemplate(xlsx));
  assert.match(after.get("[Content_Types].xml")!, /spreadsheetml\.template\.main\+xml/);
  assert.doesNotMatch(after.get("[Content_Types].xml")!, /spreadsheetml\.sheet\.main\+xml/);
  assert.deepEqual([...after.keys()], [...before.keys()]);
  for (const [name, text] of before) {
    if (name !== "[Content_Types].xml") assert.equal(after.get(name), text, name);
  }
});

test("annat än arbetsböcker stoppas med ett begripligt fel", async () => {
  await assert.rejects(
    toTemplate(new TextEncoder().encode("inte en zip-fil")),
    /inte en Excel-arbetsbok/,
  );
});
