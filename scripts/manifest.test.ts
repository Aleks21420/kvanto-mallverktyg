import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { DEVELOPMENT_ID, PRODUCTION_ID, renderManifest } from "./manifest.ts";

const template = readFileSync(new URL("../manifest.template.xml", import.meta.url), "utf8");

test("manifestet fylls i och varje knapp har texter och ikoner", () => {
  const xml = renderManifest(template, {
    id: PRODUCTION_ID,
    name: "Kvanto Mallverktyg",
    url: "https://exempel.se/",
  });
  assert.ok(!xml.includes("{{"));
  assert.ok(xml.includes(`<Id>${PRODUCTION_ID}</Id>`));
  assert.ok(xml.includes('DefaultValue="https://exempel.se/index.html"'));
  // Allt som pekas ut med resid finns bland resurserna.
  const defined = new Set([...xml.matchAll(/ id="([^"]+)" DefaultValue/g)].map((m) => m[1]));
  for (const [, id] of xml.matchAll(/resid="([^"]+)"/g)) assert.ok(defined.has(id), id);
  // Funktionerna som knapparna anropar finns i src/ribbon.ts.
  const ribbon = readFileSync(new URL("../src/ribbon.ts", import.meta.url), "utf8");
  for (const [, name] of xml.matchAll(/<FunctionName>(\w+)<\/FunctionName>/g)) {
    assert.ok(ribbon.includes(`${name}:`), name);
  }
  assert.notEqual(PRODUCTION_ID, DEVELOPMENT_ID);
});

test("ikonerna som manifestet pekar på finns", () => {
  const xml = renderManifest(template, { id: PRODUCTION_ID, name: "x", url: "https://exempel.se" });
  for (const [, path] of xml.matchAll(/https:\/\/exempel\.se\/(icons\/[^"]+)"/g)) {
    assert.ok(readFileSync(new URL(`../public/${path}`, import.meta.url)).length > 0, path);
  }
});
