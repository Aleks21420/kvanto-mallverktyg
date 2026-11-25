// INSTALLERA.md visar samma kommandon som installationssidan.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
// @ts-expect-error installationssidans skript är vanlig JavaScript utan typer.
import { ADDIN_ID, installCommand, uninstallCommand } from "../public/installera.js";
import { PRODUCTION_ID } from "./manifest.ts";

test("INSTALLERA.md och installationssidan har samma kommandon och rätt id", () => {
  const url: string = JSON.parse(
    readFileSync(new URL("../mallverktyg.json", import.meta.url), "utf8"),
  ).url;
  const text = readFileSync(new URL("../INSTALLERA.md", import.meta.url), "utf8");
  assert.equal(ADDIN_ID, PRODUCTION_ID);
  assert.ok(text.includes(installCommand(url)));
  assert.ok(text.includes(uninstallCommand()));
  assert.ok(text.includes(`${url}/mallverktyg-manifest.xml`));
});
