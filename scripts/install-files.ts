// Körs efter bygget: lägger det publicerade manifestet i dist, så att
// installationssidan och Microsoft 365-administrationen kan hämta det.
import { readFileSync, writeFileSync } from "node:fs";
import { PRODUCTION_ID, renderManifest } from "./manifest.ts";

export const PUBLIC_URL: string = JSON.parse(readFileSync("mallverktyg.json", "utf8")).url;

if (process.argv[1]?.endsWith("install-files.ts")) {
  const template = readFileSync("manifest.template.xml", "utf8");
  writeFileSync(
    "dist/mallverktyg-manifest.xml",
    renderManifest(template, { id: PRODUCTION_ID, name: "Kvanto Mallverktyg", url: PUBLIC_URL }),
  );
  console.log(`dist/mallverktyg-manifest.xml (${PUBLIC_URL})`);
}
