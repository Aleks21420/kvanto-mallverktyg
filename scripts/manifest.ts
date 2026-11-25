// Gör Mallverktygets manifest av mallen manifest.template.xml.
//
//   node scripts/manifest.ts      skriver manifest.xml för utveckling (localhost)
//
// Det publicerade manifestet görs av scripts/install-files.ts vid bygget.
import { readFileSync, writeFileSync } from "node:fs";

/** Id för det publicerade Mallverktyget. Ändras aldrig, annars blir det ett nytt tillägg i Excel. */
export const PRODUCTION_ID = "8c511835-2eb9-4d7e-b149-c617fc608e15";
/** Id för utvecklingsversionen, så att den kan vara installerad samtidigt. */
export const DEVELOPMENT_ID = "61b43214-ac87-4e60-a224-abdecdd75ceb";

export const DEVELOPMENT_URL = "https://localhost:3100";

export function renderManifest(
  template: string,
  options: { id: string; name: string; url: string },
): string {
  const url = options.url.replace(/\/+$/, "");
  const values: Record<string, string> = {
    ID: options.id,
    NAME: escapeXml(options.name),
    URL: escapeXml(url),
    ORIGIN: escapeXml(new URL(url).origin),
  };
  const body = template.replace(
    /<!--[\s\S]*?-->/,
    "<!-- Gjord av manifest.template.xml med scripts/manifest.ts. Ändra i mallen. -->",
  );
  return body.replace(/\{\{(\w+)\}\}/g, (match, key: string) => {
    if (!(key in values)) throw new Error(`Okänd plats i manifestmallen: ${match}`);
    return values[key];
  });
}

function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

if (process.argv[1]?.endsWith("manifest.ts")) {
  const template = readFileSync("manifest.template.xml", "utf8");
  writeFileSync(
    "manifest.xml",
    renderManifest(template, {
      id: DEVELOPMENT_ID,
      name: "Kvanto Mallverktyg (utveckling)",
      url: DEVELOPMENT_URL,
    }),
  );
}
