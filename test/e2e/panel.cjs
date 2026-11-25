// Buggtest i webbläsaren: panelen med låtsas-arbetsboken.
// Kör: npm run build && npx vite preview --mode browser --port 4174, sedan node test/e2e/panel.cjs
const { chromium } = require("playwright");
const SHOTS = process.env.SHOTS;
let failures = 0;
const check = (name, ok, detail = "") => {
  if (!ok) failures++;
  console.log(`${ok ? "OK  " : "FEL "} ${name}${detail ? ` — ${detail}` : ""}`);
};

(async () => {
  // UTF-8, annars döper Chromium om filer med å, ä och ö till "download".
  const browser = await chromium.launch({
    env: { ...process.env, LANG: "C.UTF-8", LC_ALL: "C.UTF-8" },
  });
  const page = await browser.newPage({ viewport: { width: 1400, height: 900 } });
  const errors = [];
  page.on("pageerror", (e) => errors.push(e.message));
  page.on(
    "console",
    (m) =>
      m.type() === "error" && !/appsforoffice|ERR_|net::/.test(m.text()) && errors.push(m.text()),
  );
  await page.goto(process.env.URL ?? "http://localhost:4174/");
  await page.waitForSelector("text=Uppdraget");
  // Svenska tusentalsavgränsare är hårda mellanslag: jämför med vanliga.
  const cell = async (r, c) =>
    (await page.locator(".grid tbody tr").nth(r).locator("td").nth(c).innerText()).replace(
      /[\u00a0\u202f]/g,
      " ",
    );

  check(
    "1 påminnelse om byråns namn",
    await page.isVisible("text=Börja med att skriva byråns namn"),
  );
  await page.click("role=tab[name='Byrån']");
  await page.fill("input[placeholder='Revisionsbyrån AB']", "Revisionsbyrån i Norr AB");
  await page.click("role=tab[name='Anpassa']");
  await page.fill("input[placeholder='Företaget AB']", "Kunden AB");
  await page.fill("input[placeholder='556677-8899']", "556677-8899");
  await page.fill("input[placeholder='2026-01-01 – 2026-12-31']", "2026-01-01 – 2026-12-31");
  await page.click("button:has-text('Anpassa ”Kundfordringar”')");
  await page.waitForSelector(".status.ok");
  check("2 huvudet: byrå", (await cell(0, 0)) === "Revisionsbyrån i Norr AB", await cell(0, 0));
  check("3 huvudet: klient", (await cell(1, 1)) === "Kunden AB");
  check("4 tabellen flyttad ner", (await cell(9, 0)) === "Kund");
  // Kolumnen har ett belopp med decimaler: då får hela kolumnen två decimaler.
  check("5 talformat", (await cell(10, 3)) === "48 250,00", await cell(10, 3));
  check("6 fakturanumret utan tusentalsavgränsare", (await cell(10, 1)) === "24117", await cell(10, 1));
  check(
    "7 status i listan",
    /Inte anpassad/.test(await page.locator(".sheets li").nth(1).innerText()),
  );

  // Signera utan initialer → begripligt fel.
  await page.click("button:has-text('Upprättat')");
  check("8 initialer saknas", /initialer/.test(await page.locator(".status").innerText()));
  await page.fill("input[placeholder='AB']", "AP");
  await page.click("button:has-text('Upprättat')");
  await page.waitForSelector(".status.ok");
  const today = new Date().toLocaleDateString("sv-SE");
  check(
    "9 upprättat: initialer och datum",
    (await cell(1, 6)) === "AP" && (await cell(1, 7)) === today,
    `${await cell(1, 6)} ${await cell(1, 7)}`,
  );
  check("10 listan visar U AP", /U AP/.test(await page.locator(".sheets li").first().innerText()));

  // Anpassa igen: ingen dubblering, initialerna kvar.
  await page.click("button:has-text('Anpassa ”Kundfordringar”')");
  await page.waitForTimeout(300);
  check("11 anpassa igen: tabellen på samma rad", (await cell(9, 0)) === "Kund");
  check("12 anpassa igen: initialerna kvar", (await cell(1, 6)) === "AP");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/mallverktyg-anpassad.png` });

  // Flera flikar på en gång.
  await page.locator(".sheets li").nth(1).locator("input").check();
  await page.locator(".sheets li").nth(2).locator("input").check();
  await page.click("button:has-text('Anpassa markerade (2)')");
  await page.waitForSelector("text=2 flikar är anpassade.");
  check(
    "13 två flikar anpassade",
    !/Inte anpassad/.test(await page.locator(".sheets").innerText()),
  );

  // Signera en flik som inte är anpassad: byt till en ny tom flik går inte här, så kontrollera texten för Bank.
  await page.click(".sheet-tabs button:has-text('Bank')");
  await page.waitForTimeout(700);
  check("14 aktiv flik följer med", await page.isVisible("button:has-text('Anpassa ”Bank”')"));

  // Mall: töm klientens uppgifter med bekräftelse.
  await page.click("role=tab[name='Mall']");
  await page.click("button:has-text('Töm klientens uppgifter')");
  await page.click("button:has-text('Ja, töm')");
  await page.waitForSelector("text=Klientens uppgifter är tömda.");
  await page.click(".sheet-tabs button:has-text('Kundfordringar')");
  check("15 klient och signering tömda", (await cell(1, 1)) === "" && (await cell(1, 6)) === "");
  check("16 byrån står kvar", (await cell(0, 0)) === "Revisionsbyrån i Norr AB");
  await page.click("button:has-text('Spara som mall (.xltx)')");
  check(
    "17 mall i webbläsarläget: förklaring",
    /bara i Excel/.test(await page.locator(".status").innerText()),
  );

  // Byrån: ändra färg och markeringar, dela som fil.
  await page.click("role=tab[name='Byrån']");
  if (SHOTS) await page.screenshot({ path: `${SHOTS}/mallverktyg-byran.png` });
  const [download] = await Promise.all([
    page.waitForEvent("download"),
    page.click("button:has-text('Spara som fil')"),
  ]);
  check(
    "18 inställningarna sparas som fil",
    /\.json$/.test(download.suggestedFilename()),
    download.suggestedFilename(),
  );

  check("Inga fel i konsolen", errors.length === 0, errors.join(" | "));
  console.log(failures ? `\n${failures} FEL` : "\nAllt gick igenom");
  await browser.close();
  process.exitCode = failures ? 1 : 0;
})();
