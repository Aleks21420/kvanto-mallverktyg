# Projekt

Kvanto Mallverktyg: Excel-tillägg som gör om Excel-filer till byråns revisionspapper (huvud,
signering, fasta avsnitt, byråns utseende, utskrift och mallar). För revisorer och
redovisningskonsulter på små byråer i Sverige. Syskon till Kvanto (repot dokumenttolkning) och
Kvanto Transaktionsanalys.

# Principer

- Arbetsböckerna läses och ändras bara i Excel på användarens dator och skickas ALDRIG till någon
  server. Säkerhetspolicyn (scripts/csp.ts) spärrar anrop till andra adresser; försvaga den aldrig.
- Användarna är revisorer, inte tekniker: vardagsord på svenska, inga tekniska felmeddelanden.
- Att anpassa en flik ska gå att göra flera gånger: inget läggs till två gånger och det som är
  ifyllt skrivs aldrig över.
- Enkel, läsbar kod. Inga onödiga beroenden (bara React).
- Konkurrenters namn och produktbenämningar används inte i kod, texter eller marknadsföring.

# Teknik

- Office Add-in (Office.js, ExcelApi 1.7; utskrift kräver 1.9), React, TypeScript, Vite.
- src/lib/layout.ts räknar ut en plan (lista med steg) för en flik; src/excel/excelHost.ts utför
  den i Excel och src/excel/memory.ts i låtsas-Excel (tester och webbläsarläget).
- Flikens rutor hittas med namn per flik (Mall_Huvud, Mall_Referens, Mall_UpprattatAv …).
  Uppdragets uppgifter står på det dolda bladet Mallverktyg, med namnen Mall_Klient m.fl.
  Namnen får inte bytas: då hittas rutorna inte i befintliga arbetsböcker.
- Byråns inställningar (src/lib/profile.ts) sparas i datorn och delas som en .json-fil.
- Spara som mall: arbetsboken hämtas från Excel och görs om till .xltx (src/lib/template.ts).
- Knapparna på fliken Mallverktyg anropar panelens kod via delad körmiljö (src/ribbon.ts).

# Utveckling

- `npm test` kör testerna, `npm run build` typkontrollerar, bygger och lägger det publicerade
  manifestet i dist.
- `npm run dev:browser` startar panelen i webbläsaren, bredvid en låtsas-arbetsbok.
- `npm run certs` en gång och sedan `npm run dev` för att prova i riktiga Excel (manifest.xml).
- Manifest: ändra i `manifest.template.xml`. Det publicerade id:t (scripts/manifest.ts) får
  aldrig ändras.
- Buggtest i webbläsaren: `npm run build`, `npx vite preview --mode browser --port 4174`, sedan
  `node test/e2e/panel.cjs`.
- Publicering: varje push till main publiceras på Cloudflare (.github/workflows/publicera.yml);
  adressen står i mallverktyg.json. Ikoner: `npm run ikoner`.
