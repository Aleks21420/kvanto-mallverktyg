# Kvanto Mallverktyg

Excel-tillägg som gör om en Excel-fil till byråns revisionspapper, med en knapp.

- **Huvud på varje flik:** byrå, klient, org.nr, räkenskapsår, referens, upprättat och granskat
  (initialer och datum), rubrik, syfte och utfört arbete. Uppdragets uppgifter fylls i en gång
  per arbetsbok.
- **Fasta avsnitt:** syfte och utfört arbete överst, slutsats och en förklaring av
  granskningsmarkeringarna (✓ ✗ Σ …) efter tabellen.
- **Byråns utseende:** typsnitt, talformat (nummerkolumner som fakturanummer lämnas orörda),
  tabellrubrik i byråns färg, kolumnbredder och utan stödlinjer.
- **Utskrift:** en sida bred, liggande för breda tabeller, sidfot med byrå, flik och sidnummer.
- **Signera:** knapparna Upprättat och Granskat fyller i initialer och dagens datum.
- **Mallar:** töm klientens uppgifter och spara arbetsboken som en Excel-mall (.xltx).
- **Byråns inställningar** delas med kollegorna som en fil.

Att anpassa en flik går att göra flera gånger: huvudet uppdateras, men det som är ifyllt står
kvar och inget läggs till två gånger.

## Installera

Se [INSTALLERA.md](INSTALLERA.md) (kommandot för `cmd`, med kopieringsknapp) eller
installationssidan `<adressen>/installera.html`.

## Utveckling

```bash
npm install
npm run dev:browser
```

- `npm test` kör testerna, `npm run build` bygger.
- Prova i riktiga Excel: `npm run certs` en gång, sedan `npm run dev` och lägg till `manifest.xml`.
