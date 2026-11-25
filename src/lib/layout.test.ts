import { test } from "node:test";
import assert from "node:assert/strict";
import { MemoryWorkbook } from "../excel/memory.ts";
import {
  HEADER_ROWS,
  columnName,
  findTableHeader,
  numberFormats,
  parseCell,
  planSheet,
} from "./layout.ts";
import { DEFAULT_PROFILE, cleanProfile, tint } from "./profile.ts";

const OPTIONS = { style: true, print: true };

/** En flik som de flesta revisionspapper: en rubrik och en tabell med tal. */
function workbook() {
  const book = new MemoryWorkbook();
  book.firm = "Revisionsbyrån AB";
  book.engagement = { client: "Kunden AB", orgNr: "556677-8899", year: "2026-01-01 – 2026-12-31" };
  const sheet = book.addSheet("Kundfordringar");
  sheet.set("A1", [
    ["Kund", "Fakturanr", "Belopp", "Betalt"],
    ["Alfa AB", 1001, 12500, 12500],
    ["Beta AB", 1002, 8400.5, 0],
    ["Gamma AB", 1003, 31000, 31000],
  ]);
  sheet.set("C6", [[51900.5]]);
  return { book, sheet };
}

test("hjälpfunktioner", () => {
  assert.equal(columnName(0), "A");
  assert.equal(columnName(27), "AB");
  assert.deepEqual(parseCell("Blad1!$C$4"), { row: 3, col: 2 });
  assert.equal(tint("#1F4E79", 0.85), "#DDE4EB");
  assert.equal(
    findTableHeader([
      ["Kund", "Belopp"],
      ["A", 1],
    ]),
    0,
  );
  assert.equal(
    findTableHeader([
      ["Kund", "Belopp"],
      ["A", "B"],
    ]),
    null,
  );
  assert.deepEqual(
    numberFormats(
      [
        [1, 2.5, "x"],
        [3, 4, 5],
      ],
      [
        ["General", "General", "General"],
        ["0%", "General", "General"],
      ],
      "auto",
    ),
    [
      ["#,##0", "#,##0.00", "General"],
      ["0%", "#,##0.00", "#,##0"],
    ],
  );
});

test("talformat: nummerkolumner lämnas som de är", () => {
  const values = [
    ["Konto", "Fakturanr", "År", "Belopp", "Antal"],
    [1930, 24117, 2026, 1500, 3],
  ];
  const formats = values.map((r) => r.map(() => "General"));
  assert.deepEqual(numberFormats(values, formats, "auto", 0)[1], [
    "General",
    "General",
    "General",
    "#,##0",
    "#,##0",
  ]);
});

test("en flik får huvud, utseende, slutsats och markeringar", () => {
  const { sheet } = workbook();
  sheet.apply(
    planSheet(sheet.snapshot(), { ...DEFAULT_PROFILE, firm: "Revisionsbyrån AB" }, OPTIONS),
  );
  const H = HEADER_ROWS;
  // Huvudet.
  assert.equal(sheet.display(0, 0), "Revisionsbyrån AB");
  assert.equal(sheet.display(1, 1), "Kunden AB");
  assert.equal(sheet.display(2, 1), "556677-8899");
  assert.equal(sheet.display(3, 1), "2026-01-01 – 2026-12-31");
  assert.equal(sheet.display(4, 0), "Kundfordringar");
  assert.equal(sheet.display(5, 0), "Syfte");
  assert.equal(sheet.names.get("Mall_Referens"), "G1");
  assert.equal(sheet.display(0, 5), "Referens");
  // Innehållet flyttat ner, med talformat och tabellrubrik.
  assert.equal(sheet.display(H, 0), "Kund");
  assert.equal(sheet.get(H, 0)!.style.bold, true);
  assert.equal(sheet.get(H + 1, 2)!.format, "#,##0.00");
  // Fakturanumret är ett nummer, inget belopp: inget talformat.
  assert.equal(sheet.get(H + 1, 1)!.format, "General");
  // Slutsats två rader under tabellen (sista raden är summan på rad 6 → H + 5).
  assert.equal(sheet.names.get("Mall_Slutsats"), `B${H + 5 + 3}`);
  assert.equal(sheet.display(H + 7, 0), "Slutsats");
  assert.equal(sheet.display(H + 9, 0), "Markeringar");
  assert.equal(sheet.display(H + 10, 0), "✓");
  assert.equal(sheet.gridlines, false);
  assert.equal(sheet.print?.landscape, false);
  assert.equal(sheet.print?.rightFooter, "Sida &P av &N");
});

test("anpassa igen: inget läggs till två gånger och det ifyllda står kvar", () => {
  const { sheet } = workbook();
  sheet.apply(planSheet(sheet.snapshot(), DEFAULT_PROFILE, OPTIONS));
  const before = sheet.text();
  sheet.set(sheet.names.get("Mall_Referens")!, [["B1"]]);
  sheet.set(sheet.names.get("Mall_Syfte")!, [["Att kundfordringarna finns"]]);
  sheet.set(sheet.names.get("Mall_Slutsats")!, [["Utan anmärkning"]]);
  // Byrån byter rubrik på ett avsnitt och tar bort en markering.
  const profile = cleanProfile({
    ...DEFAULT_PROFILE,
    sections: { ...DEFAULT_PROFILE.sections, work: "Granskningsåtgärder" },
    markings: DEFAULT_PROFILE.markings.slice(0, 2),
  });
  sheet.apply(planSheet(sheet.snapshot(), profile, OPTIONS));
  const H = HEADER_ROWS;
  assert.equal(sheet.display(H, 0), "Kund", "innehållet har inte flyttats");
  assert.equal(sheet.display(6, 0), "Granskningsåtgärder");
  assert.equal(sheet.display(0, 6), "B1");
  assert.equal(sheet.display(5, 1), "Att kundfordringarna finns");
  assert.equal(sheet.display(H + 7, 1), "Utan anmärkning");
  assert.equal(sheet.display(H + 12, 0), "", "borttagen markering är borta");
  assert.equal(sheet.display(H + 11, 0), "✗");
  assert.equal(sheet.text().split("\n").length, before.split("\n").length - 3);
});

test("tom flik och bred tabell", () => {
  const book = new MemoryWorkbook();
  const empty = book.addSheet("Tom");
  empty.apply(planSheet(empty.snapshot(), DEFAULT_PROFILE, OPTIONS));
  assert.equal(empty.names.get("Mall_Slutsats"), `B${HEADER_ROWS + 1}`);

  const wide = book.addSheet("Bred");
  wide.set("A1", [
    Array.from({ length: 12 }, (_, i) => `K${i}`),
    Array.from({ length: 12 }, (_, i) => i),
  ]);
  wide.apply(planSheet(wide.snapshot(), DEFAULT_PROFILE, OPTIONS));
  // Referensen till höger om tabellen, och liggande utskrift.
  assert.equal(wide.names.get("Mall_Referens"), "K1");
  assert.equal(wide.print?.landscape, true);
});

test("byråns inställningar: felaktiga värden ersätts med standard", () => {
  const p = cleanProfile({
    accent: "röd",
    fontSize: 99,
    markings: [{ symbol: "" }, { symbol: "✓", meaning: "OK" }],
    numbers: "x",
  });
  assert.equal(p.accent, DEFAULT_PROFILE.accent);
  assert.equal(p.fontSize, 11);
  assert.deepEqual(p.markings, [{ symbol: "✓", meaning: "OK" }]);
  assert.equal(p.numbers, "auto");
});
