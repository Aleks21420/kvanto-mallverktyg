// Låtsas-Excel för att prova panelen i en vanlig webbläsare (npm run dev:browser):
// en arbetsbok i minnet med några typiska arbetspapper.

import {
  CLIENT_FIELDS,
  NAMES,
  excelToday,
  parseCell,
  planSheet,
  type Options,
} from "../lib/layout.ts";
import type { Profile } from "../lib/profile.ts";
import { dateText, type Host, type SheetStatus, type SignKind } from "./host.ts";
import { MemoryWorkbook } from "./memory.ts";

export function exampleWorkbook(): MemoryWorkbook {
  const book = new MemoryWorkbook();
  book.addSheet("Kundfordringar").set("A1", [
    ["Kund", "Fakturanr", "Fakturadatum", "Belopp", "Betalt efter bokslut", "Kommentar"],
    ["Byggfirman Ek AB", 24117, "2026-11-28", 48250, 48250, ""],
    ["Lindqvist Konsult", 24121, "2026-12-03", 12875.5, 12875.5, ""],
    ["Norrvik Fastigheter", 24130, "2026-12-15", 96400, 0, "Påminnelse skickad"],
    ["Café Solsidan", 24138, "2026-12-22", 4310, 4310, ""],
    ["Summa", "", "", 161835.5, 65435.5, ""],
  ]);
  book.addSheet("Bank").set("A1", [
    ["Konto", "Enligt bokföringen", "Enligt kontoutdrag", "Differens"],
    ["1930 Företagskonto", 512780.5, 512780.5, 0],
    ["1940 Placeringskonto", 250000, 250000, 0],
  ]);
  book
    .addSheet("Anteckningar")
    .set("A1", [["Möte med VD 2027-01-14: inga kända tvister. Lager inventerat 2026-12-31."]]);
  return book;
}

export function createBrowserHost(
  book = exampleWorkbook(),
): Host & { book: MemoryWorkbook; activate(index: number): void } {
  const listeners = new Set<() => void>();
  const changed = () => listeners.forEach((l) => l());

  return {
    kind: "browser",
    book,
    activate(index: number) {
      book.active = index;
      changed();
    },
    async getEngagement() {
      return { ...book.engagement };
    },
    async setEngagement(engagement, profile) {
      book.engagement = { ...engagement };
      book.firm = profile.firm;
      changed();
    },
    async sheets() {
      return book.sheets.map((sheet): SheetStatus => {
        const at = (name: string) => {
          const address = sheet.names.get(name);
          if (!address) return "";
          const { row, col } = parseCell(address)!;
          return sheet.display(row, col);
        };
        return {
          name: sheet.name,
          adapted: sheet.names.has(NAMES.header),
          reference: String(at(NAMES.reference)),
          prepared: { by: String(at(NAMES.preparedBy)), date: dateText(at(NAMES.preparedDate)) },
          reviewed: { by: String(at(NAMES.reviewedBy)), date: dateText(at(NAMES.reviewedDate)) },
        };
      });
    },
    async activeSheet() {
      return book.sheets[book.active].name;
    },
    async adapt(name: string, profile: Profile, options: Options) {
      book.firm = profile.firm;
      const sheet = book.sheet(name);
      sheet.apply(planSheet(sheet.snapshot(), profile, options));
      changed();
    },
    async sign(name: string, kind: SignKind, initials: string) {
      const sheet = book.sheet(name);
      const [by, date] =
        kind === "prepared"
          ? [NAMES.preparedBy, NAMES.preparedDate]
          : [NAMES.reviewedBy, NAMES.reviewedDate];
      const byAddress = sheet.names.get(by);
      const dateAddress = sheet.names.get(date);
      if (!byAddress || !dateAddress) return false;
      sheet.set(byAddress, [[initials]]);
      sheet.set(dateAddress, [[excelToday()]]);
      changed();
      return true;
    },
    async clearClientData() {
      book.engagement = { client: "", orgNr: "", year: "" };
      for (const sheet of book.sheets) {
        for (const name of CLIENT_FIELDS) {
          const address = sheet.names.get(name);
          if (address) sheet.set(address, [[""]]);
        }
      }
      changed();
    },
    async file() {
      return null;
    },
    onChange(listener) {
      listeners.add(listener);
      return () => void listeners.delete(listener);
    },
  };
}
