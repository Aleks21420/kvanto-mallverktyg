// Riktiga Excel (Office.js): läser flikarna och utför planens steg.

import {
  CLIENT_FIELDS,
  ENGAGEMENT_NAMES,
  ENGAGEMENT_SHEET,
  NAMES,
  cellsInArea,
  excelToday,
  planSheet,
  type Options,
  type Snapshot,
  type Step,
  type Value,
} from "../lib/layout.ts";
import type { Profile } from "../lib/profile.ts";
import { dateText, type Engagement, type Host, type SheetStatus, type SignKind } from "./host.ts";

/** "Blad1!$G$1" → "G1". */
const local = (address: string) => address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, "");

const supports = (version: string) => {
  try {
    return Office.context.requirements.isSetSupported("ExcelApi", version);
  } catch {
    return false;
  }
};

/** Det dolda bladet med uppdragets uppgifter, och namnen som pekar dit. */
const ENGAGEMENT_CELLS: [keyof typeof ENGAGEMENT_NAMES, string, string][] = [
  ["firm", "B1", "Byrå"],
  ["client", "B2", "Klient"],
  ["orgNr", "B3", "Org.nr"],
  ["year", "B4", "Räkenskapsår"],
];

async function ensureEngagementSheet(context: Excel.RequestContext): Promise<Excel.Worksheet> {
  let sheet = context.workbook.worksheets.getItemOrNullObject(ENGAGEMENT_SHEET);
  await context.sync();
  if (sheet.isNullObject) {
    sheet = context.workbook.worksheets.add(ENGAGEMENT_SHEET);
    sheet.getRange("A1:A4").values = ENGAGEMENT_CELLS.map(([, , label]) => [label]);
    sheet.getRange("B1:B4").numberFormat = [["@"], ["@"], ["@"], ["@"]];
    sheet.getRange("A6").values = [
      ["Bladet används av Kvanto Mallverktyg. Ändra uppgifterna i panelen, inte här."],
    ];
    sheet.visibility = Excel.SheetVisibility.hidden;
  }
  const names = context.workbook.names;
  const existing = ENGAGEMENT_CELLS.map(([key]) =>
    names.getItemOrNullObject(ENGAGEMENT_NAMES[key]),
  );
  await context.sync();
  ENGAGEMENT_CELLS.forEach(([key, address], i) => {
    if (existing[i].isNullObject) names.add(ENGAGEMENT_NAMES[key], sheet.getRange(address));
  });
  return sheet;
}

/** Värden som text i Excel: en text som börjar med = får en apostrof, så att den inte blir en formel. */
const safe = (rows: Value[][]) =>
  rows.map((row) => row.map((v) => (typeof v === "string" && /^[=+@]/.test(v) ? `'${v}` : v)));

/** Utför planens steg på en flik. */
async function applySteps(context: Excel.RequestContext, sheet: Excel.Worksheet, steps: Step[]) {
  const sheetNames = sheet.names;
  sheetNames.load("items/name");
  await context.sync();
  const existing = new Map(sheetNames.items.map((n) => [n.name, n]));
  const later: Step[] = [];

  for (const step of steps) {
    switch (step.op) {
      case "insertRows":
        sheet.getRange(`1:${step.count}`).insert(Excel.InsertShiftDirection.down);
        break;
      case "values":
        sheet.getRange(step.address).values = safe(step.values);
        break;
      case "formulas":
        sheet.getRange(step.address).formulas = step.formulas;
        break;
      case "style": {
        const range = sheet.getRange(step.address);
        const s = step.style;
        if (s.font) range.format.font.name = s.font;
        if (s.size) range.format.font.size = s.size;
        if (s.bold !== undefined) range.format.font.bold = s.bold;
        if (s.color) range.format.font.color = s.color;
        if (s.clearFill) range.format.fill.clear();
        if (s.fill) range.format.fill.color = s.fill;
        if (s.horizontal) range.format.horizontalAlignment = s.horizontal;
        if (s.borderBottom) {
          const border = range.format.borders.getItem(Excel.BorderIndex.edgeBottom);
          border.style = Excel.BorderLineStyle.continuous;
          border.color = s.borderBottom.color;
          border.weight = s.borderBottom.weight;
        }
        if (s.numberFormat) {
          const { rows, columns } = cellsInArea(step.address);
          range.numberFormat = Array.from({ length: rows }, () =>
            Array(columns).fill(s.numberFormat),
          );
        }
        break;
      }
      case "numberFormats":
        sheet.getRange(step.address).numberFormat = step.formats;
        break;
      case "clear":
        sheet.getRange(local(step.address)).clear();
        break;
      case "name":
        existing.get(step.name)?.delete();
        existing.delete(step.name);
        sheetNames.add(step.name, sheet.getRange(step.address));
        break;
      case "gridlines":
        if (supports("1.8")) sheet.showGridlines = step.show;
        break;
      default:
        // Kolumnbredder och utskrift görs efteråt (de behöver läsa av Excel).
        later.push(step);
    }
  }
  await context.sync();

  for (const step of later) {
    if (step.op === "autofit") {
      const range = sheet.getRange(step.address);
      range.format.autofitColumns();
      const { columns } = cellsInArea(step.address);
      const cols = Array.from({ length: columns }, (_, i) => range.getColumn(i).format);
      cols.forEach((c) => c.load("columnWidth"));
      await context.sync();
      cols.forEach((c) => {
        if (c.columnWidth > step.max) c.columnWidth = step.max;
      });
    } else if (step.op === "minWidth") {
      const format = sheet.getRange(`${step.column}:${step.column}`).format;
      format.load("columnWidth");
      await context.sync();
      if (format.columnWidth < step.points) format.columnWidth = step.points;
    } else if (step.op === "print" && supports("1.9")) {
      const layout = sheet.pageLayout;
      layout.orientation = step.landscape
        ? Excel.PageOrientation.landscape
        : Excel.PageOrientation.portrait;
      layout.setPrintArea(step.area);
      const footers = layout.headersFooters.defaultForAllPages;
      // & är en styrkod i sidfoten; && skriver ut ett &.
      footers.leftFooter = step.leftFooter.replace(/&/g, "&&");
      footers.centerFooter = step.centerFooter;
      footers.rightFooter = step.rightFooter;
      await context.sync();
      try {
        // En sida bred, så många sidor höga som behövs.
        layout.zoom = { horizontalFitToPages: 1, verticalFitToPages: 0 };
        await context.sync();
      } catch {
        // Äldre Excel: utskriften fungerar ändå, bara utan anpassning till sidbredden.
      }
    }
  }
  await context.sync();
}

async function snapshotOf(
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
): Promise<Snapshot> {
  const used = sheet.getUsedRangeOrNullObject(true);
  used.load("values,numberFormat,rowIndex,columnIndex,rowCount,columnCount");
  sheet.load("name");
  const named = Object.entries(NAMES).map(([key, name]) => {
    const item = sheet.names.getItemOrNullObject(name);
    return { key, item };
  });
  await context.sync();
  const ranges = named.map(({ key, item }) => {
    const range = item.isNullObject ? null : item.getRangeOrNullObject();
    range?.load("address");
    return { key, range };
  });
  await context.sync();
  const names: Snapshot["names"] = {};
  for (const { key, range } of ranges) {
    if (range && !range.isNullObject) names[key as keyof typeof NAMES] = local(range.address);
  }
  return {
    used: used.isNullObject
      ? null
      : {
          top: used.rowIndex,
          left: used.columnIndex,
          rows: used.rowCount,
          columns: used.columnCount,
        },
    values: used.isNullObject ? [] : (used.values as Value[][]),
    formats: used.isNullObject ? [] : (used.numberFormat as string[][]),
    names,
    sheetName: sheet.name,
  };
}

/** Rutorna (namn per flik) för en signering. */
const SIGN_NAMES: Record<SignKind, [string, string]> = {
  prepared: [NAMES.preparedBy, NAMES.preparedDate],
  reviewed: [NAMES.reviewedBy, NAMES.reviewedDate],
};

export function createExcelHost(): Host {
  const listeners = new Set<() => void>();
  let watching = false;

  async function watch() {
    if (watching) return;
    watching = true;
    await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.onActivated.add(async () => listeners.forEach((l) => l()));
      if (supports("1.9")) sheets.onChanged.add(async () => listeners.forEach((l) => l()));
      await context.sync();
    }).catch(() => (watching = false));
  }

  return {
    kind: "excel",

    getEngagement: () =>
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItemOrNullObject(ENGAGEMENT_SHEET);
        await context.sync();
        if (sheet.isNullObject) return { client: "", orgNr: "", year: "" };
        const range = sheet.getRange("B2:B4");
        range.load("values");
        await context.sync();
        const [[client], [orgNr], [year]] = range.values.map((r) => r.map((v) => String(v ?? "")));
        return { client, orgNr, year };
      }),

    setEngagement: (engagement: Engagement, profile: Profile) =>
      Excel.run(async (context) => {
        const sheet = await ensureEngagementSheet(context);
        sheet.getRange("B1:B4").values = [
          [profile.firm],
          [engagement.client],
          [engagement.orgNr],
          [engagement.year],
        ];
        await context.sync();
      }),

    sheets: () =>
      Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        const list = sheets.items.filter((s) => s.name !== ENGAGEMENT_SHEET);
        const lookups = list.map((sheet) => ({
          sheet,
          items: [
            NAMES.header,
            NAMES.reference,
            ...SIGN_NAMES.prepared,
            ...SIGN_NAMES.reviewed,
          ].map((name) => sheet.names.getItemOrNullObject(name)),
        }));
        await context.sync();
        const ranges = lookups.map(({ items }) =>
          items.map((item) => {
            if (item.isNullObject) return null;
            const range = item.getRangeOrNullObject();
            range.load("values");
            return range;
          }),
        );
        await context.sync();
        return lookups.map(({ sheet }, i): SheetStatus => {
          const value = (j: number) => {
            const range = ranges[i][j];
            return range && !range.isNullObject ? range.values[0][0] : "";
          };
          return {
            name: sheet.name,
            adapted: !!ranges[i][0],
            reference: String(value(1) ?? ""),
            prepared: { by: String(value(2) ?? ""), date: dateText(value(3)) },
            reviewed: { by: String(value(4) ?? ""), date: dateText(value(5)) },
          };
        });
      }),

    activeSheet: () =>
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await context.sync();
        return sheet.name;
      }),

    adapt: (name: string, profile: Profile, options: Options) =>
      Excel.run(async (context) => {
        const engagement = await ensureEngagementSheet(context);
        engagement.getRange("B1").values = [[profile.firm]];
        const sheet = context.workbook.worksheets.getItem(name);
        const snapshot = await snapshotOf(context, sheet);
        await applySteps(context, sheet, planSheet(snapshot, profile, options));
      }),

    sign: (name: string, kind: SignKind, initials: string) =>
      Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(name);
        const [byName, dateName] = SIGN_NAMES[kind];
        const by = sheet.names.getItemOrNullObject(byName);
        const date = sheet.names.getItemOrNullObject(dateName);
        await context.sync();
        if (by.isNullObject || date.isNullObject) return false;
        by.getRange().values = safe([[initials]]);
        date.getRange().values = [[excelToday()]];
        await context.sync();
        return true;
      }),

    clearClientData: () =>
      Excel.run(async (context) => {
        const engagement = context.workbook.worksheets.getItemOrNullObject(ENGAGEMENT_SHEET);
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (!engagement.isNullObject) engagement.getRange("B2:B4").values = [[""], [""], [""]];
        const items = sheets.items.flatMap((sheet) =>
          CLIENT_FIELDS.map((name) => sheet.names.getItemOrNullObject(name)),
        );
        await context.sync();
        for (const item of items) if (!item.isNullObject) item.getRange().values = [[""]];
        await context.sync();
      }),

    file: () =>
      new Promise((resolve, reject) => {
        Office.context.document.getFileAsync(
          Office.FileType.Compressed,
          { sliceSize: 4 * 1024 * 1024 },
          (result) => {
            if (result.status !== Office.AsyncResultStatus.Succeeded) {
              reject(new Error("Excel lämnade inte ut arbetsboken. Spara den och försök igen."));
              return;
            }
            const file = result.value;
            const parts: Uint8Array[] = [];
            const next = (index: number) => {
              if (index >= file.sliceCount) {
                file.closeAsync();
                const out = new Uint8Array(parts.reduce((s, p) => s + p.length, 0));
                let at = 0;
                for (const p of parts) {
                  out.set(p, at);
                  at += p.length;
                }
                resolve(out);
                return;
              }
              file.getSliceAsync(index, (slice) => {
                if (slice.status !== Office.AsyncResultStatus.Succeeded) {
                  file.closeAsync();
                  reject(new Error("Arbetsboken gick inte att läsa. Försök igen."));
                  return;
                }
                parts.push(new Uint8Array(slice.value.data as number[]));
                next(index + 1);
              });
            };
            next(0);
          },
        );
      }),

    onChange(listener) {
      listeners.add(listener);
      void watch();
      return () => void listeners.delete(listener);
    },
  };
}
