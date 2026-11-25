// Hur en flik görs om till ett revisionspapper. Här räknas allt ut som en lista
// med steg (en plan); src/excel/ utför stegen i Excel, eller i låtsas-Excel när
// panelen provas i webbläsaren. Planen går att testa utan Excel.
//
// Fliken får:
//   överst  byråns huvud: byrå, klient, org.nr, räkenskapsår, referens,
//           upprättat och granskat (initialer och datum), rubrik, syfte och
//           utfört arbete
//   sedan   fliken som den var, med byråns typsnitt och talformat
//   sist    slutsats och en förklaring av granskningsmarkeringarna

import { tint, type Profile } from "./profile.ts";

/** Rader som huvudet tar upp (sista raden är tom). */
export const HEADER_ROWS = 9;

/** Namnen (per flik) som pekar ut rutorna, så att de hittas igen. */
export const NAMES = {
  header: "Mall_Huvud",
  reference: "Mall_Referens",
  preparedBy: "Mall_UpprattatAv",
  preparedDate: "Mall_UpprattatDatum",
  reviewedBy: "Mall_GranskatAv",
  reviewedDate: "Mall_GranskatDatum",
  title: "Mall_Rubrik",
  purpose: "Mall_Syfte",
  work: "Mall_Utfort",
  conclusion: "Mall_Slutsats",
  markings: "Mall_Markeringar",
} as const;

/** Namnen (för hela arbetsboken) på uppdragets uppgifter, på det dolda bladet. */
export const ENGAGEMENT_SHEET = "Mallverktyg";
export const ENGAGEMENT_NAMES = {
  firm: "Mall_Byra",
  client: "Mall_Klient",
  orgNr: "Mall_Orgnr",
  year: "Mall_Rakenskapsar",
} as const;

export type Value = string | number | boolean;

export interface Snapshot {
  /** Flikens använda område (0-baserat), eller null om fliken är tom. */
  used: { top: number; left: number; rows: number; columns: number } | null;
  /** Värden och talformat i det använda området. */
  values: Value[][];
  formats: string[][];
  /** Flikens namn (NAMES) och deras adresser, om fliken redan är anpassad. */
  names: Partial<Record<keyof typeof NAMES, string>>;
  sheetName: string;
}

export interface Style {
  font?: string;
  size?: number;
  bold?: boolean;
  color?: string;
  fill?: string;
  /** Ta bort bakgrundsfärgen (före andra färger i samma steg). */
  clearFill?: boolean;
  borderBottom?: { color: string; weight: "Thin" | "Medium" };
  horizontal?: "Left" | "Center" | "Right";
  numberFormat?: string;
}

export type Step =
  | { op: "insertRows"; count: number }
  | { op: "values"; address: string; values: Value[][] }
  | { op: "formulas"; address: string; formulas: string[][] }
  | { op: "style"; address: string; style: Style }
  | { op: "numberFormats"; address: string; formats: string[][] }
  | { op: "clear"; address: string }
  | { op: "name"; name: string; address: string }
  | { op: "autofit"; address: string; max: number }
  | { op: "minWidth"; column: string; points: number }
  | { op: "gridlines"; show: boolean }
  | {
      op: "print";
      area: string;
      landscape: boolean;
      leftFooter: string;
      centerFooter: string;
      rightFooter: string;
    };

export interface Options {
  /** Byråns typsnitt, talformat, tabellrubrik och kolumnbredder på flikens innehåll. */
  style: boolean;
  /** Utskriftsinställningar. */
  print: boolean;
}

// ---------------------------------------------------------------------------
// Adresser

export function columnName(index: number): string {
  let name = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    name = String.fromCharCode(65 + ((n - 1) % 26)) + name;
  }
  return name;
}

export const cell = (row: number, col: number) => `${columnName(col)}${row + 1}`;
export const area = (r1: number, c1: number, r2: number, c2: number) =>
  r1 === r2 && c1 === c2 ? cell(r1, c1) : `${cell(r1, c1)}:${cell(r2, c2)}`;

/** "Blad1!$C$4" eller "C4:D5" → rad och kolumn för första cellen (0-baserat). */
export function parseCell(address: string): { row: number; col: number } | null {
  const m = /([A-Z]+)\$?(\d+)/.exec(address.slice(address.lastIndexOf("!") + 1).replace(/\$/g, ""));
  if (!m) return null;
  let col = 0;
  for (const ch of m[1]) col = col * 26 + ch.charCodeAt(0) - 64;
  return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Antal rader och kolumner i ett område ("A3:F9" → 7 rader, 6 kolumner). */
export function cellsInArea(address: string): { rows: number; columns: number } {
  const [a, b = a] = address.slice(address.lastIndexOf("!") + 1).split(":");
  const from = parseCell(a)!;
  const to = parseCell(b)!;
  return { rows: to.row - from.row + 1, columns: to.col - from.col + 1 };
}

/** Sista raden i ett område ("A3:F9" → 8). */
export function lastRow(address: string): number | null {
  const parts = address.slice(address.lastIndexOf("!") + 1).split(":");
  return parseCell(parts[parts.length - 1])?.row ?? null;
}

// ---------------------------------------------------------------------------

const isEmpty = (v: Value | undefined) => v === "" || v === null || v === undefined;

/** Formel som visar ett namn, eller inget om det är tomt (inte 0). */
const show = (name: string) => `=IF(${name}="","",${name})`;

/**
 * Tabellens rubrikrad: den första raden med innehåll, om den mest består av
 * text och någon av de följande raderna innehåller tal.
 */
export function findTableHeader(values: Value[][]): number | null {
  const first = values.findIndex((row) => row.some((v) => !isEmpty(v)));
  if (first === -1) return null;
  const filled = values[first].filter((v) => !isEmpty(v));
  const texts = filled.filter((v) => typeof v === "string").length;
  if (filled.length < 2 || texts < Math.ceil(filled.length / 2)) return null;
  const numbersBelow = values
    .slice(first + 1, first + 6)
    .some((row) => row.some((v) => typeof v === "number"));
  return numbersBelow ? first : null;
}

/** Kolumnrubriker för nummer som inte är belopp (fakturanummer, konton, år): de får inget talformat. */
const IDENTIFIER_WORDS = new Set([
  "konto",
  "år",
  "id",
  "ref",
  "referens",
  "kod",
  "ver",
  "verifikation",
  "org",
  "period",
  "månad",
]);
export function isIdentifier(header: string): boolean {
  return header
    .toLowerCase()
    .split(/[^a-zåäö0-9]+/)
    .some((w) => /(nr|nummer)$/.test(w) || IDENTIFIER_WORDS.has(w));
}

/**
 * Talformat för tal som står utan format ("General"). Kolumner vars rubrik
 * (raden `header`) ser ut som ett nummer lämnas som de är.
 */
export function numberFormats(
  values: Value[][],
  formats: string[][],
  mode: Profile["numbers"],
  header: number | null = null,
) {
  const identifiers = new Set(
    header === null
      ? []
      : values[header].flatMap((v, c) => (typeof v === "string" && isIdentifier(v) ? [c] : [])),
  );
  const columns = Math.max(0, ...values.map((r) => r.length));
  const decimals = Array.from({ length: columns }, (_, c) =>
    values.some((row) => typeof row[c] === "number" && !Number.isInteger(row[c])),
  );
  return values.map((row, r) =>
    row.map((v, c) => {
      const current = formats[r]?.[c] ?? "General";
      if (typeof v !== "number" || current !== "General" || identifiers.has(c)) return current;
      const withDecimals = mode === "decimaler" || (mode === "auto" && decimals[c]);
      return withDecimals ? "#,##0.00" : "#,##0";
    }),
  );
}

/**
 * Planen för en flik. Är fliken redan anpassad (den har namnet Mall_Huvud)
 * uppdateras huvud, utseende och markeringar, utan att något läggs till två
 * gånger och utan att det som fyllts i skrivs över.
 */
export function planSheet(snapshot: Snapshot, profile: Profile, options: Options): Step[] {
  const steps: Step[] = [];
  const H = HEADER_ROWS;
  const updating = !!snapshot.names.header;
  const used = snapshot.used;
  const accent = profile.accent;
  const input: Style = { fill: profile.inputColor };

  // Innehållet (0-baserat, efter att huvudet lagts in).
  let contentTop: number | null = null;
  let contentBottom: number | null = null;
  let left = 0;
  let right = 0;
  let values: Value[][] = [];
  let formats: string[][] = [];
  const conclusionAt = snapshot.names.conclusion ? parseCell(snapshot.names.conclusion)?.row : null;

  if (used) {
    left = used.left;
    right = used.left + used.columns - 1;
    // Raderna i ögonblicksbilden som är flikens eget innehåll.
    const shift = updating ? 0 : H;
    snapshot.values.forEach((row, i) => {
      const absolute = used.top + i;
      if (updating && (absolute < H || (conclusionAt != null && absolute >= conclusionAt - 1)))
        return;
      values.push(row);
      formats.push(snapshot.formats[i] ?? []);
      if (contentTop === null) contentTop = absolute + shift;
      contentBottom = absolute + shift;
    });
    // Tomma rader sist räknas inte.
    while (values.length && values[values.length - 1].every(isEmpty)) {
      values.pop();
      formats.pop();
      contentBottom = contentBottom! - 1;
    }
    if (!values.length) contentTop = contentBottom = null;
    if (updating && contentTop !== null) {
      // Kolumnerna som innehållet faktiskt använder.
      const usedCols = values.flatMap((row) => row.map((v, c) => (isEmpty(v) ? -1 : c)));
      right = used.left + Math.max(0, ...usedCols);
    }
  }

  // Högra delen av huvudet (referens, upprättat, granskat): slutar vid
  // tabellens högra kant, men börjar aldrig före kolumn F. Vid uppdatering där
  // den redan står.
  const existingRef = snapshot.names.reference ? parseCell(snapshot.names.reference) : null;
  const R = existingRef ? existingRef.col - 1 : Math.max(5, right - 2);
  const last = Math.max(right, R + 2);

  if (!updating) steps.push({ op: "insertRows", count: H });

  // Huvudet: ledtexter och formler skrivs alltid (de kan ha ändrats i
  // inställningarna); det som fylls i skrivs bara första gången.
  steps.push(
    { op: "formulas", address: cell(0, 0), formulas: [[show("Mall_Byra")]] },
    {
      op: "values",
      address: area(1, 0, 3, 0),
      values: [["Klient"], ["Org.nr"], ["Räkenskapsår"]],
    },
    {
      op: "formulas",
      address: area(1, 1, 3, 1),
      formulas: [[show("Mall_Klient")], [show("Mall_Orgnr")], [show("Mall_Rakenskapsar")]],
    },
    { op: "values", address: cell(0, R), values: [["Referens"]] },
    {
      op: "values",
      address: area(1, R, 2, R),
      values: [["Upprättat"], ["Granskat"]],
    },
    {
      op: "values",
      address: area(5, 0, 6, 0),
      values: [[profile.sections.purpose], [profile.sections.work]],
    },
  );
  if (!updating) {
    steps.push({ op: "values", address: cell(4, 0), values: [[snapshot.sheetName]] });
  }
  // Mallverktygets delar får byråns bakgrund (eller ingen), före rutorna som
  // fylls i, så att de behåller sin egen färg.
  const background: Style = profile.background ? { fill: profile.background } : { clearFill: true };
  steps.push(
    {
      op: "style",
      address: area(0, 0, H - 1, last),
      style: { font: profile.font, size: profile.fontSize },
    },
    { op: "style", address: area(0, 0, H - 2, last), style: background },
    {
      op: "style",
      address: cell(0, 0),
      style: { bold: true, color: accent, size: profile.fontSize + 2 },
    },
    { op: "style", address: area(1, 0, 3, 0), style: { color: "#595959" } },
    { op: "style", address: area(0, R, 2, R), style: { color: "#595959" } },
    {
      op: "style",
      address: cell(0, R + 1),
      style: { ...input, bold: true, size: profile.fontSize + 2, horizontal: "Center" },
    },
    { op: "style", address: area(1, R + 1, 2, R + 1), style: { ...input, horizontal: "Center" } },
    {
      op: "style",
      address: area(1, R + 2, 2, R + 2),
      style: { ...input, horizontal: "Center", numberFormat: "yyyy-mm-dd" },
    },
    {
      op: "style",
      address: cell(4, 0),
      style: { bold: true, size: profile.fontSize + 4, color: accent },
    },
    { op: "style", address: area(5, 0, 6, 0), style: { bold: true } },
    { op: "style", address: area(5, 1, 6, Math.max(1, R - 1)), style: input },
    {
      op: "style",
      address: area(7, 0, 7, last),
      style: { borderBottom: { color: accent, weight: "Medium" } },
    },
    { op: "name", name: NAMES.header, address: area(0, 0, H - 1, last) },
    { op: "name", name: NAMES.reference, address: cell(0, R + 1) },
    { op: "name", name: NAMES.preparedBy, address: cell(1, R + 1) },
    { op: "name", name: NAMES.preparedDate, address: cell(1, R + 2) },
    { op: "name", name: NAMES.reviewedBy, address: cell(2, R + 1) },
    { op: "name", name: NAMES.reviewedDate, address: cell(2, R + 2) },
    { op: "name", name: NAMES.title, address: cell(4, 0) },
    { op: "name", name: NAMES.purpose, address: cell(5, 1) },
    { op: "name", name: NAMES.work, address: cell(6, 1) },
  );

  // Flikens innehåll i byråns utseende.
  if (options.style && contentTop !== null && contentBottom !== null) {
    const content = area(contentTop, left, contentBottom, right);
    const width = Math.max(1, ...values.map((r) => r.length));
    const header = findTableHeader(values);
    steps.push(
      { op: "style", address: content, style: { font: profile.font, size: profile.fontSize } },
      {
        op: "numberFormats",
        address: area(contentTop, left, contentBottom, left + width - 1),
        formats: numberFormats(
          values.map((r) => pad(r, width, "")),
          formats.map((r) => pad(r, width, "General")),
          profile.numbers,
          header,
        ),
      },
    );
    if (header !== null) {
      steps.push({
        op: "style",
        address: area(contentTop + header, left, contentTop + header, right),
        style: {
          bold: true,
          fill: tint(accent, 0.85),
          borderBottom: { color: accent, weight: "Thin" },
        },
      });
    }
    steps.push({ op: "autofit", address: content, max: 300 });
  }
  steps.push({ op: "minWidth", column: "A", points: 84 });

  // Slutsats och markeringar, efter innehållet.
  const conclusionRow =
    conclusionAt ?? (contentBottom !== null ? Math.max(contentBottom + 2, H) : H);
  const legendTop = conclusionRow + 2;
  const legend = [
    ["Markeringar", ""],
    ...profile.markings.filter((m) => m.symbol.trim()).map((m) => [m.symbol, m.meaning]),
  ];
  const legendBottom = legendTop + legend.length - 1;
  // De gamla markeringarna bort först (de kan ha varit fler).
  if (snapshot.names.markings) steps.push({ op: "clear", address: snapshot.names.markings });
  steps.push(
    { op: "style", address: area(conclusionRow, 0, legendBottom, last), style: background },
    { op: "values", address: cell(conclusionRow, 0), values: [[profile.sections.conclusion]] },
    {
      op: "style",
      address: cell(conclusionRow, 0),
      style: { bold: true, font: profile.font, size: profile.fontSize },
    },
    {
      op: "style",
      address: area(conclusionRow, 1, conclusionRow, Math.max(1, R - 1)),
      style: input,
    },
    { op: "name", name: NAMES.conclusion, address: cell(conclusionRow, 1) },
  );
  steps.push(
    { op: "values", address: area(legendTop, 0, legendBottom, 1), values: legend },
    {
      op: "style",
      address: area(legendTop, 0, legendBottom, 1),
      style: { font: profile.font, size: profile.fontSize - 1 },
    },
    { op: "style", address: cell(legendTop, 0), style: { bold: true } },
    { op: "name", name: NAMES.markings, address: area(legendTop, 0, legendBottom, 1) },
  );
  if (legend.length > 1) {
    steps.push({
      op: "style",
      address: area(legendTop + 1, 0, legendBottom, 0),
      style: { bold: true, color: accent, horizontal: "Center" },
    });
  }

  if (profile.hideGridlines) steps.push({ op: "gridlines", show: false });

  if (options.print) {
    const columns = last + 1;
    steps.push({
      op: "print",
      area: area(0, 0, legendBottom, last),
      landscape:
        profile.orientation === "liggande" || (profile.orientation === "auto" && columns > 8),
      leftFooter: profile.firm,
      centerFooter: "&A",
      rightFooter: "Sida &P av &N",
    });
  }
  return steps;
}

function pad<T>(row: T[], width: number, fill: T): T[] {
  return [...row, ...Array<T>(Math.max(0, width - row.length)).fill(fill)];
}

// ---------------------------------------------------------------------------
// Signering och mall

/** Dagens datum som Excels datumtal. */
export function excelToday(now = new Date()): number {
  const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
  return (utc - Date.UTC(1899, 11, 30)) / 86400000;
}

/** Rutorna som töms när arbetsboken görs till en mall (klientens uppgifter). */
export const CLIENT_FIELDS = [
  NAMES.preparedBy,
  NAMES.preparedDate,
  NAMES.reviewedBy,
  NAMES.reviewedDate,
  NAMES.conclusion,
] as const;
