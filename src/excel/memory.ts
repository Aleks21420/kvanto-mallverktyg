// Låtsas-Excel i minnet: utför planens steg på en enkel arbetsbok. Används av
// testerna och när panelen provas i en vanlig webbläsare (npm run dev:browser).

import {
  ENGAGEMENT_NAMES,
  NAMES,
  parseCell,
  type Snapshot,
  type Step,
  type Style,
  type Value,
} from "../lib/layout.ts";

export interface MemoryCell {
  value: Value;
  formula?: string;
  format: string;
  style: Style;
}

export interface Engagement {
  client: string;
  orgNr: string;
  year: string;
}

/** Alla celler i ett område ("B2:D5" eller "B2"), 0-baserat. */
export function cellsIn(address: string): { row: number; col: number }[] {
  const [a, b = a] = address.slice(address.lastIndexOf("!") + 1).split(":");
  const from = parseCell(a)!;
  const to = parseCell(b)!;
  const out = [];
  for (let row = from.row; row <= to.row; row++)
    for (let col = from.col; col <= to.col; col++) out.push({ row, col });
  return out;
}

export class MemorySheet {
  cells = new Map<string, MemoryCell>();
  names = new Map<string, string>();
  gridlines = true;
  widths = new Map<number, number>();
  print: Extract<Step, { op: "print" }> | null = null;
  name: string;
  private book: MemoryWorkbook;

  constructor(book: MemoryWorkbook, name: string) {
    this.book = book;
    this.name = name;
  }

  get(row: number, col: number): MemoryCell | undefined {
    return this.cells.get(`${row},${col}`);
  }

  private ensure(row: number, col: number): MemoryCell {
    const key = `${row},${col}`;
    let c = this.cells.get(key);
    if (!c) {
      c = { value: "", format: "General", style: {} };
      this.cells.set(key, c);
    }
    return c;
  }

  /** Skriver värden som en användare, från cellen `address` och åt höger/nedåt. */
  set(address: string, rows: Value[][]) {
    const start = parseCell(address)!;
    rows.forEach((row, r) =>
      row.forEach((value, c) => {
        const target = this.ensure(start.row + r, start.col + c);
        target.value = value;
        delete target.formula;
      }),
    );
  }

  /** Värdet som Excel skulle visa (formlerna som visar uppdragets namn räknas ut). */
  display(row: number, col: number): Value {
    const c = this.get(row, col);
    if (!c) return "";
    if (c.formula) {
      const m = /^=IF\((\w+)=""/.exec(c.formula);
      return m ? this.book.nameValue(m[1]) : c.formula;
    }
    return c.value;
  }

  /** Som Excels getUsedRange(true): rutan runt alla celler med värde eller formel. */
  used() {
    let top = Infinity;
    let left = Infinity;
    let bottom = -1;
    let right = -1;
    for (const [key, c] of this.cells) {
      if (c.value === "" && !c.formula) continue;
      const [r, col] = key.split(",").map(Number);
      top = Math.min(top, r);
      left = Math.min(left, col);
      bottom = Math.max(bottom, r);
      right = Math.max(right, col);
    }
    return bottom < 0 ? null : { top, left, rows: bottom - top + 1, columns: right - left + 1 };
  }

  snapshot(): Snapshot {
    const used = this.used();
    const values: Value[][] = [];
    const formats: string[][] = [];
    if (used) {
      for (let r = 0; r < used.rows; r++) {
        values.push([]);
        formats.push([]);
        for (let c = 0; c < used.columns; c++) {
          values[r].push(this.display(used.top + r, used.left + c));
          formats[r].push(this.get(used.top + r, used.left + c)?.format ?? "General");
        }
      }
    }
    const names: Snapshot["names"] = {};
    for (const [key, name] of Object.entries(NAMES)) {
      const address = this.names.get(name);
      if (address) names[key as keyof typeof NAMES] = address;
    }
    return { used, values, formats, names, sheetName: this.name };
  }

  apply(steps: Step[]) {
    for (const step of steps) {
      switch (step.op) {
        case "insertRows": {
          const moved = new Map<string, MemoryCell>();
          for (const [key, c] of this.cells) {
            const [r, col] = key.split(",").map(Number);
            moved.set(`${r + step.count},${col}`, c);
          }
          this.cells = moved;
          for (const [name, address] of this.names) {
            this.names.set(
              name,
              address.replace(
                /([A-Z]+)(\d+)/g,
                (_, col, row) => `${col}${Number(row) + step.count}`,
              ),
            );
          }
          break;
        }
        case "values":
        case "formulas": {
          const start = parseCell(step.address)!;
          const rows = step.op === "values" ? step.values : step.formulas;
          rows.forEach((row, r) =>
            row.forEach((v, c) => {
              const target = this.ensure(start.row + r, start.col + c);
              if (step.op === "values") {
                target.value = v;
                delete target.formula;
              } else {
                target.formula = String(v);
                target.value = "";
              }
            }),
          );
          break;
        }
        case "style":
          for (const { row, col } of cellsIn(step.address)) {
            const target = this.ensure(row, col);
            const { numberFormat, clearFill, ...style } = step.style;
            if (clearFill) delete target.style.fill;
            target.style = { ...target.style, ...style };
            if (numberFormat) target.format = numberFormat;
          }
          break;
        case "numberFormats": {
          const start = parseCell(step.address)!;
          step.formats.forEach((row, r) =>
            row.forEach((f, c) => {
              const target = this.get(start.row + r, start.col + c);
              if (target) target.format = f;
              else if (f !== "General") this.ensure(start.row + r, start.col + c).format = f;
            }),
          );
          break;
        }
        case "clear":
          for (const { row, col } of cellsIn(step.address)) this.cells.delete(`${row},${col}`);
          break;
        case "name":
          this.names.set(step.name, step.address);
          break;
        case "autofit":
          for (const { row, col } of cellsIn(step.address)) {
            const text = String(this.display(row, col));
            const width = Math.min(
              step.max,
              Math.max(this.widths.get(col) ?? 48, text.length * 7 + 10),
            );
            this.widths.set(col, width);
          }
          break;
        case "minWidth": {
          const col = parseCell(`${step.column}1`)!.col;
          this.widths.set(col, Math.max(this.widths.get(col) ?? 48, step.points));
          break;
        }
        case "gridlines":
          this.gridlines = step.show;
          break;
        case "print":
          this.print = step;
          break;
      }
    }
  }

  /** Text för felsökning: rad för rad, celler med | emellan. */
  text(): string {
    const used = this.used();
    if (!used) return "";
    const lines = [];
    for (let r = 0; r < used.top + used.rows; r++) {
      const row = [];
      for (let c = 0; c < used.left + used.columns; c++) row.push(String(this.display(r, c)));
      lines.push(row.join(" | ").replace(/[ |]+$/, ""));
    }
    return lines.join("\n");
  }
}

export class MemoryWorkbook {
  sheets: MemorySheet[] = [];
  engagement: Engagement = { client: "", orgNr: "", year: "" };
  firm = "";
  active = 0;

  addSheet(name: string): MemorySheet {
    const sheet = new MemorySheet(this, name);
    this.sheets.push(sheet);
    return sheet;
  }

  sheet(name: string): MemorySheet {
    const found = this.sheets.find((s) => s.name === name);
    if (!found) throw new Error(`Fliken ${name} finns inte`);
    return found;
  }

  nameValue(name: string): Value {
    if (name === ENGAGEMENT_NAMES.firm) return this.firm;
    if (name === ENGAGEMENT_NAMES.client) return this.engagement.client;
    if (name === ENGAGEMENT_NAMES.orgNr) return this.engagement.orgNr;
    if (name === ENGAGEMENT_NAMES.year) return this.engagement.year;
    return "";
  }
}
