import { useEffect, useReducer } from "react";
import type { Host } from "./excel/host.ts";
import { dateText } from "./excel/host.ts";
import type { MemoryCell, MemoryWorkbook } from "./excel/memory.ts";
import { columnName, type Value } from "./lib/layout.ts";

/** Hur ett värde visas med sitt talformat (förenklat, som i svenska Excel). */
function shown(value: Value, format: string): string {
  if (typeof value !== "number") return String(value);
  if (format === "yyyy-mm-dd") return dateText(value);
  const decimals = format === "#,##0.00" ? 2 : format === "#,##0" ? 0 : undefined;
  return value.toLocaleString("sv-SE", {
    minimumFractionDigits: decimals,
    maximumFractionDigits: decimals ?? 10,
    useGrouping: decimals !== undefined,
  });
}

function style(c: MemoryCell | undefined): React.CSSProperties {
  const s = c?.style ?? {};
  return {
    fontFamily: s.font,
    fontSize: s.size ? `${s.size + 2}px` : undefined,
    fontWeight: s.bold ? 600 : undefined,
    color: s.color,
    background: s.fill,
    borderBottom: s.borderBottom
      ? `${s.borderBottom.weight === "Medium" ? 2 : 1}px solid ${s.borderBottom.color}`
      : undefined,
    textAlign:
      s.horizontal === "Center"
        ? "center"
        : s.horizontal === "Right" || typeof c?.value === "number"
          ? "right"
          : undefined,
  };
}

/** Låtsas-arbetsboken bredvid panelen i webbläsarläget. */
export function Preview({
  host,
}: {
  host: Host & { book: MemoryWorkbook; activate(i: number): void };
}) {
  const [, redraw] = useReducer((n: number) => n + 1, 0);
  useEffect(() => host.onChange(redraw), [host]);
  const book = host.book;
  const sheet = book.sheets[book.active];
  const used = sheet.used();
  const rows = Math.max(30, (used ? used.top + used.rows : 0) + 3);
  const cols = Math.max(10, (used ? used.left + used.columns : 0) + 1);

  return (
    <div className="preview">
      <p className="preview-note">
        Provversion i webbläsaren: arbetsboken nedan finns bara här. I Excel anpassas dina riktiga
        flikar.
      </p>
      <div className={`grid-wrap${sheet.gridlines ? "" : " no-grid"}`}>
        <table className="grid">
          <thead>
            <tr>
              <th />
              {Array.from({ length: cols }, (_, c) => (
                <th key={c} style={{ minWidth: sheet.widths.get(c) ?? 64 }}>
                  {columnName(c)}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {Array.from({ length: rows }, (_, r) => (
              <tr key={r}>
                <th>{r + 1}</th>
                {Array.from({ length: cols }, (_, c) => {
                  const cell = sheet.get(r, c);
                  return (
                    <td key={c} style={style(cell)}>
                      {shown(sheet.display(r, c), cell?.format ?? "General")}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <div className="sheet-tabs" role="tablist">
        {book.sheets.map((s, i) => (
          <button
            key={s.name}
            role="tab"
            aria-selected={i === book.active}
            className={i === book.active ? "active" : ""}
            onClick={() => host.activate(i)}
          >
            {s.name}
          </button>
        ))}
      </div>
    </div>
  );
}
