// Det panelen behöver av Excel. Två varianter: riktiga Excel (excelHost.ts)
// och låtsas-Excel i webbläsaren (browserHost.ts).

import type { Options } from "../lib/layout.ts";
import type { Profile } from "../lib/profile.ts";
import type { Engagement } from "./memory.ts";

export type { Engagement };

export interface SheetStatus {
  name: string;
  /** Fliken har fått byråns huvud. */
  adapted: boolean;
  reference: string;
  prepared: { by: string; date: string };
  reviewed: { by: string; date: string };
}

export type SignKind = "prepared" | "reviewed";

export interface Host {
  kind: "excel" | "browser";
  getEngagement(): Promise<Engagement>;
  setEngagement(engagement: Engagement, profile: Profile): Promise<void>;
  sheets(): Promise<SheetStatus[]>;
  activeSheet(): Promise<string>;
  /** Anpassar en flik till byråns revisionspapper (se lib/layout.ts). */
  adapt(sheet: string, profile: Profile, options: Options): Promise<void>;
  /** Initialer och dagens datum i fliken. Falskt om fliken inte är anpassad. */
  sign(sheet: string, kind: SignKind, initials: string): Promise<boolean>;
  /** Tömmer klientens uppgifter: uppdraget, signeringar och slutsatser. */
  clearClientData(): Promise<void>;
  /** Arbetsboken som .xlsx (null där det inte går, t.ex. i webbläsarläget). */
  file(): Promise<Uint8Array | null>;
  /** Anropas när användaren byter flik eller ändrar i arbetsboken. */
  onChange(listener: () => void): () => void;
}

/** Excels datumtal → "2026-11-25" (tomt för annat). */
export function dateText(value: unknown): string {
  if (typeof value === "number" && value > 0) {
    return new Date(Date.UTC(1899, 11, 30) + Math.round(value) * 86400000)
      .toISOString()
      .slice(0, 10);
  }
  return typeof value === "string" ? value : "";
}
