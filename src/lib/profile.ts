// Byråns inställningar: hur revisionspapperen ska se ut. Sparas i datorn och
// kan delas med kollegorna som en fil, så att hela byrån får samma utseende.

export interface Marking {
  symbol: string;
  meaning: string;
}

export interface Profile {
  /** Byråns namn, överst på varje flik. */
  firm: string;
  font: string;
  fontSize: number;
  /** Byråns färg (rubriker, linjer), som #rrggbb. */
  accent: string;
  /** Bakgrund för rutor som ska fyllas i, som #rrggbb. */
  inputColor: string;
  /** Talformat för tal som saknar format: heltal, två decimaler eller efter talet. */
  numbers: "heltal" | "decimaler" | "auto";
  /** Rubrikerna för avsnitten. */
  sections: { purpose: string; work: string; conclusion: string };
  /** Granskningsmarkeringar och vad de betyder. */
  markings: Marking[];
  /** Utskrift: liggande, stående, eller liggande när tabellen är bred. */
  orientation: "auto" | "liggande" | "stående";
  /** Dölj stödlinjerna, så att flikarna ser ut som papper. */
  hideGridlines: boolean;
}

/** Förslaget till standard. Byrån ändrar det under Byråns inställningar. */
export const DEFAULT_PROFILE: Profile = {
  firm: "",
  font: "Calibri",
  fontSize: 11,
  accent: "#1F4E79",
  inputColor: "#FFF5CC",
  numbers: "auto",
  sections: { purpose: "Syfte", work: "Utfört arbete", conclusion: "Slutsats" },
  markings: [
    { symbol: "✓", meaning: "Stämt mot underlag" },
    { symbol: "✗", meaning: "Avvikelse, se kommentar" },
    { symbol: "Σ", meaning: "Summerat och kontrollerat" },
    { symbol: "FÅ", meaning: "Stämt mot föregående år" },
    { symbol: "HB", meaning: "Stämt mot huvudbok" },
  ],
  orientation: "auto",
  hideGridlines: true,
};

export const FONTS = ["Calibri", "Aptos", "Arial", "Segoe UI", "Verdana", "Times New Roman"];

const isColor = (s: unknown): s is string => typeof s === "string" && /^#[0-9a-f]{6}$/i.test(s);
const text = (s: unknown, fallback: string, max = 120) =>
  typeof s === "string" ? s.slice(0, max) : fallback;

/**
 * En giltig profil ur vad som helst (sparad data, en importerad fil): okända
 * eller felaktiga fält ersätts med standardvärdet.
 */
export function cleanProfile(input: unknown): Profile {
  const p = (input && typeof input === "object" ? input : {}) as Record<string, unknown>;
  const d = DEFAULT_PROFILE;
  const sections = (p.sections && typeof p.sections === "object" ? p.sections : {}) as Record<
    string,
    unknown
  >;
  const markings = Array.isArray(p.markings)
    ? p.markings
        .filter((m): m is Record<string, unknown> => !!m && typeof m === "object")
        .map((m) => ({ symbol: text(m.symbol, "", 6).trim(), meaning: text(m.meaning, "").trim() }))
        .filter((m) => m.symbol)
        .slice(0, 20)
    : d.markings;
  const size = Number(p.fontSize);
  return {
    firm: text(p.firm, d.firm).trim(),
    font: text(p.font, d.font, 60).trim() || d.font,
    fontSize: Number.isFinite(size) && size >= 6 && size <= 24 ? Math.round(size) : d.fontSize,
    accent: isColor(p.accent) ? p.accent.toUpperCase() : d.accent,
    inputColor: isColor(p.inputColor) ? p.inputColor.toUpperCase() : d.inputColor,
    numbers: p.numbers === "heltal" || p.numbers === "decimaler" ? p.numbers : "auto",
    sections: {
      purpose: text(sections.purpose, d.sections.purpose, 40).trim() || d.sections.purpose,
      work: text(sections.work, d.sections.work, 40).trim() || d.sections.work,
      conclusion:
        text(sections.conclusion, d.sections.conclusion, 40).trim() || d.sections.conclusion,
    },
    markings,
    orientation:
      p.orientation === "liggande" || p.orientation === "stående" ? p.orientation : "auto",
    hideGridlines: typeof p.hideGridlines === "boolean" ? p.hideGridlines : d.hideGridlines,
  };
}

/** En ljus nyans av en färg (för tabellrubriker): andelen vitt, 0–1. */
export function tint(color: string, white: number): string {
  const n = parseInt(color.slice(1), 16);
  const mix = (c: number) => Math.round(c + (255 - c) * white);
  const r = mix(n >> 16);
  const g = mix((n >> 8) & 255);
  const b = mix(n & 255);
  return `#${((r << 16) | (g << 8) | b).toString(16).padStart(6, "0").toUpperCase()}`;
}

const KEY = "kvanto-mallverktyg-profil";

export function loadProfile(): Profile {
  try {
    const saved = localStorage.getItem(KEY);
    return saved ? cleanProfile(JSON.parse(saved)) : DEFAULT_PROFILE;
  } catch {
    return DEFAULT_PROFILE;
  }
}

export function saveProfile(profile: Profile) {
  try {
    localStorage.setItem(KEY, JSON.stringify(profile));
  } catch {
    // Utan lagring gäller inställningarna tills panelen stängs.
  }
}

/** Filen som delas med kollegorna. */
export function exportProfile(profile: Profile): string {
  return JSON.stringify({ typ: "Kvanto Mallverktyg, byråns inställningar", ...profile }, null, 2);
}
