import { useRef, useState } from "react";
import {
  DEFAULT_PROFILE,
  FONTS,
  cleanProfile,
  exportProfile,
  tint,
  type Profile,
} from "../lib/profile.ts";

interface Props {
  profile: Profile;
  onChange: (profile: Profile) => void;
}

/** Byråns inställningar: utseende, avsnitt och markeringar. */
export function ProfileView({ profile, onChange }: Props) {
  const file = useRef<HTMLInputElement>(null);
  const [message, setMessage] = useState("");
  const [confirmReset, setConfirmReset] = useState(false);
  /** Senast valda bakgrund, så att den kommer tillbaka när bakgrunden slås på igen. */
  const [lastBackground, setLastBackground] = useState(
    profile.background || DEFAULT_PROFILE.background,
  );
  const set = (patch: Partial<Profile>) => onChange({ ...profile, ...patch });
  const setMarking = (i: number, patch: Partial<Profile["markings"][number]>) =>
    set({ markings: profile.markings.map((m, j) => (j === i ? { ...m, ...patch } : m)) });

  function download() {
    const link = document.createElement("a");
    link.href = URL.createObjectURL(
      new Blob([exportProfile(profile)], { type: "application/json" }),
    );
    link.download = "Mallverktyg - byråns inställningar.json";
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 10000);
    setMessage("Filen är sparad i Hämtade filer. Skicka den till kollegorna.");
  }

  async function upload(chosen: File | undefined) {
    if (!chosen) return;
    try {
      onChange(cleanProfile(JSON.parse(await chosen.text())));
      setMessage("Byråns inställningar är inlästa.");
    } catch {
      setMessage("Filen gick inte att läsa. Välj filen med byråns inställningar (.json).");
    }
  }

  return (
    <>
      <section>
        <h2>Byrån</h2>
        <label className="field">
          Byråns namn
          <input
            value={profile.firm}
            onChange={(e) => set({ firm: e.target.value })}
            placeholder="Revisionsbyrån AB"
          />
        </label>
        <p className="muted small">
          Inställningarna gäller flikar som anpassas från och med nu. Anpassa igen för att uppdatera
          gamla.
        </p>
      </section>

      <section>
        <h2>Utseende</h2>
        <div className="row">
          <label className="field">
            Typsnitt
            <select value={profile.font} onChange={(e) => set({ font: e.target.value })}>
              {[...new Set([profile.font, ...FONTS])].map((f) => (
                <option key={f}>{f}</option>
              ))}
            </select>
          </label>
          <label className="field narrow">
            Storlek
            <input
              type="number"
              min={6}
              max={24}
              value={profile.fontSize}
              onChange={(e) =>
                set({ fontSize: cleanProfile({ fontSize: e.target.value }).fontSize })
              }
            />
          </label>
        </div>
        <div className="row">
          <label className="field narrow">
            Byråns färg
            <input
              type="color"
              value={profile.accent}
              onChange={(e) => set({ accent: e.target.value.toUpperCase() })}
            />
          </label>
          <label className="field narrow">
            Fyll i-rutor
            <input
              type="color"
              value={profile.inputColor}
              onChange={(e) => set({ inputColor: e.target.value.toUpperCase() })}
            />
          </label>
          <label className="field narrow">
            Bakgrund
            <input
              type="color"
              value={profile.background || lastBackground}
              disabled={!profile.background}
              onChange={(e) => {
                setLastBackground(e.target.value.toUpperCase());
                set({ background: e.target.value.toUpperCase() });
              }}
            />
          </label>
          <div className="sample" aria-hidden style={{ fontFamily: profile.font }}>
            <span
              style={{
                color: profile.accent,
                fontWeight: 600,
                background: profile.background || undefined,
              }}
            >
              {profile.firm || "Byrån"}
            </span>
            <span
              style={{
                background: tint(profile.accent, 0.85),
                borderBottom: `1px solid ${profile.accent}`,
              }}
            >
              Tabellrubrik
            </span>
            <span style={{ background: profile.inputColor }}>Fyll i</span>
          </div>
        </div>
        <label className="check">
          <input
            type="checkbox"
            checked={!!profile.background}
            onChange={(e) => set({ background: e.target.checked ? lastBackground : "" })}
          />
          Bakgrund i det Mallverktyget lägger till (huvud, slutsats och markeringar), så att det
          ursprungliga innehållet syns tydligt
        </label>
        <label className="field">
          Tal utan format
          <select
            value={profile.numbers}
            onChange={(e) => set({ numbers: e.target.value as Profile["numbers"] })}
          >
            <option value="auto">1 234 eller 1 234,50, efter kolumnen</option>
            <option value="heltal">Alltid heltal: 1 234</option>
            <option value="decimaler">Alltid två decimaler: 1 234,50</option>
          </select>
        </label>
        <div className="row">
          <label className="field">
            Utskrift
            <select
              value={profile.orientation}
              onChange={(e) => set({ orientation: e.target.value as Profile["orientation"] })}
            >
              <option value="auto">Liggande när tabellen är bred</option>
              <option value="stående">Alltid stående</option>
              <option value="liggande">Alltid liggande</option>
            </select>
          </label>
        </div>
        <label className="check">
          <input
            type="checkbox"
            checked={profile.hideGridlines}
            onChange={(e) => set({ hideGridlines: e.target.checked })}
          />
          Dölj stödlinjerna, så att flikarna ser ut som papper
        </label>
      </section>

      <section>
        <h2>Avsnitt</h2>
        <div className="row">
          {(
            [
              ["purpose", "Överst, första"],
              ["work", "Överst, andra"],
              ["conclusion", "Efter tabellen"],
            ] as [keyof Profile["sections"], string][]
          ).map(([key, where]) => (
            <label key={key} className="field">
              {where}
              <input
                value={profile.sections[key]}
                onChange={(e) => set({ sections: { ...profile.sections, [key]: e.target.value } })}
              />
            </label>
          ))}
        </div>
      </section>

      <section>
        <h2>Markeringar</h2>
        <p className="muted small">Förklaringen sist på varje flik.</p>
        <ul className="markings">
          {profile.markings.map((m, i) => (
            <li key={i}>
              <input
                className="symbol"
                value={m.symbol}
                maxLength={6}
                aria-label="Markering"
                onChange={(e) => setMarking(i, { symbol: e.target.value })}
              />
              <input
                value={m.meaning}
                aria-label="Betydelse"
                onChange={(e) => setMarking(i, { meaning: e.target.value })}
              />
              <button
                className="icon-button"
                aria-label={`Ta bort ${m.symbol}`}
                onClick={() => set({ markings: profile.markings.filter((_, j) => j !== i) })}
              >
                ×
              </button>
            </li>
          ))}
        </ul>
        <button
          onClick={() => set({ markings: [...profile.markings, { symbol: "", meaning: "" }] })}
        >
          Lägg till markering
        </button>
      </section>

      <section>
        <h2>Dela med kollegorna</h2>
        <p className="muted small">
          Spara inställningarna som en fil och läs in den hos kollegorna, så ser allas
          revisionspapper likadana ut.
        </p>
        <div className="row">
          <button onClick={download}>Spara som fil</button>
          <button onClick={() => file.current?.click()}>Läs in fil</button>
          <input
            ref={file}
            type="file"
            accept=".json,application/json"
            hidden
            onChange={(e) => {
              void upload(e.target.files?.[0]);
              e.target.value = "";
            }}
          />
        </div>
        {message && <p className="small">{message}</p>}
        {confirmReset ? (
          <p className="confirm">
            Återställ allt utom byråns namn?{" "}
            <button
              className="danger"
              onClick={() => {
                onChange({ ...DEFAULT_PROFILE, firm: profile.firm });
                setConfirmReset(false);
              }}
            >
              Ja, återställ
            </button>{" "}
            <button onClick={() => setConfirmReset(false)}>Avbryt</button>
          </p>
        ) : (
          <button className="link" onClick={() => setConfirmReset(true)}>
            Återställ förslaget
          </button>
        )}
      </section>
    </>
  );
}
