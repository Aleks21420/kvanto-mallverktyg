import { useCallback, useEffect, useRef, useState } from "react";
import type { Engagement, Host, SheetStatus, SignKind } from "./excel/host.ts";
import type { Options } from "./lib/layout.ts";
import { loadProfile, saveProfile, type Profile } from "./lib/profile.ts";
import { toTemplate } from "./lib/template.ts";
import { onRibbon, showPane, type RibbonCommand } from "./ribbon.ts";
import { ProfileView } from "./views/ProfileView.tsx";

type Tab = "anpassa" | "mall" | "byran";

const INITIALS_KEY = "kvanto-mallverktyg-initialer";
const OPTIONS_KEY = "kvanto-mallverktyg-val";

function stored<T>(key: string, fallback: T): T {
  try {
    const value = localStorage.getItem(key);
    return value ? { ...fallback, ...JSON.parse(value) } : fallback;
  } catch {
    return fallback;
  }
}
function store(key: string, value: unknown) {
  try {
    localStorage.setItem(key, typeof value === "string" ? value : JSON.stringify(value));
  } catch {
    // Bara en bekvämlighet.
  }
}

/** Ett fel från Excel som text att visa. */
function problem(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (/protect/i.test(message))
    return "Fliken är skyddad. Ta bort skyddet (Granska → Ta bort skydd) och försök igen.";
  if (/edit mode|InvalidOperationInCellEditMode/i.test(message))
    return "Excel håller på att redigera en cell. Tryck Enter eller Esc i Excel och försök igen.";
  return `Det gick inte: ${message}`;
}

export function App({ host }: { host: Host }) {
  const [tab, setTab] = useState<Tab>("anpassa");
  const [profile, setProfile] = useState<Profile>(loadProfile);
  const [engagement, setEngagement] = useState<Engagement>({ client: "", orgNr: "", year: "" });
  const [engagementDirty, setEngagementDirty] = useState(false);
  const [sheets, setSheets] = useState<SheetStatus[]>([]);
  const [active, setActive] = useState("");
  const [selected, setSelected] = useState<Set<string>>(new Set());
  const [options, setOptions] = useState<Options>(() =>
    stored(OPTIONS_KEY, { style: true, print: true }),
  );
  const [initials, setInitials] = useState(() => {
    try {
      return localStorage.getItem(INITIALS_KEY) ?? "";
    } catch {
      return "";
    }
  });
  const [status, setStatus] = useState<{ text: string; kind: "ok" | "error" } | null>(null);
  const [busy, setBusy] = useState(false);
  const [templateName, setTemplateName] = useState("Revisionspapper");
  const [confirmClear, setConfirmClear] = useState(false);

  const refresh = useCallback(async () => {
    try {
      const [list, current] = await Promise.all([host.sheets(), host.activeSheet()]);
      setSheets(list);
      setActive(current);
    } catch {
      // Excel kan vara upptaget en kort stund; nästa ändring läser in igen.
    }
  }, [host]);

  useEffect(() => {
    void refresh();
    void host
      .getEngagement()
      .then(setEngagement)
      .catch(() => {});
    let timer: ReturnType<typeof setTimeout> | undefined;
    const off = host.onChange(() => {
      clearTimeout(timer);
      timer = setTimeout(() => void refresh(), 300);
    });
    return () => {
      off();
      clearTimeout(timer);
    };
  }, [host, refresh]);

  const ok = (text: string) => setStatus({ text, kind: "ok" });
  const fail = (text: string) => setStatus({ text, kind: "error" });

  async function run(work: () => Promise<void>) {
    setBusy(true);
    setStatus(null);
    try {
      await work();
    } catch (error) {
      fail(problem(error));
    } finally {
      setBusy(false);
      void refresh();
    }
  }

  async function saveEngagement() {
    await host.setEngagement(engagement, profile);
    setEngagementDirty(false);
  }

  async function adapt(names: string[]) {
    if (!names.length) return;
    if (engagementDirty) await saveEngagement();
    else await host.setEngagement(engagement, profile);
    for (const name of names) await host.adapt(name, profile, options);
    ok(
      names.length === 1
        ? `Fliken ${names[0]} är anpassad.`
        : `${names.length} flikar är anpassade.`,
    );
  }

  async function sign(kind: SignKind, sheet = active): Promise<boolean> {
    if (!initials.trim()) {
      setTab("anpassa");
      showPane();
      fail("Skriv dina initialer under Signera först.");
      return false;
    }
    const done = await host.sign(sheet, kind, initials.trim());
    if (!done) {
      showPane();
      fail(`Fliken ${sheet} har inget huvud från Mallverktyget än. Anpassa den först.`);
      return false;
    }
    ok(
      `${kind === "prepared" ? "Upprättat" : "Granskat"}: ${initials.trim()}, ${new Date().toLocaleDateString("sv-SE")} (${sheet}).`,
    );
    return true;
  }

  async function saveTemplate() {
    const bytes = await host.file();
    if (!bytes) {
      fail("Att spara som mall fungerar bara i Excel, inte i provversionen i webbläsaren.");
      return;
    }
    const template = await toTemplate(bytes);
    const link = document.createElement("a");
    link.href = URL.createObjectURL(
      new Blob([template as BlobPart], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.template",
      }),
    );
    link.download = `${templateName.trim().replace(/[\\/:*?"<>|]/g, "") || "Mall"}.xltx`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 10000);
    ok(`Mallen ${link.download} är sparad i Hämtade filer.`);
  }

  // Knapparna i menyfliksområdet.
  const latest = useRef({ adapt, sign, active });
  latest.current = { adapt, sign, active };
  useEffect(
    () =>
      onRibbon(async (command: RibbonCommand) => {
        const { adapt, sign } = latest.current;
        const current = await host.activeSheet();
        setActive(current);
        if (command === "template") {
          setTab("mall");
          showPane();
          return;
        }
        await run(async () => {
          if (command === "adapt") await adapt([current]);
          else await sign(command, current);
        });
      }),
    [host],
  );

  function changeProfile(next: Profile) {
    setProfile(next);
    saveProfile(next);
  }

  const toggle = (name: string) => {
    const next = new Set(selected);
    if (next.has(name)) next.delete(name);
    else next.add(name);
    setSelected(next);
  };
  const changeOptions = (next: Options) => {
    setOptions(next);
    store(OPTIONS_KEY, next);
  };
  const edit = (key: keyof Engagement) => (e: React.ChangeEvent<HTMLInputElement>) => {
    setEngagement({ ...engagement, [key]: e.target.value });
    setEngagementDirty(true);
  };

  return (
    <div className="pane">
      <header className="pane-head">
        <img src="./icons/icon-32.png" alt="" width={24} height={24} />
        <h1>Mallverktyg</h1>
      </header>
      <nav className="tabs" role="tablist">
        {(
          [
            ["anpassa", "Anpassa"],
            ["mall", "Mall"],
            ["byran", "Byrån"],
          ] as [Tab, string][]
        ).map(([id, label]) => (
          <button
            key={id}
            role="tab"
            aria-selected={tab === id}
            className={tab === id ? "active" : ""}
            onClick={() => {
              setTab(id);
              setStatus(null);
            }}
          >
            {label}
          </button>
        ))}
      </nav>

      {status && (
        <p className={`status ${status.kind}`} role="status">
          {status.text}
        </p>
      )}
      {!profile.firm && tab !== "byran" && (
        <p className="notice">
          Börja med att skriva byråns namn under{" "}
          <button className="link" onClick={() => setTab("byran")}>
            Byrån
          </button>
          .
        </p>
      )}

      {tab === "anpassa" && (
        <>
          <section>
            <h2>Uppdraget</h2>
            <p className="muted small">
              Gäller hela arbetsboken. Alla flikars huvud hämtar uppgifterna härifrån.
            </p>
            <label className="field">
              Klient
              <input
                value={engagement.client}
                onChange={edit("client")}
                placeholder="Företaget AB"
              />
            </label>
            <div className="row">
              <label className="field">
                Org.nr
                <input
                  value={engagement.orgNr}
                  onChange={edit("orgNr")}
                  placeholder="556677-8899"
                />
              </label>
              <label className="field">
                Räkenskapsår
                <input
                  value={engagement.year}
                  onChange={edit("year")}
                  placeholder="2026-01-01 – 2026-12-31"
                />
              </label>
            </div>
            {engagementDirty && (
              <button
                disabled={busy}
                onClick={() =>
                  void run(async () => {
                    await saveEngagement();
                    ok("Uppdragets uppgifter är sparade i arbetsboken.");
                  })
                }
              >
                Spara uppgifterna
              </button>
            )}
          </section>

          <section>
            <h2>Flikar</h2>
            <ul className="sheets">
              {sheets.map((s) => (
                <li key={s.name} className={s.name === active ? "current" : ""}>
                  <label>
                    <input
                      type="checkbox"
                      checked={selected.has(s.name)}
                      onChange={() => toggle(s.name)}
                    />
                    <span className="sheet-name">{s.name}</span>
                  </label>
                  <span className="sheet-status">
                    {s.adapted ? (
                      <>
                        {s.reference && <b>{s.reference} </b>}
                        <span title="Upprättat" className={s.prepared.by ? "done" : "todo"}>
                          U {s.prepared.by || "–"}
                        </span>{" "}
                        <span title="Granskat" className={s.reviewed.by ? "done" : "todo"}>
                          G {s.reviewed.by || "–"}
                        </span>
                      </>
                    ) : (
                      <span className="muted">Inte anpassad</span>
                    )}
                  </span>
                </li>
              ))}
            </ul>
            <div className="row">
              <button
                className="primary"
                disabled={busy || !active}
                onClick={() => void run(() => adapt([active]))}
              >
                Anpassa {active ? `”${active}”` : "fliken"}
              </button>
              <button
                disabled={busy || !selected.size}
                onClick={() => void run(() => adapt([...selected]))}
              >
                Anpassa markerade ({selected.size})
              </button>
            </div>
            <label className="check">
              <input
                type="checkbox"
                checked={options.style}
                onChange={(e) => changeOptions({ ...options, style: e.target.checked })}
              />
              Byråns typsnitt, talformat och tabellrubrik på flikens innehåll
            </label>
            <label className="check">
              <input
                type="checkbox"
                checked={options.print}
                onChange={(e) => changeOptions({ ...options, print: e.target.checked })}
              />
              Utskrift: en sida bred, sidfot med byrå, flik och sidnummer
            </label>
            <p className="muted small">
              Går att göra flera gånger: huvudet uppdateras, men det som är ifyllt står kvar och
              inget läggs till två gånger.
            </p>
          </section>

          <section>
            <h2>Signera</h2>
            <div className="row">
              <label className="field narrow">
                Dina initialer
                <input
                  value={initials}
                  maxLength={6}
                  onChange={(e) => {
                    setInitials(e.target.value);
                    store(INITIALS_KEY, e.target.value);
                  }}
                  placeholder="AB"
                />
              </label>
              <button
                disabled={busy || !active}
                onClick={() => void run(async () => void (await sign("prepared")))}
              >
                Upprättat
              </button>
              <button
                disabled={busy || !active}
                onClick={() => void run(async () => void (await sign("reviewed")))}
              >
                Granskat
              </button>
            </div>
            <p className="muted small">
              Fyller i initialer och dagens datum i den aktiva flikens huvud.
            </p>
          </section>
        </>
      )}

      {tab === "mall" && (
        <>
          <section>
            <h2>Spara som mall</h2>
            <p>
              En mall (.xltx) är en arbetsbok som man börjar nya arbetspapper från. Öppnar man
              mallen blir det en ny, namnlös arbetsbok, och mallen själv ändras inte.
            </p>
            <ol className="steps">
              <li>
                Anpassa flikarna och skriv in det som ska stå i alla uppdrag, t.ex. syfte och utfört
                arbete.
              </li>
              <li>
                Töm klientens uppgifter: klient, org.nr, räkenskapsår, signeringar och slutsatser.
                <br />
                {/* Office-tillägg kan inte visa confirm(): bekräftelsen görs här. */}
                {confirmClear ? (
                  <span className="confirm">
                    Töm i hela arbetsboken?{" "}
                    <button
                      className="danger"
                      disabled={busy}
                      onClick={() => {
                        setConfirmClear(false);
                        void run(async () => {
                          await host.clearClientData();
                          setEngagement({ client: "", orgNr: "", year: "" });
                          setEngagementDirty(false);
                          ok("Klientens uppgifter är tömda.");
                        });
                      }}
                    >
                      Ja, töm
                    </button>{" "}
                    <button onClick={() => setConfirmClear(false)}>Avbryt</button>
                  </span>
                ) : (
                  <button disabled={busy} onClick={() => setConfirmClear(true)}>
                    Töm klientens uppgifter
                  </button>
                )}
              </li>
              <li>Spara arbetsboken (Ctrl + S).</li>
              <li>
                <label className="field">
                  Mallens namn
                  <input value={templateName} onChange={(e) => setTemplateName(e.target.value)} />
                </label>
                <button className="primary" disabled={busy} onClick={() => void run(saveTemplate)}>
                  Spara som mall (.xltx)
                </button>
              </li>
            </ol>
            <p className="muted small">
              Lägg mallen i byråns gemensamma mapp. Dubbelklicka på den för att börja ett nytt
              arbetspapper.
            </p>
          </section>
        </>
      )}

      {tab === "byran" && <ProfileView profile={profile} onChange={changeProfile} />}
    </div>
  );
}
