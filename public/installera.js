// Installationssidan: länken till Mallverktygets manifest (för Microsoft 365-
// administrationen) och ett kommando som lägger till Mallverktyget i Excel för
// Windows. Samma sätt som för Kvanto: ingen nedladdad installationsfil, som
// Windows Defender skulle stoppa.

/** Det publicerade Mallverktygets id, samma som PRODUCTION_ID i scripts/manifest.ts. */
export const ADDIN_ID = "8c511835-2eb9-4d7e-b149-c617fc608e15";

const KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";
const DIR = "%LOCALAPPDATA%\\KvantoMallverktyg";
const FILE = `${DIR}\\manifest.xml`;

/** Kommandot för cmd som installerar Mallverktyget från `origin`. */
export function installCommand(origin) {
  return (
    `mkdir "${DIR}" 2>nul & ` +
    `curl -fsSL ${origin}/mallverktyg-manifest.xml -o "${FILE}" && ` +
    `reg add "${KEY}" /v ${ADDIN_ID} /t REG_SZ /d "${FILE}" /f`
  );
}

/** Kommandot som tar bort Mallverktyget ur Excel igen. */
export function uninstallCommand() {
  return `reg delete "${KEY}" /v ${ADDIN_ID} /f & rmdir /s /q "${DIR}"`;
}

/** Fyller i kommandona på sidan och gör kopieringsknapparna klickbara. */
function setUp() {
  const commands = {
    manifest: `${location.origin}/mallverktyg-manifest.xml`,
    install: installCommand(location.origin),
    uninstall: uninstallCommand(),
  };
  for (const [name, command] of Object.entries(commands)) {
    const code =
      document.getElementById(`${name}-command`) ?? document.getElementById(`${name}-url`);
    const button = document.getElementById(`${name}-copy`);
    if (!code || !button) continue;
    code.textContent = command;
    button.addEventListener("click", async () => {
      try {
        await navigator.clipboard.writeText(command);
      } catch {
        const range = document.createRange();
        range.selectNodeContents(code);
        getSelection()?.removeAllRanges();
        getSelection()?.addRange(range);
      }
      const text = button.textContent;
      button.textContent = "Kopierat ✓";
      setTimeout(() => (button.textContent = text), 2000);
    });
  }
}

if (typeof document !== "undefined") setUp();
