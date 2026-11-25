// Knapparna på fliken Mallverktyg i Excels menyfliksområde (manifest.template.xml).
// Panelen och knapparna delar körmiljö (SharedRuntime 1.1), så en knapp anropar
// panelens kod direkt, även när panelen är stängd.

export type RibbonCommand = "adapt" | "prepared" | "reviewed" | "template";

/** Funktionsnamnen i manifestet och vad de gör. */
const FUNCTIONS: Record<string, RibbonCommand> = {
  mallAnpassa: "adapt",
  mallUpprattat: "prepared",
  mallGranskat: "reviewed",
  mallMall: "template",
};

const listeners = new Set<(command: RibbonCommand) => Promise<void>>();

function hasSharedRuntime(): boolean {
  try {
    return Office.context.requirements.isSetSupported("SharedRuntime", "1.1");
  } catch {
    return false;
  }
}

/** Kopplar knapparna till panelen. Anropas en gång, när Office är redo. */
export function setUpRibbon() {
  if (!hasSharedRuntime()) return;
  for (const [name, command] of Object.entries(FUNCTIONS)) {
    Office.actions.associate(name, async (event: Office.AddinCommands.Event) => {
      try {
        for (const listener of listeners) await listener(command);
      } finally {
        event.completed();
      }
    });
  }
}

/** Visar panelen, om den är stängd. */
export function showPane() {
  if (hasSharedRuntime()) Office.addin.showAsTaskpane().catch(() => {});
}

/** Lyssnar på knapparna. Returnerar en funktion som slutar lyssna. */
export function onRibbon(listener: (command: RibbonCommand) => Promise<void>): () => void {
  listeners.add(listener);
  return () => void listeners.delete(listener);
}
