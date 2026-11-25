import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { App } from "./App.tsx";
import { createBrowserHost } from "./excel/browserHost.ts";
import { createExcelHost } from "./excel/excelHost.ts";
import { Preview } from "./Preview.tsx";
import { setUpRibbon } from "./ribbon.ts";
import "./styles.css";

const root = createRoot(document.getElementById("root")!);

function start(inExcel: boolean) {
  if (inExcel) {
    setUpRibbon();
    root.render(
      <StrictMode>
        <App host={createExcelHost()} />
      </StrictMode>,
    );
    return;
  }
  // Utanför Excel: panelen bredvid en låtsas-arbetsbok, för att prova.
  const host = createBrowserHost();
  root.render(
    <StrictMode>
      <div className="browser">
        <div className="browser-pane">
          <App host={host} />
        </div>
        <Preview host={host} />
      </div>
    </StrictMode>,
  );
}

if (typeof Office !== "undefined") {
  Office.onReady((info) => start(info.host === Office.HostType.Excel));
} else {
  start(false);
}
