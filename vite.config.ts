import { defineConfig, type Plugin } from "vite";
import react from "@vitejs/plugin-react";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { CSP, HEADERS } from "./scripts/csp.ts";

/** Säkerhetspolicyn (scripts/csp.ts) i den byggda sidan och som http-huvud hos Cloudflare. */
function contentSecurityPolicy(): Plugin {
  return {
    name: "mallverktyg-csp",
    apply: "build",
    transformIndexHtml: (html) =>
      html.replace(
        "<head>",
        `<head>\n    <meta http-equiv="Content-Security-Policy" content="${CSP}" />`,
      ),
    generateBundle() {
      this.emitFile({ type: "asset", fileName: "_headers", source: HEADERS });
    },
  };
}

// Excel kräver https även lokalt (`npm run certs` installerar certifikatet).
// `npm run dev:browser` kör utan https, för att prova panelen i en vanlig webbläsare.
export default defineConfig(async ({ command, mode }) => ({
  base: "./",
  plugins: [react(), contentSecurityPolicy()],
  server: {
    port: 3100,
    strictPort: true,
    https: command === "serve" && mode !== "browser" ? await getHttpsServerOptions() : undefined,
  },
}));
