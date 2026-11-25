// Mallverktygets säkerhetspolicy (Content Security Policy). Webbläsaren i
// panelen vägrar att skicka något till andra servrar än Mallverktygets egen
// adress: arbetsböckerna kan inte lämna datorn, inte ens av misstag.
// Läggs i den byggda sidan och som http-huvud hos Cloudflare (_headers).
// Försvaga den aldrig.

const POLICY = [
  "default-src 'self'",
  // Microsofts office.js måste hämtas från Microsoft.
  "script-src 'self' https://appsforoffice.microsoft.com",
  "style-src 'self' 'unsafe-inline'",
  "img-src 'self' data: blob:",
  "font-src 'self'",
  // Det viktigaste: bara den egna adressen.
  "connect-src 'self' blob:",
  "worker-src 'none'",
  "frame-src 'none'",
  "object-src 'none'",
  "base-uri 'self'",
  "form-action 'none'",
];

export const CSP = POLICY.join("; ");

export const HEADERS = `# Gjord av vite.config.ts vid bygget. Ändra i scripts/csp.ts.
/*
  Content-Security-Policy: ${CSP}
  Referrer-Policy: no-referrer
  X-Content-Type-Options: nosniff
  Permissions-Policy: camera=(), microphone=(), geolocation=(), payment=(), usb=()
`;
