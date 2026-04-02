/**
 * authConfig.ts — MSAL (Microsoft Authentication Library) configuration
 *
 * This configures the browser-based OAuth 2.0 flow for Microsoft Entra ID.
 *
 * Key design choices:
 *   - "common" authority: supports both personal Microsoft accounts and organizational
 *     (work/school) accounts. Users from any tenant can sign in without pre-configuration.
 *   - localStorage cache: tokens persist across browser sessions so users don't have to
 *     re-authenticate on every visit. (sessionStorage would require login per tab.)
 *   - Notes.Read scope: the minimum permission needed to read OneNote content. This is a
 *     user-consentable delegated permission — no admin consent required, so anyone can
 *     use the app without IT approval.
 *   - SPA redirect: MSAL handles the OAuth redirect flow in the browser. The redirect URI
 *     must match one registered in the Entra ID app registration.
 */
import type { Configuration } from "@azure/msal-browser";
import { LogLevel } from "@azure/msal-browser";

// Client ID from Entra ID app registration. Can be overridden via env var for self-hosted instances.
export const CLIENT_ID = import.meta.env.VITE_ENTRA_CLIENT_ID || "0aca83cc-ae07-4b1f-a62d-0e4a82fa00d4";

// BASE_PATH comes from Vite's base config (e.g. "/" for root domain, "/oneintranote/" for subpath).
// Used to construct redirect URIs and SW registration scope.
export const BASE_PATH = import.meta.env.BASE_URL || "/";

export const msalConfig: Configuration = {
  auth: {
    clientId: CLIENT_ID,
    // "common" authority enables multi-tenant support (personal + org accounts).
    // For single-tenant deployments, override with VITE_ENTRA_AUTHORITY env var.
    authority: import.meta.env.VITE_ENTRA_AUTHORITY || "https://login.microsoftonline.com/common",
    // Redirect URI must match what's registered in the Entra ID app.
    // Uses origin + base path (e.g. "https://example.com/" or "https://example.com/oneintranote/").
    redirectUri: window.location.origin + BASE_PATH,
  },
  cache: {
    // localStorage persists tokens across browser sessions, so the user stays
    // logged in until the token expires or is revoked. This is important for
    // stale-while-revalidate — returning users see cached content instantly.
    cacheLocation: "localStorage",
  },
  system: {
    loggerOptions: {
      logLevel: LogLevel.Warning,
    },
  },
};

// Graph API scopes needed by the app.
// Notes.Read.All allows reading OneNote notebooks across all locations the user
// can access: personal notebooks, shared notebooks, site notebooks, and group notebooks.
// Notes.Read only covers /me/onenote/* (personal notebooks on the user's own OneDrive).
// Notes.Read.All also covers /sites/{id}/onenote/* and /groups/{id}/onenote/*.
// This is a delegated permission — users can consent to it themselves (no admin needed).
export const graphScopes = {
  onenote: ["Notes.Read.All"],
};
