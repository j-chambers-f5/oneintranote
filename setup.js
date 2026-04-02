#!/usr/bin/env node
/**
 * setup.js — Entra ID app registration setup script
 *
 * This script automates creating a Microsoft Entra ID (formerly Azure AD) app registration
 * for a self-hosted OneIntraNote instance. It uses the Azure CLI (az) to:
 *
 *   1. Create an app registration with multi-tenant support (personal + org accounts)
 *   2. Configure SPA redirect URIs for the web app (localhost dev + deployed URL)
 *   3. Enable public client flows for the CLI tool (publish.js uses localhost redirect)
 *   4. Add delegated OneNote permissions (Notes.Read + Notes.ReadWrite)
 *   5. Set access token version to v2 (required for personal Microsoft accounts)
 *   6. Write a .env file with the app's client ID and authority
 *
 * Why each Entra ID configuration is needed:
 *   - AzureADandPersonalMicrosoftAccount sign-in audience: allows both work/school and
 *     personal Microsoft accounts to sign in (multi-tenant + personal)
 *   - SPA redirect URIs: MSAL.js uses the redirect flow; the URI must be registered
 *   - isFallbackPublicClient + publicClient redirect: publish.js is a native/CLI app
 *     that uses a localhost redirect for OAuth (no client secret)
 *   - requestedAccessTokenVersion: 2 is required for personal Microsoft accounts;
 *     v1 tokens don't work with the "common" or "consumers" endpoints
 *   - Notes.Read (371361e4...): delegated permission to read OneNote content (viewer)
 *   - Notes.ReadWrite (615e26af...): delegated permission to create/update pages (publisher)
 *   - Both are user-consentable — no admin consent required
 *
 * Usage:
 *   node setup.js [deployed-url]
 */
import { execSync } from "child_process";
import { writeFileSync, existsSync } from "fs";

// Optional: pass the deployed URL as an argument to add it as a redirect URI
const DEPLOY_URL = process.argv[2] || "";

function run(cmd) {
  return execSync(cmd, { encoding: "utf-8", stdio: ["pipe", "pipe", "pipe"] }).trim();
}

function runJson(cmd) {
  return JSON.parse(run(cmd));
}

console.log("OneIntraNote Setup\n");

// --- Prerequisite: Azure CLI must be installed and logged in ---
try {
  run("az --version");
} catch {
  console.error("Error: Azure CLI (az) is not installed. Install it from https://aka.ms/install-azure-cli");
  process.exit(1);
}

let account;
try {
  account = runJson("az account show");
} catch {
  console.log("Not logged in. Opening browser to sign in...\n");
  execSync("az login", { stdio: "inherit" });
  account = runJson("az account show");
}

console.log(`Logged in as: ${account.user.name}`);
console.log(`Tenant: ${account.tenantId}\n`);

// --- Build redirect URIs ---
// localhost:5173 is the Vite dev server default.
// The deployed URL is added if provided as a CLI argument.
const redirectUris = ["http://localhost:5173", "https://j-chambers-f5.github.io/oneintranote/"];
if (DEPLOY_URL) {
  const cleaned = DEPLOY_URL.replace(/\/$/, "");
  redirectUris.push(cleaned);
}

// --- Create the Entra ID app registration ---
// "AzureADandPersonalMicrosoftAccount" allows both org and personal Microsoft accounts.
console.log("Creating Entra ID app registration...");
const app = runJson(`az ad app create --display-name "OneIntraNote" --sign-in-audience AzureADandPersonalMicrosoftAccount`);
const appId = app.appId;
console.log(`App ID: ${appId}`);

// --- Configure SPA redirect URIs ---
// These are needed for the MSAL.js redirect-based auth flow in the browser.
const spaJson = JSON.stringify({ redirectUris }).replace(/"/g, '\\"');
run(`az ad app update --id ${appId} --set spa=\\"${spaJson}\\"`);

// Workaround: the az CLI escaping for --set is fragile, so we also try via REST API
try {
  run(`az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/${app.id}" --body '{"spa":${JSON.stringify({ redirectUris })}}'`);
} catch {
  // Fallback: try direct update
  run(`az ad app update --id ${appId} --set 'spa={"redirectUris":${JSON.stringify(redirectUris)}}'`);
}
console.log(`SPA redirect URIs: ${redirectUris.join(", ")}`);

// --- Enable public client flows ---
// This allows publish.js (a CLI/native app) to use the OAuth authorization code flow
// with a localhost redirect URI, without requiring a client secret.
run(`az ad app update --id ${appId} --set isFallbackPublicClient=true --set 'publicClient={"redirectUris":["https://login.microsoftonline.com/common/oauth2/nativeclient"]}'`);
console.log("Public client flows: enabled");

// --- Add OneNote API permissions (delegated, user-consentable) ---
// Set access token version to v2 (required for personal Microsoft accounts)
run(`az rest --method PATCH --uri "https://graph.microsoft.com/v1.0/applications/${app.id}" --body '{"api":{"requestedAccessTokenVersion":2}}'`);

// Permission GUIDs are from the Microsoft Graph API:
// Notes.Read: 371361e4-b9e2-4a3f-8315-2a301a3b0a3d (read notebooks, used by the web viewer)
// Notes.ReadWrite: 615e26af-c38a-4150-ae3e-c3b0d4cb1d6a (create/update pages, used by publish.js)
run(`az ad app permission add --id ${appId} --api 00000003-0000-0000-c000-000000000000 --api-permissions 371361e4-b9e2-4a3f-8315-2a301a3b0a3d=Scope 615e26af-c38a-4150-ae3e-c3b0d4cb1d6a=Scope`);
console.log("API permissions: Notes.Read, Notes.ReadWrite (no admin consent required)");

// --- Write .env file for the Vite dev server ---
const envContent = `VITE_ENTRA_CLIENT_ID=${appId}\nVITE_ENTRA_AUTHORITY=https://login.microsoftonline.com/common\n`;
writeFileSync(".env", envContent);
console.log("\nWrote .env file");

console.log(`
Setup complete!

  App ID: ${appId}
  .env:   written

Next steps:
  npm install
  npm run dev            # start dev server at http://localhost:5173
  npm run build          # build for production

To publish a site:
  node publish.js ./dist "My Site"

To deploy the app:
  npm run build
  swa deploy ./dist
`);
