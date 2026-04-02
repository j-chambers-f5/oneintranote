#!/usr/bin/env node
/**
 * publish.js — OneIntraNote CLI publishing tool
 *
 * This script publishes a static site directory to a OneNote notebook.
 * It zips the directory, authenticates the user, finds (or creates) the target
 * notebook and "Sites" section, and uploads the zip as a page attachment.
 *
 * The uploaded page can then be viewed via the OneIntraNote web app, which
 * downloads the zip, unpacks it in the browser, and serves it via a Service Worker.
 *
 * Authentication:
 *   - Uses browser-based OAuth 2.0 with PKCE (Proof Key for Code Exchange)
 *   - Starts a local HTTP server on port 53682 to receive the OAuth redirect
 *   - Opens the user's browser to the Microsoft login page
 *   - Exchanges the auth code for tokens using the PKCE code verifier
 *   - Caches tokens (including refresh_token) at ~/.oneintranote/token.json
 *   - On subsequent runs, silently refreshes the token without opening the browser
 *
 * Usage:
 *   node publish.js <directory> <onenote-url> [site-name]
 *   node publish.js <directory> <site-name> [notebook-name]
 */
import { readFileSync, writeFileSync, existsSync, statSync, mkdirSync } from "fs";
import { execSync } from "child_process";
import { resolve, join } from "path";
import { createServer } from "http";
import { homedir } from "os";

// --- Configuration ---
// The same Entra ID app registration is used for both the web app and the CLI.
// Can be overridden via environment variables for self-hosted instances.
const APP_ID = process.env.ONEINTRANOTE_CLIENT_ID || "0aca83cc-ae07-4b1f-a62d-0e4a82fa00d4";

// "common" supports both personal and org accounts. Auto-detected to "consumers"
// for personal OneDrive URLs. Can be overridden via ONEINTRANOTE_TENANT_ID.
let TENANT_ID = process.env.ONEINTRANOTE_TENANT_ID || "common";

// Notes.ReadWrite.All is needed to create/update pages in site notebooks.
// openid and profile are needed to extract the user's identity from the id_token.
const SCOPES = "Notes.ReadWrite.All openid profile";

// Convention: all sites are stored in a section named "Sites" within the notebook.
const SECTION_NAME = "Sites";

// Base URL of the deployed OneIntraNote web viewer.
const VIEWER_URL = "https://j-chambers-f5.github.io/oneintranote";

// --- OneNote URL parsing ---
// When the user pastes a OneNote URL, we extract the notebook GUID and optional page name.
// This avoids requiring users to manually look up notebook IDs.
//
// Supported URL formats:
//   Business: https://...sharepoint.com/...?sourcedoc={GUID}&wd=target(Section.one|.../PageName|...)
//   Personal: https://onedrive.live.com/...?sourcedoc={GUID}&wd=target(Section.one|/PageName|...)
function parseOneNoteUrl(url) {
  const result = { notebookGuid: null, pageName: null };

  // Decode URL-encoded characters first
  const decoded = decodeURIComponent(url);

  // Extract sourcedoc GUID (this identifies the notebook in OneDrive/SharePoint)
  const sourcedoc = decoded.match(/sourcedoc=\{?([a-f0-9-]+)\}?/i);
  if (sourcedoc) {
    result.notebookGuid = sourcedoc[1].toLowerCase();
  }

  // Extract page name from wd=target(Section.one|.../PageName|...)
  const wd = decoded.match(/target\([^|]+\|[^|]*?([^|/]+)\|/);
  if (wd) {
    result.pageName = wd[1];
  }

  return result;
}

// --- Token caching ---
// Tokens are stored at ~/.oneintranote/token.json so the user only needs to
// authenticate via browser once. Subsequent runs use the refresh_token.
const TOKEN_DIR = join(homedir(), ".oneintranote");
const TOKEN_FILE = join(TOKEN_DIR, "token.json");

// OAuth redirect configuration for the local HTTP server.
// The Entra ID app registration must include this URI as a "Mobile and desktop" redirect.
const REDIRECT_PORT = 53682;
const REDIRECT_URI = `http://localhost:${REDIRECT_PORT}/callback`;

function loadCachedToken() {
  try {
    if (existsSync(TOKEN_FILE)) return JSON.parse(readFileSync(TOKEN_FILE, "utf-8"));
  } catch { /* ignore */ }
  return null;
}

function saveCachedToken(tokenData) {
  if (!existsSync(TOKEN_DIR)) mkdirSync(TOKEN_DIR, { recursive: true });
  writeFileSync(TOKEN_FILE, JSON.stringify(tokenData, null, 2));
}

/** Attempt to refresh an expired access token using the stored refresh_token. */
async function refreshToken(refreshToken) {
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: `client_id=${APP_ID}&grant_type=refresh_token&refresh_token=${refreshToken}&scope=${encodeURIComponent(SCOPES)}`,
  });
  const data = await res.json();
  if (data.error) return null;
  return data;
}

/**
 * Browser-based OAuth 2.0 authentication with PKCE.
 *
 * PKCE (Proof Key for Code Exchange) is required because this is a public client
 * (no client secret). The flow:
 *   1. Generate a random code_verifier and compute its SHA-256 hash (code_challenge)
 *   2. Open the browser to the authorization URL with the code_challenge
 *   3. User signs in and consents; Entra ID redirects to localhost with an auth code
 *   4. Exchange the auth code + code_verifier for access/refresh tokens
 *
 * The local HTTP server on port 53682 receives the redirect and extracts the code.
 */
async function browserAuth() {
  const { createHash, randomBytes } = await import("crypto");
  // Generate PKCE parameters: random verifier + its SHA-256 hash as the challenge
  const state = randomBytes(16).toString("hex");
  const codeVerifier = randomBytes(32).toString("base64url");
  const codeChallenge = createHash("sha256").update(codeVerifier).digest("base64url");

  // Build the authorization URL with PKCE challenge and offline_access for refresh tokens
  const authUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize?client_id=${APP_ID}&response_type=code&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&scope=${encodeURIComponent(SCOPES + " offline_access")}&state=${state}&response_mode=query&code_challenge=${codeChallenge}&code_challenge_method=S256`;

  return new Promise((resolveAuth, rejectAuth) => {
    // Start a local HTTP server to receive the OAuth redirect
    const server = createServer(async (req, res) => {
      const url = new URL(req.url, `http://localhost:${REDIRECT_PORT}`);
      if (!url.pathname.startsWith("/callback")) { res.writeHead(404); res.end(); return; }

      const code = url.searchParams.get("code");
      const error = url.searchParams.get("error");

      if (error) {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end("<html><body><h2>Authentication failed</h2><p>You can close this window.</p></body></html>");
        server.close();
        rejectAuth(new Error(url.searchParams.get("error_description") || error));
        return;
      }

      if (code) {
        res.writeHead(200, { "Content-Type": "text/html" });
        res.end("<html><body><h2>Authenticated!</h2><p>You can close this window.</p></body></html>");
        server.close();

        // Exchange the authorization code for tokens, proving possession of the code_verifier
        const tokenRes = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
          method: "POST",
          headers: { "Content-Type": "application/x-www-form-urlencoded" },
          body: `client_id=${APP_ID}&grant_type=authorization_code&code=${code}&redirect_uri=${encodeURIComponent(REDIRECT_URI)}&scope=${encodeURIComponent(SCOPES + " offline_access")}&code_verifier=${codeVerifier}`,
        });
        const tokenData = await tokenRes.json();
        if (tokenData.error) { rejectAuth(new Error(tokenData.error_description)); return; }
        resolveAuth(tokenData);
      }
    });

    server.listen(REDIRECT_PORT, () => {
      console.log("Opening browser for authentication...");
      // Open the default browser to the authorization URL
      const open = process.platform === "darwin" ? "open" : process.platform === "win32" ? "start" : "xdg-open";
      execSync(`${open} "${authUrl}"`);
    });

    // Timeout after 2 minutes if the user doesn't complete sign-in
    setTimeout(() => { server.close(); rejectAuth(new Error("Authentication timed out")); }, 120000);
  });
}

/**
 * Extract the user's identity (email/name) from the JWT id_token.
 * Used for display purposes only ("Signed in as: user@example.com").
 */
function getIdentityFromToken(tokenData) {
  const idt = tokenData.id_token;
  if (!idt) return null;
  try {
    const parts = idt.split(".");
    const payload = parts[1] + "=".repeat(4 - (parts[1].length % 4));
    const claims = JSON.parse(Buffer.from(payload, "base64url").toString());
    return claims.preferred_username || claims.email || claims.name || null;
  } catch { return null; }
}

/**
 * Get an access token, trying cached refresh first, then falling back to browser auth.
 * This is the main entry point for authentication in the CLI.
 */
async function getAccessToken() {
  // Try cached refresh token first (avoids opening the browser)
  const cached = loadCachedToken();
  if (cached?.refresh_token) {
    const identity = getIdentityFromToken(cached);
    console.log(`Refreshing saved token${identity ? ` for ${identity}` : ""}...`);
    const refreshed = await refreshToken(cached.refresh_token);
    if (refreshed?.access_token) {
      // Preserve id_token from original auth if refresh doesn't return one
      if (!refreshed.id_token && cached.id_token) refreshed.id_token = cached.id_token;
      saveCachedToken(refreshed);
      return refreshed.access_token;
    }
    console.log("Refresh failed, need to re-authenticate.");
  }

  // No cached token or refresh failed — open browser for interactive auth
  const tokenData = await browserAuth();
  const identity = getIdentityFromToken(tokenData);
  if (identity) console.log(`Signed in as: ${identity}`);
  saveCachedToken(tokenData);
  return tokenData.access_token;
}

// --- OneNote Graph API helpers ---

/**
 * Find a notebook by matching its Graph API ID against a GUID extracted from a OneNote URL.
 * Business notebook IDs look like "1-{guid-with-dashes}" and personal IDs look like
 * "0-{hex}!s{guid-no-dashes}", so we match both with and without dashes.
 */
async function findNotebookByGuid(token, guid) {
  const headers = { Authorization: `Bearer ${token}` };
  let url = "https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName,id&$top=50";
  while (url) {
    const res = await fetch(url, { headers });
    const data = await res.json();
    for (const nb of data.value || []) {
      const nbIdLower = nb.id.toLowerCase();
      const guidNoDashes = guid.replace(/-/g, "");
      if (nbIdLower.includes(guid) || nbIdLower.includes(guidNoDashes)) return nb;
    }
    url = data["@odata.nextLink"] || null;
  }
  return null;
}

/** Find a notebook by display name, or create it if it doesn't exist. */
async function findOrCreateNotebook(token, name) {
  const headers = { Authorization: `Bearer ${token}` };
  let res = await fetch(`https://graph.microsoft.com/v1.0/me/onenote/notebooks?$filter=displayName eq '${name.replace(/'/g, "''")}'`, { headers });
  let data = await res.json();
  if (data.value?.length) return data.value[0];

  console.log(`Creating "${name}" notebook...`);
  res = await fetch("https://graph.microsoft.com/v1.0/me/onenote/notebooks", {
    method: "POST", headers: { ...headers, "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: name }),
  });
  return await res.json();
}

/** Returns the Graph API OneNote base URL. */
function oneNoteBase(graphSiteId) {
  return graphSiteId
    ? `https://graph.microsoft.com/v1.0/sites/${graphSiteId}/onenote`
    : "https://graph.microsoft.com/v1.0/me/onenote";
}

/** Find the "Sites" section in a notebook, or create it if it doesn't exist. */
async function findOrCreateSection(token, notebookId, graphSiteId) {
  const headers = { Authorization: `Bearer ${token}` };
  const base = oneNoteBase(graphSiteId);
  let res = await fetch(`${base}/notebooks/${notebookId}/sections`, { headers });
  let data = await res.json();
  let section = data.value.find((s) => s.displayName === SECTION_NAME);
  if (section) return section.id;

  res = await fetch(`${base}/notebooks/${notebookId}/sections`, {
    method: "POST", headers: { ...headers, "Content-Type": "application/json" },
    body: JSON.stringify({ displayName: SECTION_NAME }),
  });
  data = await res.json();
  return data.id;
}

/**
 * Upload a site zip to a OneNote page.
 *
 * If a page with this name already exists, it's deleted first (there's no
 * in-place update for page attachments in the OneNote API).
 *
 * The upload uses the OneNote multipart page creation API:
 *   - Part 1 ("Presentation"): HTML content with an <object> tag referencing the zip
 *   - Part 2 ("site.zip"): the raw zip bytes
 *
 * The <object> tag in the HTML uses data="name:site.zip" which tells the OneNote API
 * to attach the binary data from the "site.zip" form part to this element.
 */
async function uploadSite(token, sectionId, siteName, zipBuffer, graphSiteId) {
  const headers = { Authorization: `Bearer ${token}` };
  const base = oneNoteBase(graphSiteId);

  // Delete existing page with same name (can't update attachments in-place)
  const res = await fetch(`${base}/sections/${sectionId}/pages`, { headers });
  const pages = await res.json();
  const existing = pages.value?.find((p) => p.title === siteName);
  if (existing) {
    console.log(`Replacing existing "${siteName}"...`);
    await fetch(`${base}/pages/${existing.id}`, { method: "DELETE", headers });
    // Wait for deletion to propagate (OneNote API is eventually consistent)
    await new Promise((r) => setTimeout(r, 2000));
  }

  // Build the multipart form body manually (Node's fetch doesn't handle mixed binary well)
  const boundary = "OneIntraNote" + Date.now();
  const pageHtml = `<!DOCTYPE html>\n<html><head><title>${siteName}</title></head><body>\n<p data-site-type="static" data-updated="${new Date().toISOString()}">${siteName}</p>\n<object data-attachment="site.zip" data="name:site.zip" type="application/zip" />\n</body></html>`;

  // Assemble multipart body: text parts (boundary + headers + HTML) + binary zip + closing boundary
  const textParts = `--${boundary}\r\nContent-Disposition: form-data; name="Presentation"\r\nContent-Type: text/html\r\n\r\n${pageHtml}\r\n--${boundary}\r\nContent-Disposition: form-data; name="site.zip"\r\nContent-Type: application/zip\r\n\r\n`;
  const closing = `\r\n--${boundary}--`;

  const encoder = new TextEncoder();
  const t = encoder.encode(textParts);
  const c = encoder.encode(closing);
  const body = new Uint8Array(t.length + zipBuffer.length + c.length);
  body.set(t, 0);
  body.set(zipBuffer, t.length);
  body.set(c, t.length + zipBuffer.length);

  console.log("Uploading...");
  const uploadRes = await fetch(`${base}/sections/${sectionId}/pages`, {
    method: "POST",
    headers: { ...headers, "Content-Type": `multipart/form-data; boundary=${boundary}` },
    body,
  });
  const data = await uploadRes.json();
  if (data.error) { console.error("Upload failed:", data.error.message); process.exit(1); }
  return data;
}

// --- CLI entry point ---
async function main() {
  // --- Argument parsing ---
  // Supports two modes:
  //   1. URL mode:  node publish.js <dir> <onenote-url> [site-name]
  //   2. Name mode: node publish.js <dir> <site-name> [notebook-name]
  const args = process.argv.slice(2);

  // Extract --site flag if present
  let graphSiteId = null;
  const filteredArgs = [];
  for (let i = 0; i < args.length; i++) {
    if (args[i] === "--site" && i + 1 < args.length) {
      graphSiteId = args[++i];
    } else {
      filteredArgs.push(args[i]);
    }
  }

  if (filteredArgs.length < 2) {
    console.log(`Usage: node publish.js <directory> <onenote-url> [site-name]
       node publish.js <directory> <site-name> [notebook-name]
       node publish.js <directory> <site-name> --site <sharepoint-site-id>

  Publish a static site to a OneNote notebook.

  You can pass a OneNote URL (copied from your browser) and the script
  will automatically find the right notebook:

    node publish.js ./dist "https://f5-my.sharepoint.com/...?sourcedoc=..."
    node publish.js ./dist "https://onedrive.live.com/...?sourcedoc=..."

  Or specify a site name and optional notebook name:

    node publish.js ./dist "My App"
    node publish.js ./dist "My App" "Work Notebook"

  For SharePoint site notebooks, use --site with the site ID:

    node publish.js ./dist "My App" --site "contoso.sharepoint.com,guid1,guid2"

Environment variables:
  ONEINTRANOTE_CLIENT_ID   Entra ID app client ID
  ONEINTRANOTE_TENANT_ID   Entra ID tenant ID (default: "common")`);
    process.exit(1);
  }

  const directory = filteredArgs[0];
  const secondArg = filteredArgs[1];
  const thirdArg = filteredArgs[2];

  // Validate the source directory exists and is actually a directory
  const dir = resolve(directory);
  if (!existsSync(dir) || !statSync(dir).isDirectory()) {
    console.error(`Error: "${dir}" is not a directory`);
    process.exit(1);
  }
  if (!existsSync(resolve(dir, "index.html"))) {
    console.warn(`Warning: No index.html found in "${dir}"`);
  }

  // Auto-detect personal Microsoft accounts from OneDrive URLs.
  // Personal accounts use the "consumers" endpoint instead of "common".
  const isUrl = secondArg.startsWith("http://") || secondArg.startsWith("https://") || secondArg.includes("sourcedoc=");
  if (isUrl && secondArg.includes("onedrive.live.com") && !process.env.ONEINTRANOTE_TENANT_ID) {
    TENANT_ID = "consumers";
    console.log("Detected personal Microsoft account URL — using consumers endpoint");
  }

  // Zip the directory using the system zip command
  const zipPath = `/tmp/oneintranote-${Date.now()}.zip`;
  console.log(`Zipping ${dir}`);
  execSync(`cd "${dir}" && zip -r "${zipPath}" .`, { stdio: "pipe" });
  const zipBuffer = readFileSync(zipPath);
  console.log(`Zip size: ${(zipBuffer.length / 1024).toFixed(0)} KB`);

  // Authenticate (uses cached token or opens browser)
  const token = await getAccessToken();
  console.log("Authenticated!\n");

  let notebookId, notebookName, siteName;

  // --- Determine target notebook ---
  if (graphSiteId) {
    // Site mode: find the first notebook on the SharePoint site (or use name match)
    siteName = secondArg;
    const base = oneNoteBase(graphSiteId);
    const headers = { Authorization: `Bearer ${token}` };
    const res = await fetch(`${base}/notebooks?$select=displayName,id`, { headers });
    const data = await res.json();
    if (data.error) { console.error("Failed to list site notebooks:", data.error.message); process.exit(1); }
    const nbs = data.value || [];
    if (nbs.length === 0) { console.error("No notebooks found on this SharePoint site."); process.exit(1); }
    // If third arg given, match by name; otherwise use first notebook
    const nb = thirdArg ? nbs.find(n => n.displayName === thirdArg) || nbs[0] : nbs[0];
    notebookId = nb.id;
    notebookName = nb.displayName;
    console.log(`Using site notebook: "${notebookName}" (${notebookId})`);
  } else if (secondArg.startsWith("http://") || secondArg.startsWith("https://") || secondArg.includes("sourcedoc=")) {
    // URL mode: extract notebook GUID from the pasted OneNote URL
    const parsed = parseOneNoteUrl(secondArg);

    if (!parsed.notebookGuid) {
      console.error("Could not extract notebook ID from URL. Make sure it contains a sourcedoc parameter.");
      process.exit(1);
    }

    console.log(`Extracted notebook GUID: ${parsed.notebookGuid}`);
    if (parsed.pageName) console.log(`Extracted page name: ${parsed.pageName}`);

    // Find the notebook in the user's account by matching the GUID
    const nb = await findNotebookByGuid(token, parsed.notebookGuid);
    if (!nb) {
      console.error(`No notebook found matching GUID ${parsed.notebookGuid}. Make sure you're signed into the right account.`);
      process.exit(1);
    }

    notebookId = nb.id;
    notebookName = nb.displayName;
    // Site name priority: CLI arg > page name from URL > default "Site"
    siteName = thirdArg || parsed.pageName || "Site";

    console.log(`Found notebook: "${notebookName}" (${notebookId})`);
  } else {
    // Name mode: second arg is site name, third arg is optional notebook name
    siteName = secondArg;
    const nbName = thirdArg || "Digital Garden";
    const nb = await findOrCreateNotebook(token, nbName);
    notebookId = nb.id;
    notebookName = nb.displayName;
  }

  // --- Upload ---
  console.log(`Publishing "${siteName}" to "${notebookName}"...\n`);

  const sectionId = await findOrCreateSection(token, notebookId, graphSiteId);
  const result = await uploadSite(token, sectionId, siteName, zipBuffer, graphSiteId);

  const sitePrefix = graphSiteId ? `/s/${encodeURIComponent(graphSiteId)}` : "";
  console.log(`\n  Published "${siteName}" to "${notebookName}"`);
  console.log(`  Page ID: ${result.id}`);
  console.log(`  View: ${VIEWER_URL}${sitePrefix}/nb/${encodeURIComponent(notebookId)}/${encodeURIComponent(siteName)}`);
}

main().catch((err) => { console.error(err); process.exit(1); });
