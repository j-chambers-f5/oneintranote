/**
 * sw.js — Service Worker for OneIntraNote
 *
 * This Service Worker serves static site files from an in-memory cache.
 * The React app unpacks a zip from OneNote, sends the files here via
 * MessageChannel, and then loads the site in an iframe pointing to /site/*.
 *
 * Message protocol (all messages use MessageChannel for replies):
 *
 *   CACHE_FILES — Store files in memory for serving.
 *     Input:  { type: "CACHE_FILES", files: Record<string, ArrayBuffer>, prefix: string, siteId: string }
 *     Reply:  { type: "CACHED", count: number }
 *     The prefix (e.g. "/site/") determines which fetch requests are intercepted.
 *     The siteId (e.g. "notebookId/siteName") is stored for stale-while-revalidate checks.
 *
 *   CHECK_CACHE — Check if files are currently cached (used for stale-while-revalidate).
 *     Input:  { type: "CHECK_CACHE" }
 *     Reply:  { type: "CACHE_STATUS", count: number, siteId: string|null }
 *
 *   SET_MANIFEST — Store a PWA manifest to serve at a specific path.
 *     Input:  { type: "SET_MANIFEST", manifest: object, path: string }
 *     No reply. The manifest is served as application/manifest+json at the given path.
 *     Each site gets a unique manifest path so multiple sites can install as separate PWAs.
 *
 * Fetch interception:
 *   - Requests to <sitePrefix>* are served from the in-memory file cache
 *   - The manifest path (if set) is served with application/manifest+json content type
 *   - All other requests pass through to the network normally
 */

// In-memory file cache: Map<filePath, ArrayBuffer>
// This is populated by CACHE_FILES messages from the React app.
// Files are keyed by their path within the zip (e.g. "index.html", "assets/style.css").
let fileCache = new Map();

// The URL prefix for serving site files (e.g. "/site/").
// Set by the CACHE_FILES message to match the React app's SW_PREFIX.
let sitePrefix = "/site/";

// Skip waiting and claim clients immediately so the SW activates on first visit.
// This is important because the React app needs the SW ready before loading the iframe.
self.addEventListener("install", () => self.skipWaiting());
self.addEventListener("activate", (event) => event.waitUntil(self.clients.claim()));

// PWA manifest storage — allows the loaded site's manifest to be served at a unique URL.
let pwaManifest = null;
let pwaManifestPath = null;

// Tracks which site is currently cached, for stale-while-revalidate.
// Format: "notebookId/siteName"
let cachedSiteId = null;

// --- Message handler ---
self.addEventListener("message", (event) => {
  if (event.data.type === "CACHE_FILES") {
    // Replace the entire file cache with new files from the React app
    fileCache = new Map(Object.entries(event.data.files));
    if (event.data.prefix) sitePrefix = event.data.prefix;
    if (event.data.siteId) cachedSiteId = event.data.siteId;
    // Reply with the number of cached files via the MessageChannel port
    if (event.ports[0]) event.ports[0].postMessage({ type: "CACHED", count: fileCache.size });
  }
  if (event.data.type === "CHECK_CACHE") {
    // Report current cache state (used by the React app on page load to decide
    // whether to show cached content immediately or wait for a fresh download)
    if (event.ports[0]) event.ports[0].postMessage({ type: "CACHE_STATUS", count: fileCache.size, siteId: cachedSiteId });
  }
  if (event.data.type === "SET_MANIFEST") {
    // Store the PWA manifest for this site so it can be served at a unique URL.
    // The React app rewrites icon paths and sets start_url/scope before sending.
    pwaManifest = event.data.manifest;
    pwaManifestPath = event.data.path || null;
  }
});

/**
 * MIME type lookup by file extension.
 * Covers the common types needed for static sites (HTML, CSS, JS, images, fonts).
 * Falls back to application/octet-stream for unknown types.
 */
function getMimeType(path) {
  const ext = path.split(".").pop().toLowerCase();
  return {
    html: "text/html", css: "text/css", js: "application/javascript",
    json: "application/json", svg: "image/svg+xml", png: "image/png",
    jpg: "image/jpeg", jpeg: "image/jpeg", webp: "image/webp",
    gif: "image/gif", ico: "image/x-icon", woff: "font/woff",
    woff2: "font/woff2", ttf: "font/ttf", xml: "application/xml",
    txt: "text/plain", map: "application/json",
  }[ext] || "application/octet-stream";
}

/**
 * Resolve a file path against the cache, trying multiple fallbacks:
 *   1. Exact match (e.g. "about.html")
 *   2. Without leading slash (e.g. "/about.html" -> "about.html")
 *   3. Append .html (e.g. "about" -> "about.html") for clean URLs
 *   4. Append /index.html (e.g. "docs/" -> "docs/index.html") for directory indexes
 */
function resolve(filePath) {
  for (const c of [filePath, filePath.replace(/^\//, ""), filePath + ".html", filePath + "/index.html"]) {
    if (fileCache.has(c)) return c;
  }
  return null;
}

// --- Fetch interception ---
// Intercepts requests under the site prefix and serves them from the in-memory cache.
// Also serves the PWA manifest at its unique path if configured.
self.addEventListener("fetch", (event) => {
  const url = new URL(event.request.url);

  // Serve PWA manifest at the app-specific path (e.g. /nb/<id>/<site>/manifest.webmanifest)
  if (pwaManifest && pwaManifestPath && url.pathname === pwaManifestPath) {
    event.respondWith(new Response(JSON.stringify(pwaManifest), {
      status: 200,
      headers: { "Content-Type": "application/manifest+json" },
    }));
    return;
  }

  // Only intercept requests under the site prefix (e.g. /site/*)
  if (!url.pathname.startsWith(sitePrefix)) return;

  // Strip the prefix to get the file path within the zip
  // e.g. "/site/assets/style.css" -> "assets/style.css"
  // Default to "index.html" if the path is empty (root request)
  let filePath = url.pathname.slice(sitePrefix.length).replace(/^\//, "") || "index.html";

  event.respondWith(
    (async () => {
      const resolved = resolve(filePath);
      if (resolved) {
        return new Response(fileCache.get(resolved), {
          status: 200,
          headers: { "Content-Type": getMimeType(resolved), "Cache-Control": "no-cache" },
        });
      }
      return new Response("Not found: " + filePath, { status: 404 });
    })()
  );
});
