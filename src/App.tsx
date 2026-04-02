/**
 * App.tsx — OneIntraNote React Shell
 *
 * This is the main entry point for the OneIntraNote viewer. The app:
 *   1. Authenticates the user via MSAL (Microsoft Entra ID)
 *   2. Lists notebooks that contain a "Sites" section
 *   3. When a site is selected, downloads the zip attachment from its OneNote page
 *   4. Unpacks the zip with JSZip and sends files to the Service Worker via MessageChannel
 *   5. Renders the site fullscreen in an iframe; the SW intercepts /site/* fetches
 *
 * Stale-while-revalidate flow:
 *   - On load, if the SW already has cached files for this site (same siteId), show them
 *     immediately in the iframe and revalidate from OneNote in the background.
 *   - If the SW cache is empty or for a different site, fetch fresh from OneNote first.
 *
 * URL routing scheme:
 *   /                           -> notebook list (home)
 *   /nb/<notebook-id>           -> sites list for that notebook
 *   /nb/<notebook-id>/<site>    -> load and display the site
 *   /nb/<notebook-id>/<site>/x  -> deep link into the site (x is passed to iframe)
 *
 * PWA manifest passthrough:
 *   - If the loaded site contains a manifest.json / manifest.webmanifest, we parse it,
 *     rewrite icon URLs to point through the SW, set start_url/scope to the current page,
 *     and tell the SW to serve it at a unique path so each site can install as its own PWA.
 *
 * SW communication protocol (via MessageChannel):
 *   - CACHE_FILES: send { files, prefix, siteId } -> SW stores files in memory, replies { count }
 *   - CHECK_CACHE: ask if SW has cached files -> replies { count, siteId }
 *   - SET_MANIFEST: send { manifest, path } -> SW will serve the manifest JSON at that path
 */
import { useState, useRef, useEffect } from "react";
import {
  MsalProvider,
  useMsal,
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from "@azure/msal-react";
import {
  PublicClientApplication,
  InteractionStatus,
} from "@azure/msal-browser";
import JSZip from "jszip";
import ReactMarkdown from "react-markdown";
import remarkGfm from "remark-gfm";
import { msalConfig, graphScopes, BASE_PATH } from "./services/authConfig";
import readmeContent from "../README.md?raw";

// --- MSAL initialization ---
// Create the MSAL instance at module scope so it's ready before any component renders.
// handleRedirectPromise() completes the redirect-based login flow if returning from Entra ID.
const msalInstance = new PublicClientApplication(msalConfig);
await msalInstance.initialize();
await msalInstance.handleRedirectPromise();

// SW_PREFIX is the URL path prefix under which the Service Worker serves site files.
// All site assets are served under <BASE_PATH>site/* so the SW can intercept them.
const SW_PREFIX = `${BASE_PATH}site/`;

// --- Service Worker communication helpers ---

/**
 * Send unpacked site files to the Service Worker for in-memory caching.
 * Uses MessageChannel so the SW can reply with the count of cached files.
 * The siteId (e.g. "notebookId/siteName") lets the SW track which site is cached
 * for the stale-while-revalidate check.
 */
async function sendFilesToSW(files: Record<string, ArrayBuffer>, siteId?: string): Promise<number> {
  const reg = await navigator.serviceWorker.ready;
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = (e) => resolve(e.data.count);
    reg.active!.postMessage({ type: "CACHE_FILES", files, prefix: SW_PREFIX, siteId }, [channel.port2]);
  });
}

/**
 * Check if the Service Worker already has files cached (and for which site).
 * Used on page load to decide whether to show cached content immediately
 * (stale-while-revalidate) or wait for a fresh download.
 * Times out after 500ms in case the SW isn't responding.
 */
async function checkSWCache(): Promise<{ count: number; siteId: string | null }> {
  const reg = await navigator.serviceWorker.ready;
  if (!reg.active) return { count: 0, siteId: null };
  return new Promise((resolve) => {
    const channel = new MessageChannel();
    channel.port1.onmessage = (e) => resolve({ count: e.data.count, siteId: e.data.siteId });
    setTimeout(() => resolve({ count: 0, siteId: null }), 500);
    reg.active!.postMessage({ type: "CHECK_CACHE" }, [channel.port2]);
  });
}

// --- Type definitions for Graph API responses ---
interface NotebookInfo { name: string; id: string; webUrl: string; hasSites?: boolean; siteId?: string }
interface SiteInfo { title: string; id: string; updated: string }

/** Returns the Graph API OneNote base URL — /me/onenote for personal, /sites/{id}/onenote for site notebooks. */
function oneNoteBase(graphSiteId?: string | null) {
  return graphSiteId
    ? `https://graph.microsoft.com/v1.0/sites/${graphSiteId}/onenote`
    : "https://graph.microsoft.com/v1.0/me/onenote";
}

/** Build an app URL, including /s/{siteId} prefix if it's a site notebook. */
function nbUrl(notebookId: string, graphSiteId?: string | null, suffix = "") {
  const sitePrefix = graphSiteId ? `/s/${encodeURIComponent(graphSiteId)}` : "";
  return appUrl(`${sitePrefix}/nb/${encodeURIComponent(notebookId)}${suffix}`);
}

// --- Inline styles (dark theme) ---
const darkBg: React.CSSProperties = { background: "#161618", minHeight: "100vh", color: "#e0e0e4", fontFamily: "sans-serif" };
const mutedText: React.CSSProperties = { color: "#8a8a92" };
const btnStyle: React.CSSProperties = { padding: "12px 32px", background: "#7b97aa", border: "none", borderRadius: 6, color: "#fff", cursor: "pointer", fontSize: 15 };
const cardStyle: React.CSSProperties = { background: "#1e1e21", border: "1px solid #2e2e32", borderRadius: 8, padding: "16px 20px", cursor: "pointer" };
const inputStyle: React.CSSProperties = { width: "100%", maxWidth: 600, padding: "10px 14px", background: "#1e1e21", border: "1px solid #2e2e32", borderRadius: 6, color: "#e0e0e4", fontFamily: "sans-serif", fontSize: 14, outline: "none" };

// --- URL routing helpers ---
// These translate between the browser URL bar and the app's internal route state.

// URL scheme: <BASE_PATH>nb/<notebook-id>/<site-name>/<path...>
function getAppPath() {
  const p = decodeURIComponent(window.location.pathname);
  const base = BASE_PATH.replace(/\/$/, "");
  return p.startsWith(base) ? p.slice(base.length) : p;
}

function appUrl(subpath: string) {
  return `${BASE_PATH.replace(/\/$/, "")}${subpath.startsWith("/") ? subpath : "/" + subpath}`;
}

/**
 * Parse the current URL into route components.
 * Returns { notebookId, siteName, sitePath } where sitePath is the sub-path
 * within the loaded site (passed to the iframe).
 */
function parseRoute() {
  const path = getAppPath();
  // /s/<siteId>/nb/<notebookId>/<siteName>/<path...>
  const siteMatch = path.match(/^\/s\/([^/]+)\/nb\/([^/]+)(?:\/([^/]+))?(\/.*)?$/);
  if (siteMatch) {
    return {
      graphSiteId: decodeURIComponent(siteMatch[1]),
      notebookId: decodeURIComponent(siteMatch[2]),
      siteName: siteMatch[3] ? decodeURIComponent(siteMatch[3]) : null,
      sitePath: siteMatch[4] || "/",
    };
  }
  // /nb/<notebookId>/<siteName>/<path...>
  const match = path.match(/^\/nb\/([^/]+)(?:\/([^/]+))?(\/.*)?$/);
  if (match) {
    return {
      graphSiteId: null as string | null,
      notebookId: decodeURIComponent(match[1]),
      siteName: match[2] ? decodeURIComponent(match[2]) : null,
      sitePath: match[3] || "/",
    };
  }
  return { graphSiteId: null as string | null, notebookId: null as string | null, siteName: null as string | null, sitePath: "/" };
}

/**
 * Try to match a pasted OneNote URL against the user's notebook list.
 * Supports direct webUrl match, resid/cid parameter extraction (OneDrive URLs),
 * and fuzzy matching by notebook name in the URL.
 */
function findNotebookByUrl(nbs: NotebookInfo[], pastedUrl: string): NotebookInfo | null {
  const lower = pastedUrl.toLowerCase();
  // Direct webUrl match
  for (const nb of nbs) {
    if (nb.webUrl && decodeURIComponent(nb.webUrl).toLowerCase() === decodeURIComponent(lower).replace(/\/$/, "")) return nb;
  }
  // Extract resid/cid from onedrive.live.com URLs and match
  const resid = pastedUrl.match(/resid=([^&!]+)/i)?.[1]?.toLowerCase();
  const cid = pastedUrl.match(/cid=([^&]+)/i)?.[1]?.toLowerCase();
  if (resid || cid) {
    for (const nb of nbs) {
      const nbUrl = nb.webUrl.toLowerCase();
      if (resid && nbUrl.includes(resid)) return nb;
      if (cid && nbUrl.includes(cid)) return nb;
    }
  }
  // Fuzzy: match by notebook name in URL
  for (const nb of nbs) {
    if (lower.includes(encodeURIComponent(nb.name).toLowerCase()) || lower.includes(nb.name.toLowerCase())) return nb;
  }
  return null;
}

// --- Main authenticated component ---
// This component handles all post-login functionality: listing notebooks/sites,
// loading site zips, communicating with the SW, and rendering the iframe.
function Main() {
  const { instance, accounts } = useMsal();
  const [status, setStatus] = useState("");
  const [notebooks, setNotebooks] = useState<NotebookInfo[]>([]);
  const [currentNotebook, setCurrentNotebook] = useState<NotebookInfo | null>(null);
  const [sites, setSites] = useState<SiteInfo[]>([]);
  const [ready, setReady] = useState(false);
  const [urlInput, setUrlInput] = useState("");
  const iframeRef = useRef<HTMLIFrameElement>(null);
  const iframePathRef = useRef(SW_PREFIX);
  const routeRef = useRef(parseRoute());

  /** Acquire an access token silently (from cache or via refresh). */
  async function getToken() {
    const account = accounts[0];
    if (!account) throw new Error("Not authenticated");
    return (await instance.acquireTokenSilent({ scopes: graphScopes.onenote, account })).accessToken;
  }

  /** Fetch all notebooks from Graph API, following pagination links. */
  async function fetchAllNotebooks(token: string): Promise<NotebookInfo[]> {
    const all: NotebookInfo[] = [];
    let url: string | null = "https://graph.microsoft.com/v1.0/me/onenote/notebooks?$select=displayName,id,links&$top=50";
    while (url) {
      const res: Response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      const data: { value?: Array<{ displayName: string; id: string; links?: { oneNoteWebUrl?: { href?: string } } }>; "@odata.nextLink"?: string } = await res.json();
      for (const n of data.value || []) {
        all.push({ name: n.displayName, id: n.id, webUrl: n.links?.oneNoteWebUrl?.href || "" });
      }
      url = data["@odata.nextLink"] || null;
    }
    return all;
  }

  // --- Initialization on mount ---
  useEffect(() => { init(); }, []);

  /**
   * Main initialization logic. Determines what to show based on the URL route:
   *   - If route has notebookId + siteName: load the site (with stale-while-revalidate)
   *   - If route has notebookId only: show the sites list for that notebook
   *   - If no route: show notebooks that have a "Sites" section
   */
  async function init() {
    const route = routeRef.current;
    try {
      // --- Stale-while-revalidate check ---
      // If we're loading a specific site, check if the SW already has it cached.
      // If yes, show the cached version immediately and revalidate from OneNote in the background.
      if (route.notebookId && route.siteName) {
        await navigator.serviceWorker.register(`${BASE_PATH}sw.js`, { scope: BASE_PATH });
        await navigator.serviceWorker.ready;
        const currentSiteId = `${route.notebookId}/${route.siteName}`;
        const cached = await checkSWCache();
        if (cached.count > 0 && cached.siteId === currentSiteId) {
          // Serve from cache immediately
          iframePathRef.current = `${SW_PREFIX.replace(/\/$/, "")}${route.sitePath}`;
          setReady(true);
          // Extract title from cached index.html
          try {
            const indexRes = await fetch(`${SW_PREFIX}index.html`);
            if (indexRes.ok) {
              const html = await indexRes.text();
              const titleMatch = html.match(/<title>([^<]+)<\/title>/i);
              if (titleMatch) document.title = titleMatch[1];
            }
          } catch { /* ignore */ }
          // Revalidate in background (fetches latest from OneNote, updates SW cache)
          revalidateInBackground(route.notebookId, route.siteName, route.graphSiteId);
          return;
        }
      }

      // --- Fresh load (no cache hit) ---
      const token = await getToken();
      setStatus("Loading...");
      const allNbs = await fetchAllNotebooks(token);

      if (route.notebookId) {
        // Try to find the notebook by ID in the user's list
        let nb = allNbs.find((n) => n.id === route.notebookId);

        // If not found in list, try a direct Graph API lookup
        if (!nb) {
          try {
            const directRes = await fetch(
              `${oneNoteBase(route.graphSiteId)}/notebooks/${route.notebookId}?$select=displayName,id`,
              { headers: { Authorization: `Bearer ${token}` } }
            );
            if (directRes.ok) {
              const directNb = await directRes.json();
              nb = { name: directNb.displayName, id: directNb.id, webUrl: "", siteId: route.graphSiteId || undefined };
            }
          } catch { /* ignore */ }
        }

        if (!nb) { setStatus(`Notebook not found: ${route.notebookId}`); return; }
        if (route.graphSiteId) nb = { ...nb, siteId: route.graphSiteId };
        setCurrentNotebook(nb);
        if (route.siteName) {
          iframePathRef.current = `${SW_PREFIX.replace(/\/$/, "")}${route.sitePath}`;
          await loadSite(token, nb.id, route.siteName, route.graphSiteId);
        } else {
          await loadSitesList(token, nb.id, route.graphSiteId);
        }
      } else {
        // Home page — show all notebooks, then check which have a "Sites" section
        setStatus("Loading notebooks...");
        setNotebooks(allNbs);
        setStatus("");
        // Check all notebooks for Sites sections in parallel (fast)
        const sitesMap = new Map<string, boolean>();
        const checks = allNbs.map(async (nb) => {
          try {
            const sectionId = await findSitesSection(token, nb.id);
            if (sectionId) {
              sitesMap.set(nb.id, true);
              // Re-render immediately when a site-enabled notebook is found
              const updated = allNbs.map(n => ({ ...n, hasSites: sitesMap.get(n.id) || false }));
              updated.sort((a, b) => (b.hasSites ? 1 : 0) - (a.hasSites ? 1 : 0));
              setNotebooks(updated);
            }
          } catch { /* ignore */ }
        });
        await Promise.all(checks);
      }
    } catch (err: unknown) {
      setStatus(`Error: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  /**
   * Background revalidation: fetches the latest zip from OneNote and updates
   * the SW cache. The iframe will pick up changes on next navigation within the site.
   * Errors are silently ignored — the user already has a working cached version.
   */
  async function revalidateInBackground(notebookId: string, siteName: string, graphSiteId?: string | null) {
    try {
      const token = await getToken();
      const base = oneNoteBase(graphSiteId);

      const sectionId = await findSitesSection(token, notebookId, graphSiteId);
      if (!sectionId) return;

      const headers = { Authorization: `Bearer ${token}` };
      const res = await fetch(`${base}/sections/${sectionId}/pages`, { headers });
      const pages = await res.json();
      const page = pages.value?.find((p: { title: string }) => p.title === siteName);
      if (!page) return;

      // Fetch the page HTML content to find the zip attachment URL
      const res2 = await fetch(`${base}/pages/${page.id}/content`, { headers });
      const html = await res2.text();
      const doc = new DOMParser().parseFromString(html, "text/html");
      let zipUrl = doc.querySelector('object[data-attachment$=".zip"]')?.getAttribute("data") || null;
      if (zipUrl) zipUrl = zipUrl.replace("/siteCollections/", "/sites/");
      if (!zipUrl) return;

      // Download and unpack the zip
      const res3 = await fetch(zipUrl, { headers });
      if (!res3.ok) return;

      const zip = await JSZip.loadAsync(await res3.arrayBuffer());
      const files: Record<string, ArrayBuffer> = {};
      for (const [name, entry] of Object.entries(zip.files)) {
        if (!entry.dir) files[name] = await entry.async("arraybuffer");
      }

      // Update SW cache — the iframe will pick up changes on next navigation
      await sendFilesToSW(files, `${notebookId}/${siteName}`);

      // Update title/manifest from new content
      if (files["index.html"]) {
        try {
          const htmlText = new TextDecoder().decode(files["index.html"]);
          const titleMatch = htmlText.match(/<title>([^<]+)<\/title>/i);
          if (titleMatch) document.title = titleMatch[1];
        } catch { /* ignore */ }
      }
    } catch { /* background revalidation failed silently */ }
  }

  /**
   * Find the "Sites" section in a notebook. Each notebook that hosts sites
   * must have a section named "Sites" — this is the convention used by publish.js.
   */
  async function findSitesSection(token: string, notebookId: string, graphSiteId?: string | null) {
    const res = await fetch(`${oneNoteBase(graphSiteId)}/notebooks/${notebookId}/sections`, { headers: { Authorization: `Bearer ${token}` } });
    const data = await res.json();
    return data.value?.find((s: { displayName: string }) => s.displayName === "Sites")?.id || null;
  }

  /** Load the list of sites (pages) in a notebook's "Sites" section. */
  async function loadSitesList(token: string, notebookId: string, graphSiteId?: string | null) {
    setStatus("Loading sites...");
    const sectionId = await findSitesSection(token, notebookId, graphSiteId);
    if (!sectionId) { setStatus("No 'Sites' section found in this notebook. Publish a site first with the CLI."); return; }
    const res = await fetch(`${oneNoteBase(graphSiteId)}/sections/${sectionId}/pages?$orderby=title`, { headers: { Authorization: `Bearer ${token}` } });
    const data = await res.json();
    setSites((data.value || []).map((p: { title: string; id: string; lastModifiedDateTime: string }) => ({
      title: p.title, id: p.id, updated: new Date(p.lastModifiedDateTime).toLocaleDateString(),
    })));
    setStatus("");
  }

  /**
   * Load a site: register SW, find the page in OneNote, download the zip attachment,
   * unpack it, send files to the SW, handle PWA manifest passthrough, and show the iframe.
   *
   * The zip attachment is found by parsing the OneNote page HTML for an <object> tag
   * with data-attachment ending in ".zip". The "data" attribute is a Graph API URL
   * that returns the raw zip bytes.
   */
  async function loadSite(token: string, notebookId: string, siteName: string, graphSiteId?: string | null) {
    const headers = { Authorization: `Bearer ${token}` };
    const base = oneNoteBase(graphSiteId);
    setStatus("Registering service worker...");
    await navigator.serviceWorker.register(`${BASE_PATH}sw.js`, { scope: BASE_PATH });
    await navigator.serviceWorker.ready;

    setStatus("Finding site...");
    const sectionId = await findSitesSection(token, notebookId, graphSiteId);
    if (!sectionId) { setStatus("No 'Sites' section in this notebook"); return; }

    const res = await fetch(`${base}/sections/${sectionId}/pages`, { headers });
    const pages = await res.json();
    const page = pages.value?.find((p: { title: string }) => p.title === siteName);
    if (!page) { setStatus(`Site "${siteName}" not found`); return; }

    // Fetch the page's HTML content to extract the zip attachment URL
    setStatus("Getting attachment...");
    const res2 = await fetch(`${base}/pages/${page.id}/content`, { headers });
    const html = await res2.text();
    const doc = new DOMParser().parseFromString(html, "text/html");
    let zipUrl = doc.querySelector('object[data-attachment$=".zip"]')?.getAttribute("data") || null;
    // OneNote API embeds "siteCollections" in resource URLs for site notebooks,
    // but Graph API only recognizes "sites". Fix the path.
    if (zipUrl) zipUrl = zipUrl.replace("/siteCollections/", "/sites/");
    if (!zipUrl) { setStatus("No zip attachment found"); return; }

    // Download and unpack the zip with JSZip
    setStatus("Downloading...");
    const res3 = await fetch(zipUrl, { headers });
    if (!res3.ok) {
      const errBody = await res3.text();
      console.error("[OneIntraNote] download error:", res3.status, errBody.substring(0, 500));
      setStatus(`Download failed: ${res3.status}`);
      return;
    }

    setStatus("Unpacking...");
    const zip = await JSZip.loadAsync(await res3.arrayBuffer());
    const files: Record<string, ArrayBuffer> = {};
    for (const [name, entry] of Object.entries(zip.files)) {
      if (!entry.dir) files[name] = await entry.async("arraybuffer");
    }

    // Extract title from the site's index.html to set as the document title
    if (files["index.html"]) {
      try {
        const htmlText = new TextDecoder().decode(files["index.html"]);
        const titleMatch = htmlText.match(/<title>([^<]+)<\/title>/i);
        if (titleMatch) document.title = titleMatch[1];
      } catch { /* ignore */ }
    }

    // --- PWA manifest passthrough ---
    // If the loaded site includes a PWA manifest, we extract it, rewrite icon URLs
    // to go through the SW (so they resolve from the in-memory cache), set unique
    // start_url/scope per site (so each site can be installed as a separate PWA),
    // and tell the SW to serve the manifest at a unique path.
    const manifestNames = ["manifest.json", "manifest.webmanifest", "site.webmanifest"];
    for (const mName of manifestNames) {
      if (files[mName]) {
        try {
          const manifestText = new TextDecoder().decode(files[mName]);
          const manifest = JSON.parse(manifestText);
          // Rewrite icon URLs to point through the service worker
          if (manifest.icons) {
            manifest.icons = manifest.icons.map((icon: { src: string }) => ({
              ...icon,
              src: `${SW_PREFIX}${icon.src.replace(/^\.?\//, "")}`,
            }));
          }
          // Set start_url and scope to current page path for unique PWA identity
          manifest.start_url = window.location.pathname;
          manifest.scope = window.location.pathname;
          // Update document title
          if (manifest.name) document.title = manifest.name;
          // Tell the SW to serve this manifest at a path unique to this app
          const manifestPath = `${window.location.pathname.replace(/\/$/, "")}/manifest.webmanifest`;
          const reg = await navigator.serviceWorker.ready;
          reg.active?.postMessage({ type: "SET_MANIFEST", manifest, path: manifestPath });
          // Point the <link rel="manifest"> in index.html to the SW-served manifest
          const manifestLink = document.getElementById("pwa-manifest");
          if (manifestLink) manifestLink.setAttribute("href", manifestPath);
        } catch { /* ignore malformed manifests */ }
        break;
      }
    }

    // Send all unpacked files to the SW and show the iframe
    await sendFilesToSW(files, `${notebookId}/${siteName}`);
    setReady(true);
    setStatus("");
  }

  // --- URL bar sync ---
  // The iframe loads content from /site/*, but the browser URL bar should show
  // /nb/<id>/<site>/<subpath>. We poll the iframe's location every 300ms and
  // update the parent URL via replaceState. We use polling (setInterval) because
  // cross-origin restrictions prevent listening to iframe navigation events.
  useEffect(() => {
    if (!ready || !currentNotebook || !routeRef.current.siteName) return;
    const interval = setInterval(() => {
      try {
        const iframe = iframeRef.current;
        if (!iframe?.contentWindow) return;
        const p = iframe.contentWindow.location.pathname;
        const prefix = SW_PREFIX.replace(/\/$/, "");
        if (p.startsWith(prefix)) {
          const innerPath = p.slice(prefix.length).replace(/\/$/, "");
          const newPath = nbUrl(currentNotebook.id, currentNotebook.siteId, `/${encodeURIComponent(routeRef.current.siteName!)}${innerPath}`);
          if (window.location.pathname !== newPath) window.history.replaceState(null, "", newPath);
        }
      } catch { /* ignore */ }
    }, 300);
    return () => clearInterval(interval);
  }, [ready, currentNotebook]);

  /** Handle the "Go" button on the URL input — find a notebook matching the pasted URL. */
  function handleGo() {
    const url = urlInput.trim();
    if (!url) return;
    const nb = findNotebookByUrl(notebooks, url);
    if (nb) {
      window.location.href = appUrl(`/nb/${encodeURIComponent(nb.id)}`);
    } else {
      setStatus(`No matching notebook found. Try selecting one from the list below.`);
    }
  }

  // --- Render ---

  // When a site is loaded, render it fullscreen in an iframe.
  // The iframe src points to /site/... which the SW intercepts.
  if (ready) {
    return <iframe ref={iframeRef} src={iframePathRef.current} style={{ position: "fixed", top: 0, left: 0, width: "100vw", height: "100vh", border: "none" }} title="Site" />;
  }

  // Loading / error state
  if (status) {
    return (
      <div style={{ ...darkBg, padding: 40, textAlign: "center" }}>
        <p style={mutedText}>{status}</p>
        {status.includes("No matching") && notebooks.length > 0 && (
          <div style={{ marginTop: 24, textAlign: "left", maxWidth: 600, margin: "24px auto" }}>
            {notebooks.map((nb) => (
              <a key={nb.id} href={appUrl(`/nb/${encodeURIComponent(nb.id)}`)} style={{ ...cardStyle, textDecoration: "none", color: "inherit", display: "block", marginBottom: 12 }}>
                <div style={{ fontSize: 16, fontWeight: 600 }}>{nb.name}</div>
              </a>
            ))}
          </div>
        )}
      </div>
    );
  }

  // Sites list view for a selected notebook
  if (currentNotebook && sites.length > 0) {
    return (
      <div style={{ ...darkBg, padding: 40 }}>
        <p style={{ ...mutedText, marginBottom: 8 }}><a href={appUrl("/")} style={{ color: "#7b97aa", textDecoration: "none" }}>Home</a></p>
        <h1 style={{ marginBottom: 8 }}>{currentNotebook.name}</h1>
        <p style={{ ...mutedText, marginBottom: 24 }}>Sites in this notebook</p>
        <div style={{ display: "grid", gap: 12, maxWidth: 600 }}>
          {sites.map((s) => (
            <a key={s.id} href={nbUrl(currentNotebook.id, currentNotebook.siteId, `/${encodeURIComponent(s.title)}`)} style={{ ...cardStyle, textDecoration: "none", color: "inherit", display: "block" }}>
              <div style={{ fontSize: 16, fontWeight: 600 }}>{s.title}</div>
              <div style={{ fontSize: 12, ...mutedText, marginTop: 4 }}>Updated {s.updated}</div>
            </a>
          ))}
        </div>
      </div>
    );
  }

  // Home page — notebook list with URL paste input
  return (
    <div style={{ ...darkBg, padding: 40 }}>
      <h1 style={{ marginBottom: 8 }}>OneIntraNote</h1>
      <p style={{ ...mutedText, marginBottom: 24 }}>Logged in as {accounts[0]?.username}</p>
      <p style={{ marginBottom: 12 }}>Paste your OneNote notebook URL:</p>
      <div style={{ display: "flex", gap: 8, maxWidth: 700 }}>
        <input type="text" value={urlInput} onChange={(e) => setUrlInput(e.target.value)} onKeyDown={(e) => e.key === "Enter" && handleGo()} placeholder="Paste any OneNote notebook URL" style={inputStyle} />
        <button onClick={handleGo} style={btnStyle}>Go</button>
      </div>
      {notebooks.length > 0 && (
        <>
          <p style={{ ...mutedText, margin: "32px 0 12px" }}>Your notebooks:</p>
          <div style={{ display: "grid", gap: 12, maxWidth: 600 }}>
            {notebooks.map((nb) => (
              <a key={nb.id} href={appUrl(`/nb/${encodeURIComponent(nb.id)}`)} style={{ ...cardStyle, textDecoration: "none", color: "inherit", display: "block", borderColor: nb.hasSites ? "#7b97aa" : "#2e2e32" }}>
                <div style={{ fontSize: 16, fontWeight: 600 }}>
                  {nb.name}
                  {nb.hasSites && <span style={{ fontSize: 11, color: "#7b97aa", marginLeft: 8 }}>has sites</span>}
                </div>
                <div style={{ fontSize: 11, ...mutedText, marginTop: 4, wordBreak: "break-all" }}>{nb.webUrl}</div>
              </a>
            ))}
          </div>
        </>
      )}

    </div>
  );
}

// --- Login/landing page (shown when unauthenticated) ---
// Renders the README.md content below the sign-in button so visitors can
// understand what OneIntraNote is before authenticating.
function LoginButton() {
  const { instance, inProgress } = useMsal();
  if (inProgress !== InteractionStatus.None) return <div style={{ ...darkBg, padding: 40 }}><p style={mutedText}>Authenticating...</p></div>;
  return (
    <div style={{ ...darkBg, padding: "40px 40px 80px" }}>
      <div style={{ maxWidth: 800, margin: "0 auto", textAlign: "center" }}>
        <h1 style={{ fontSize: 36, marginBottom: 8 }}>OneIntraNote</h1>
        <p style={{ ...mutedText, margin: "8px 0 20px" }}>Host static sites from your OneNote notebooks</p>
        <button onClick={() => instance.loginRedirect({ scopes: graphScopes.onenote, prompt: "consent" })} style={btnStyle}>Sign in with Microsoft</button>
      </div>
      <div style={{
        maxWidth: 800, margin: "48px auto 0", lineHeight: 1.7, fontSize: 15,
        color: "#c9d1d9",
      }}>
        <ReactMarkdown
          remarkPlugins={[remarkGfm]}
          components={{
            img: ({ src, alt }) => {
              // README images are at public/docs/ in the repo, served at /docs/ by Vite
              const imgSrc = src?.replace(/^public\//, "");
              return <img src={imgSrc?.startsWith("docs/") ? `${BASE_PATH}${imgSrc}` : src} alt={alt || ""} style={{ maxWidth: "100%", borderRadius: 8, margin: "16px 0" }} />;
            },
            a: ({ href, children }) => (
              <a href={href} target="_blank" rel="noopener noreferrer" style={{ color: "#7b97aa" }}>{children}</a>
            ),
            code: ({ children }) => (
              <code style={{ background: "#1e1e21", padding: "2px 6px", borderRadius: 4, fontSize: 13 }}>{children}</code>
            ),
            pre: ({ children }) => (
              <pre style={{ background: "#1e1e21", padding: 16, borderRadius: 8, overflow: "auto", fontSize: 13, margin: "16px 0" }}>{children}</pre>
            ),
            h1: () => null, // Skip the first H1 since we already have the title above
            h2: ({ children }) => <h2 style={{ borderBottom: "1px solid #2e2e32", paddingBottom: 8, marginTop: 32 }}>{children}</h2>,
          }}
        >
          {readmeContent.replace(/^# OneIntraNote\n/, "")}
        </ReactMarkdown>
      </div>
    </div>
  );
}

// --- App root ---
// MsalProvider wraps the app to provide authentication context.
// AuthenticatedTemplate / UnauthenticatedTemplate conditionally render
// based on whether the user is signed in.
export default function App() {
  return (
    <MsalProvider instance={msalInstance}>
      <AuthenticatedTemplate><Main /></AuthenticatedTemplate>
      <UnauthenticatedTemplate><LoginButton /></UnauthenticatedTemplate>
    </MsalProvider>
  );
}
