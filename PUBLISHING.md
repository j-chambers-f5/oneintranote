# Publishing Sites with OneIntraNote

This guide covers how to publish static sites to OneNote and share them with your team.

## Quick start

```bash
git clone https://github.com/j-chambers-f5/OneIntraNote.git
cd OneIntraNote
npm install
node publish.js ./my-site/dist "My Site"
```

The first run opens your browser to sign in with Microsoft. After that, the token is cached and subsequent publishes are instant.

## Publishing to a personal notebook

```bash
node publish.js <directory> <site-name> [notebook-name]
```

### Examples

```bash
# Publish a Vite/React app
npm run build
node publish.js ./dist "My Dashboard"

# Publish a Quartz/Obsidian digital garden
npx quartz build
node publish.js ./public "Digital Garden"

# Publish to a specific notebook
node publish.js ./dist "My App" "Work Notebook"

# Publish using a OneNote URL (auto-detects notebook)
node publish.js ./dist "https://onedrive.live.com/...?sourcedoc=..."
```

## Publishing to a SharePoint site notebook

For team-wide sharing, publish to a SharePoint site notebook. Everyone who is a member of the SharePoint site can access the published sites — no individual sharing required.

```bash
node publish.js <directory> <site-name> --site <sharepoint-site-id>
```

### Finding your SharePoint site ID

You can find the site ID using the Graph Explorer or the Microsoft Graph API:

```
GET https://graph.microsoft.com/v1.0/sites/{hostname}:/sites/{site-name}?$select=id
```

For example:
```
GET https://graph.microsoft.com/v1.0/sites/contoso.sharepoint.com:/sites/MyTeamSite?$select=id
```

The `id` field in the response is your site ID (format: `hostname,guid1,guid2`).

### Examples

```bash
# Publish to a SharePoint site notebook
node publish.js ./dist "Team Dashboard" --site "contoso.sharepoint.com,abc123,def456"

# Publish to a specific notebook on the site
node publish.js ./dist "My App" --site "contoso.sharepoint.com,abc123,def456" "Project Notebook"
```

## Publishing manually (no CLI required)

If you don't want to install Node.js or use the command line, you can publish directly through the OneNote app. This works for any notebook — personal or SharePoint site notebooks.

### Step 1: Zip your site

Your static site must have an `index.html` at the root. Zip the entire site directory into a single file called `site.zip`.

**macOS:**
1. Open Finder and navigate to your site's build output folder (e.g. `dist/` or `public/`)
2. Select **all files inside the folder** (not the folder itself) — Cmd+A
3. Right-click → "Compress Items"
4. Rename the resulting `Archive.zip` to `site.zip`

**Windows:**
1. Open File Explorer and navigate to your site's build output folder
2. Select all files inside the folder — Ctrl+A
3. Right-click → "Compress to ZIP file"
4. Rename the file to `site.zip`

**Important:** The zip must contain `index.html` at the top level, not inside a subfolder. If you open the zip, you should see `index.html` immediately — not a folder containing it.

### Step 2: Create the "Sites" section in OneNote

1. Open the target notebook in OneNote (desktop app or OneNote for the web)
2. Create a new **section** called exactly `Sites` (case-sensitive)

### Step 3: Create a page and attach the zip

1. Inside the "Sites" section, create a new **page**
2. Set the page **title** to whatever you want the site to be called (e.g. "Team Dashboard")
3. Attach `site.zip` to the page:
   - **OneNote desktop:** Go to Insert → File Attachment → select `site.zip` → choose "Attach File"
   - **OneNote for the web:** Drag and drop `site.zip` onto the page, or use Insert → File

That's it. The site is now published.

### Step 4: Get the shareable link

You need two pieces of information to construct the link:

- **Notebook ID:** Open the notebook in OneNote for the web. The URL will contain the notebook ID — look for a segment like `1-xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx` in the address bar.
- **Site name:** The exact page title you used in Step 3.

**For personal notebooks:**
```
https://j-chambers-f5.github.io/oneintranote/nb/<notebook-id>/<site-name>
```

**For SharePoint site notebooks:**
```
https://j-chambers-f5.github.io/oneintranote/s/<site-id>/nb/<notebook-id>/<site-name>
```

Replace spaces in the site name with `%20` (e.g. `Team%20Dashboard`).

### Updating a manually published site

1. Delete the existing page in the "Sites" section
2. Create a new page with the same title
3. Attach the updated `site.zip`

The viewer caches the previous version, so returning visitors will see the old site instantly while the new version loads in the background. A page refresh will show the update.

## Sharing

### Personal notebooks

1. **Publish** your site (see above)
2. **Share the OneNote notebook** — open OneNote, right-click the notebook, select "Share", and enter your colleague's email
3. **Send them the link** — the publish script outputs a URL like:
   ```
   https://j-chambers-f5.github.io/oneintranote/nb/<notebook-id>/<site-name>
   ```
4. **They open the shared notebook** in OneNote (one-time step from the email notification)
5. **They click your link** — sign in with Microsoft and the site loads

### SharePoint site notebooks

1. **Publish** to a site notebook with `--site` (see above)
2. **Send colleagues the link** — the publish script outputs a URL like:
   ```
   https://j-chambers-f5.github.io/oneintranote/s/<site-id>/nb/<notebook-id>/<site-name>
   ```
3. **They click the link** — anyone who is a member of the SharePoint site can view the site. No extra sharing step needed.

This is the recommended approach for teams, since SharePoint site membership is already managed by your IT admin.

### Collaborative publishing

Multiple people can publish to the same shared notebook. If you share a notebook with edit access (or use a SharePoint site notebook where members have contribute permissions), your colleagues can also run `publish.js` to update sites in that notebook.

## Authentication

**First publish:** Opens your browser to sign in with Microsoft. The token is saved to `~/.oneintranote/token.json`.

**Subsequent publishes:** Token refreshes silently. No browser needed.

**Switch accounts:** Delete `~/.oneintranote/token.json` and run again.

## PWA support

If your site includes a `manifest.json` (or `manifest.webmanifest`), it becomes installable as a standalone app. Each site gets its own PWA identity — multiple sites can be installed simultaneously with their own name, icon, and theme color.

### Requirements for PWA install

- Site must include a `manifest.json` with `name`, `icons` (192x192 and 512x512), and `display: "standalone"`
- Icons must be included in the published directory

## Self-hosting

If you want to deploy your own OneIntraNote instance:

1. **Set up the Entra ID app:**
   ```bash
   node setup.js [your-deployed-url]
   ```
   This creates an app registration via Azure CLI and writes a `.env` file.

2. **Build and deploy:**
   ```bash
   npm run build
   ```
   Deploy the `dist/` folder to any static host (GitHub Pages, Netlify, etc).

3. **Update `publish.js`:** Change the `VIEWER_URL` constant to your deployed URL.

## Development

```bash
npm install
npm run dev    # start dev server at http://localhost:5173
npm run build  # production build
```

## Troubleshooting

### "No 'Sites' section found in this notebook"
The publish script creates a "Sites" section automatically when you first publish. If you see this error in the viewer, make sure you've published at least one site to the notebook.

### "Notebook not found"
The notebook must be in your account or shared with you. If it's shared, you need to open it in OneNote first (one-time step). For SharePoint site notebooks, make sure you're a member of the site.

### Site shows old content after updating
The viewer uses stale-while-revalidate — it shows cached content instantly and updates in the background. Refresh the page to see the updated version.

### "Authentication timed out"
The browser auth window has a 2-minute timeout. Make sure you complete the sign-in promptly. If your default browser opened the wrong profile, change your default browser or sign out of the wrong account first.

### "Download failed: 400"
If you see this for a SharePoint site notebook, the OneNote API may be returning a legacy URL format. This is handled automatically in the latest version — make sure you're on the latest code.
