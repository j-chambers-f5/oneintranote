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

## Publishing

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

### What happens when you publish

1. The directory is zipped
2. You authenticate via browser (first time) or cached token (subsequent runs)
3. The script finds or creates the target notebook and a "Sites" section
4. The zip is uploaded as a file attachment on a new OneNote page
5. You get a shareable URL

### Updating a site

Just run the same publish command again. The old version is automatically replaced.

## Sharing with colleagues

1. **Publish** your site (see above)
2. **Share the OneNote notebook** — open OneNote, right-click the notebook, select "Share", and enter your colleague's email
3. **Send them the link** — the publish script outputs a URL like:
   ```
   https://j-chambers-f5.github.io/OneIntraNote/nb/<notebook-id>/<site-name>
   ```
4. **They open the shared notebook** in OneNote (one-time step from the email notification)
5. **They click your link** — sign in with Microsoft and the site loads

### Collaborative publishing

Multiple people can publish to the same shared notebook. If you share a notebook with edit access, your colleagues can also run `publish.js` to update sites in that notebook.

## Authentication

**First publish:** Opens your browser to sign in with Microsoft. The token is saved to `~/.oneintranote/token.json`.

**Subsequent publishes:** Token refreshes silently. No browser needed.

**Switch accounts:** Delete `~/.oneintranote/token.json` and run again.

**Identity:** The publish script shows which account you're signed in as:
```
Refreshing saved token for j.chambers@f5.com...
Authenticated!
```

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
The notebook must be in your account or shared with you. If it's shared, you need to open it in OneNote first (one-time step).

### Site shows old content after updating
The viewer uses stale-while-revalidate — it shows cached content instantly and updates in the background. Refresh the page to see the updated version.

### "Authentication timed out"
The browser auth window has a 2-minute timeout. Make sure you complete the sign-in promptly. If your default browser opened the wrong profile, change your default browser or sign out of the wrong account first.
