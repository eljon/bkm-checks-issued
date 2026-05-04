# BKM Checks Issued

Static web app for logging Philippine bank checks. Sign in with Google,
fill out the form, and the entry is appended directly to a Google Sheet
via the Sheets API. Supplier and bank inputs autocomplete from their own
tabs in the same sheet. A Print button lays the values out on a 175mm ×
75mm page in standard Philippine check positions.

No backend server, no Apps Script — push the files to any static host
(GitHub Pages, Netlify, Vercel, your own server) and that's the whole
deploy. Re-deploy = `git push`.

## Files

- `index.html` — the entire app
- `config.js` — your Sheet ID and OAuth Client ID
- `README.md` — this file

## One-time setup

### 1. Create the Google Sheet

Create a new Google Sheet. Copy the ID out of its URL:

```
https://docs.google.com/spreadsheets/d/<THIS_PART_IS_THE_ID>/edit
```

The app will create the `Checks`, `Suppliers`, and `Banks` tabs on first
use.

### 2. Create an OAuth Client ID

1. Go to <https://console.cloud.google.com/>, create (or pick) a project.
2. **APIs & Services → Library** → enable **Google Sheets API**.
3. **APIs & Services → OAuth consent screen** → set up an external app.
   Add yourself as a test user. Add the scope
   `https://www.googleapis.com/auth/spreadsheets`.
4. **APIs & Services → Credentials → Create credentials → OAuth client ID**:
   - Application type: **Web application**
   - **Authorized JavaScript origins**: every origin you'll serve the app
     from. Examples:
     - `http://localhost:8000` for local testing
     - `https://<your-username>.github.io` for GitHub Pages
     - `https://your-domain.com` for a custom host
5. Copy the generated **Client ID**.

### 3. Configure the app

Edit `config.js`:

```js
window.APP_CONFIG = {
  spreadsheetId: 'your-sheet-id',
  oauthClientId: 'your-client-id.apps.googleusercontent.com'
};
```

Commit and push.

## Running it

### Local

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

(The Google sign-in flow needs an `http://` or `https://` origin —
opening `index.html` via `file://` won't work.)

### GitHub Pages

Repo **Settings → Pages → Build from branch** → pick this branch,
`/ (root)` folder. Wait for the deployment, then visit the published URL.
Make sure that URL is added as an Authorized JavaScript origin in step 2.

## How it works

- Sign-in uses **Google Identity Services** (token client) — you grant
  the app permission to access spreadsheets, the browser receives a
  short-lived access token, and the token is cached in `sessionStorage`
  for the tab's lifetime.
- All reads/writes go straight to
  `https://sheets.googleapis.com/v4/spreadsheets/<id>/...`. There is no
  intermediate server.
- On first save, the three tabs are created if missing and headers are
  written.

## Adjusting the check print layout

The `:root` block at the top of `index.html` exposes all check positions
as CSS variables (`--check-w`, `--check-h`, `--check-date-top`,
`--check-payee-left`, etc.). Tweak them once to match your bank's
pre-printed check, then `git push`.
