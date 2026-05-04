# BKM Checks Issued

Static web app for logging Philippine bank checks. Form data is appended to a
Google Sheet via a tiny Apps Script "webhook" attached to the sheet. Supplier
and bank inputs autocomplete from their own tabs. A Print button lays the
values out on a 175mm × 75mm page in standard Philippine check positions.

**No Google Cloud Console.** The Apps Script is set up once and never touched
again. All UI changes live in this repo and ship with `git push`.

## Files

- `index.html` — the entire app
- `config.js` — your Apps Script Web App URL
- `apps_script.gs` — paste-once script for the Sheet
- `README.md` — this file

## One-time setup

### 1. Create the Google Sheet

1. Go to <https://sheets.google.com> → **Blank** to create a new spreadsheet.
2. Rename it (e.g., "BKM Checks Issued"). The script will create the
   `Checks`, `Suppliers`, and `Banks` tabs on first save.

### 2. Install the Apps Script webhook

1. In the Sheet: **Extensions → Apps Script**.
2. Replace the contents of the editor with everything in `apps_script.gs`
   from this repo. **File → Save**.
3. **Deploy → New deployment**:
   - Click the gear icon → **Web app**.
   - Description: `BKM Checks API`
   - **Execute as**: `Me`
   - **Who has access**: `Anyone`
   - **Deploy**.
4. Authorize when prompted (Advanced → Go to … (unsafe) → Allow).
5. Copy the **Web app URL** that ends in `/exec`.

### 3. Configure the app

Edit `config.js`:

```js
window.APP_CONFIG = {
  apiUrl: 'https://script.google.com/macros/s/AKfyc.../exec',
  secret: ''
};
```

Commit and push.

### 4. Host the page

Easiest: **GitHub Pages**.

1. Push the repo to GitHub.
2. Repo **Settings → Pages → Build from branch** → pick this branch,
   `/ (root)` → **Save**.
3. Wait ~1 minute. Open the URL Pages gives you (e.g.,
   `https://<user>.github.io/bkm-checks-issued/`).

Or run locally:

```sh
python3 -m http.server 8000
# open http://localhost:8000
```

## Updating the app

Edit any file in this repo → `git commit` → `git push`. Pages re-deploys
automatically. Hard-refresh (Ctrl/Cmd+Shift+R) to pick up changes. The Apps
Script doesn't need to be touched.

## Optional: shared secret

The Web App URL is unguessable, but if you want a second layer:

1. In `apps_script.gs`, set `const SHARED_SECRET = 'your-passphrase';` and
   redeploy (Deploy → Manage deployments → edit → New version).
2. In `config.js`, set `secret: 'your-passphrase'`.

Note this isn't a real secret — anyone who views the page source can read it.
It just stops casual abuse if the URL leaks.

## Adjusting the check print layout

The `:root` block at the top of `index.html` exposes all check positions as
CSS variables (`--check-w`, `--check-h`, `--check-date-top`,
`--check-payee-left`, etc.). Tweak them once to match your bank's pre-printed
check, then `git push`.
