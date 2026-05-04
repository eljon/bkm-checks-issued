# BKM Checks Issued

A small web app for logging Philippine bank checks issued to suppliers. Form
data is written to a Google Sheet, with separate tabs for the running check log
and for searchable supplier and bank lists.

## Form fields

- Supplier (searchable, autocompletes from the `Suppliers` tab; new entries are added automatically)
- Check Date
- Check Number
- Bank (searchable, autocompletes from the `Banks` tab; new entries are added automatically)
- Amount (Philippine Peso)
- Notes (optional)

## Files

- `Code.gs` — Google Apps Script backend (sheet I/O, validation)
- `Index.html` — Form UI served by the Apps Script web app
- `appsscript.json` — Apps Script project manifest

## Setup

1. Create a new Google Sheet. This is where checks will be saved.
2. In that Sheet: **Extensions → Apps Script**.
3. In the Apps Script editor:
   - Replace the contents of `Code.gs` with the file from this repo.
   - Click the **+** next to "Files", choose **HTML**, name it `Index`, and
     paste the contents of `Index.html`.
   - Open project settings (gear icon) and tick **"Show appsscript.json
     manifest file"**, then replace its contents with `appsscript.json` from
     this repo.
4. Click **Deploy → New deployment → Web app**:
   - Execute as: **Me**
   - Who has access: **Only myself** (or **Anyone with Google account** if you
     want to share)
5. Authorize when prompted. Open the deployment URL — the `Checks`,
   `Suppliers`, and `Banks` tabs are created on first load.

When you push code changes, run **Deploy → Manage deployments → edit → New
version** to publish them.
