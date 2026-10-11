# Raj Govin Law Practice – Case Entry & Receipt Service

Two-step web form (Case Entry → Receipt). On final Submit, both forms are saved as **one row** in a Google Sheet, and the receipt freezes into an A4 print layout.

```
raj-govin-receipts/
├── server.js          Express API + Google Sheets writer
├── public/index.html  Both forms (all CSS/JS inline)
├── package.json
├── render.yaml        Render blueprint (optional)
└── .env.example
```

## Sheet columns
Timestamp (SGT) · Lawyer · S/L No · Case Category · Item 1 · Item 2 · Remarks · Receipt Date · Bill No · Received From · Amount (S$) · Amount in Words · Payment For · Mode of Payment

The header row is written automatically on the first save.

---

## 1. Secrets needed

| Env var | What it is | Secret? |
|---|---|---|
| `GOOGLE_CREDENTIALS_JSON` | Full contents of the service-account JSON key | **Yes** |
| `GOOGLE_SHEET_ID` | ID of the Google Sheet in the `Raj&Raju` folder | Yes (keep private) |
| `GOOGLE_SHEET_TAB` | Tab name, default `Records` | No |

No Gmail password is used. The app signs in as a Google *service account* that you share the sheet with.

## 2. Google setup (logged in as agentic.avai@gmail.com)

1. **Create the folder & sheet**
   - Google Drive → New → Folder → `Raj&Raju`
   - Inside it: New → Google Sheets → name it e.g. `Raj&Raju Records`
   - Copy the ID from the URL: `docs.google.com/spreadsheets/d/`**`<SHEET_ID>`**`/edit` → this is `GOOGLE_SHEET_ID`

2. **Create a Google Cloud project** – <https://console.cloud.google.com> → project picker → New Project → `raj-raju-receipts`

3. **Enable the API** – APIs & Services → Library → *Google Sheets API* → Enable

4. **Create the service account** – APIs & Services → Credentials → Create Credentials → Service account → name `raj-raju-sheets` → Done (no roles needed)

5. **Create the key** – open the service account → Keys → Add key → Create new key → **JSON** → a `.json` file downloads. Its whole content = `GOOGLE_CREDENTIALS_JSON`. Store it safely; never commit it to GitHub.

6. **Share the sheet with the service account** – open the sheet → Share → paste the `client_email` from the JSON (`raj-raju-sheets@raj-raju-receipts.iam.gserviceaccount.com`) → **Editor** → untick "Notify" → Share

> Why create the sheet yourself? Service accounts have no Drive storage of their own, so they can't create files in a personal Gmail Drive. Sharing an existing sheet avoids that — the file stays in your 5 TB Drive, owned by agentic.avai@gmail.com.

## 3. Deploy on Render

1. Push this folder to a GitHub repo (the `.gitignore` keeps keys out).
2. Render → your existing `receipt-p122` service → **Settings**:
   - Change type: if it is a *Static Site*, create a new **Web Service** instead (static sites can't run Node). Point it at the repo.
   - Runtime **Node**, Build `npm install`, Start `npm start`, Region **Singapore**.
3. **Environment** tab → Add environment variables:
   - `GOOGLE_CREDENTIALS_JSON` → paste the entire JSON file content (multi-line is fine)
   - `GOOGLE_SHEET_ID` → the sheet ID
   - `GOOGLE_SHEET_TAB` → `Records`
4. Save → Render redeploys.
5. Check `https://<your-app>.onrender.com/healthz` → should show `"sheets":"connected"`. `"dry-run"` means a variable is missing.
6. Do one test entry and confirm the row appears in the sheet.

*Alternative:* use **New → Blueprint** with `render.yaml`; Render will prompt for the two secret values.

**Free plan note:** the free instance sleeps after ~15 min idle; the first visit then takes ~30–50 s to wake. Upgrade to *Starter* if that matters for front-desk use.

## 4. Run locally

```bash
npm install
npm start          # no env vars → DRY RUN: rows are printed to console, not saved
```
Open <http://localhost:3000>.

## Security notes
- Values are written with `RAW` input, so text starting with `=` is never executed as a formula.
- Server re-validates every field; request size capped at 20 KB.
- The site has no login. Anyone with the URL can add rows. If needed, add Render's basic-auth or a simple PIN — ask and it can be added.
