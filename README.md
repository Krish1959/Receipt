# Raj Govin Law Practice – Case Entry & Receipt Service

Three-step web form (Case Entry → **Receipt** → **Accounts**). The receipt comes before the
accounts page so the amount is typed and the receipt printed while the client is at the desk;
the office amounts are filled in afterwards, when the client has gone.

The sheet row is therefore written in **two passes**:

| Pass | When | What it writes |
|---|---|---|
| `POST /api/records` | **Submit** on the receipt page | the whole row + the Tally voucher. Accounts columns take their automatic values (NETT = AMOUNT, Govin/Raj 70/30) |
| `POST /api/records/accounts` | **Save Accounts** on the last page | columns **H–Q** of that same row (DIS/INC … TEL/IT); NETT and the Govin/Raj split are recalculated |

The receipt and the Tally voucher are never waiting on office data. If the accounts page is
never filled in, the row still carries the correct automatic split. Nothing is re-sent to Tally
on the second pass — the voucher is for AMOUNT, which the accounts page cannot change.

```
raj-govin-receipts/
├── server.js          Express API + Google Sheets writer
├── public/index.html  All three pages (all CSS/JS inline)
├── package.json
├── render.yaml        Render blueprint (optional)
└── .env.example
```

## Sheet columns
Ledger order (1–17): DATE · CLIENT · SUBJECT · O/R DAT · O/R · MATTER · AMOUNT · DIS/INC · NETT · SHARE · SALARY · GOVIN · RAJ · *(blank)* · RENT · *(blank)* · TEL/IT
Extras (18–20): LAWYER · MODE OF PAYMENT · TIMESTAMP (SGT)
Column 21 (U): TALLY – sync status

- MATTER = "Payment For" typed on the receipt page; one input serves both
- O/R = Bill No (`1-YYYYMMDD-HHMM/D` Dora Raj, `2-YYYYMMDD-HHMM/G` Govinda Raju), set on page 1 only
- DATE and O/R DAT = receipt date; NETT = AMOUNT − DIS/INC (auto)
- The accounts page matches the Bill No to find the row to update, so **the Bill No must stay unique**. If the same Bill No appears twice, the newest row is the one updated.
- Header row is written, bolded, frozen and RENT / TEL/IT highlighted yellow automatically on the first save.

---

## 1. Secrets needed

| Env var | What it is | Secret? |
|---|---|---|
| `GOOGLE_CREDENTIALS_JSON` | Full contents of the service-account JSON key | **Yes** |
| `GOOGLE_SHEET_ID` | ID of the Google Sheet in the `Receipts` folder | Yes (keep private) |
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
6. Do one test entry and confirm the row appears in the sheet, then press **Save Accounts** and confirm columns H–Q of that same row change.

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
- `/api/records/accounts` only ever writes columns H–Q of an existing row. It cannot change the receipt, the amount, the client or the Tally status.
- The site has no login. Anyone with the URL can add rows. If needed, add Render's basic-auth or a simple PIN — ask and it can be added.
