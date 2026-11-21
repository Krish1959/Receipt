# Setting up the Tally Sync on the client's PC

For whoever installs it, and for the person who uses it afterwards.

---

## There is no `requirements.txt`, and no Python

This app is **Node.js**, not Python. `pip` plays no part.

| Python | here |
|---|---|
| `requirements.txt` | **`package.json`** — already in the folder |
| `pip install -r requirements.txt` | **`npm install`** — run once |
| a venv | nothing. Node installs globally; there is no environment to activate |

The conda `AI` environment on the development PC is irrelevant. `npm` comes with Node, from
`C:\Program Files\nodejs`, and works in any Command Prompt.

---

## One-time setup — about ten minutes

**1. Install Node.js.** The **LTS** build from `https://nodejs.org`. Accept the defaults. Nothing
else needs installing.

**2. Copy the folder** to the client PC. Anywhere the user can write — `C:\RajGovin\Tally\` is
fine, or their Documents. The batch files work out their own location, so the path does not matter
and nothing inside them needs editing.

**3. Create the settings file.** Copy `secrets-template.txt` into the `secrets\` folder, rename it
to **`tally-secrets.txt`**, and fill in the three `GOOGLE_` values from the Render dashboard.

> **No quotes around the values.** A quote opened and not closed travels into the credential, and
> Google rejects it with an error that says nothing about a stray `"`. The app warns at startup if
> it sees one.

This file holds a live Google refresh token. It never goes to GitHub, and never into a chat.

**4. Double-click `install-once.bat`.** It checks Node, installs the libraries, and tells you if
the settings file is still missing. This needs the internet; it happens once per computer.

---

## The desktop icon

1. Right-click **`start-tally-sync.bat`** → **Show more options** → **Send to** → **Desktop
   (create shortcut)**
2. On the desktop, rename the shortcut to **Tally Sync**
3. Right-click it → **Properties** → **Change Icon…** → **Browse** to
   `%SystemRoot%\System32\SHELL32.dll` and pick one that stands out

Do **not** tick "Run as administrator". It does not need it, and the prompt only confuses people.

---

## Using it, day to day

1. Open **TallyPrime** and load the company. The sync cannot post to a Tally that is not open.
2. Double-click **Tally Sync**. A black window opens and the browser goes to
   `http://localhost:3000`.
3. **Check New Items**, read the list, then **Update to TALLY**.
4. When finished, close the black window.

**Leave the black window open while using the sync.** It *is* the app. Closing it stops it; that
is also how you stop it. Nothing is left running afterwards.

Double-clicking the icon twice is harmless — the second one says it is already running and closes.

---

## What the user sees when something is wrong

Every one of these stops **before** anything reaches Tally, and says what to do:

| On screen | Means |
|---|---|
| `Node.js is not installed` | step 1 was skipped |
| `The settings file is missing` | step 3 was skipped |
| `Tally not answering` (red pill) | TallyPrime is closed, or no company is loaded |
| `TallyPrime has no ledger named "X"` | the ledger does not exist, or the name differs by a character |
| `Dated 14 Nov 2026 … Educational mode` | no licence; vouchers only on the 1st, 2nd or last of a month |

Nothing is ever posted twice: each row is marked `DONE` in the sheet the moment Tally confirms it.

---

## Two things to settle before go-live

- **Buy the TallyPrime licence.** Under Educational mode a receipt dated the 14th simply cannot be
  posted. This is not a bug to work around.
- **Create the ledgers the practice needs**, spelled exactly as in `secrets\tally-secrets.txt`.
  As at 22 Nov that is `Legal_Fees`, `Bank Account`, `Cash`, `Disbursements Recovered`,
  `Office Expenses`, and the group `Sundry Debtors`.
