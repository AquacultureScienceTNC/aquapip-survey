# AquaPIP — MEL Monitoring Survey

A shareable web survey that matches aquaculture practitioners to monitoring
protocols from the TNC MEL evidence base. A farmer answers a short survey
(what they want to measure + a light farm profile), and the app returns
**one best practitioner-runnable protocol per indicator**, ranked by usability
tier, with a downloadable PDF.

The app reads your Excel database **live**. Update the workbook, push it, and
both the survey options *and* the recommendations refresh automatically — no
code changes.

---

## Contents

```
aquapip-survey/
├── app.py                       ← the app (run this)
├── mel_logic.py                 ← data loading + matching (edit the goal map here)
├── report.py                    ← PDF report generator
├── storage.py                   ← optional Google-Sheets response saving
├── requirements.txt             ← Python packages
├── README.md                    ← this file
├── .gitignore
├── data/
│   └── mel_database.xlsx         ← YOUR database (keep this exact filename)
├── protocols/                    ← for V2 downloadable files (empty for now)
│   ├── templates/  datasheets/  workbooks/
└── .streamlit/
    ├── config.toml               ← colours / theme
    └── secrets.toml.template      ← copy to secrets.toml only if enabling storage
```

---

## The fastest path: deploy for free on Streamlit Community Cloud

You will end with a public link like `https://aquapip-mel.streamlit.app` that
anyone can open. No servers to manage. ~15–20 minutes the first time.

You need three free accounts, all sign-in-with-Google friendly:
**GitHub**, **Streamlit Community Cloud**, and (optional, later) **Google Cloud**.

### 1. Make a GitHub account + install GitHub Desktop
1. Go to <https://github.com/join> and create an account.
2. Download **GitHub Desktop** from <https://desktop.github.com>, install it,
   and sign in with your new account. (This lets you move files to GitHub by
   dragging — no command line.)

### 2. Create a repository and add these files
1. In GitHub Desktop: **File → New repository**.
   - Name: `aquapip-survey`
   - Local path: anywhere you like
   - Leave the rest default; click **Create repository**.
2. Open that new folder (GitHub Desktop shows a **Show in Finder/Explorer**
   button).
3. Drag **everything inside this bundle** into that folder — `app.py`,
   `mel_logic.py`, `report.py`, `storage.py`, `requirements.txt`, `README.md`,
   `.gitignore`, and the `data/`, `protocols/`, `.streamlit/` folders.
4. Back in GitHub Desktop you'll see all the files listed as changes. In the
   bottom-left box type a summary like `Initial AquaPIP survey`, then click
   **Commit to main**.
5. Click **Publish repository** (top bar). **Uncheck "Keep this code private"**
   so Streamlit's free tier can read it → **Publish repository**.

Your code is now on GitHub.

### 3. Deploy on Streamlit Community Cloud
1. Go to <https://share.streamlit.io> and **Sign in with GitHub** → authorise.
2. Click **Create app** (or **New app**) → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository:** `your-username/aquapip-survey`
   - **Branch:** `main`
   - **Main file path:** `app.py`
4. Click **Deploy**. It installs the packages and launches (first build takes a
   couple of minutes). When it's done you have your public URL — copy and share it.

That's the whole survey, live. **Skip to “Keeping it up to date”** unless you
want to save responses.

---

## Optional: turn on response storage (Google Sheets)

Responses are saved **only** when the practitioner ticks the consent box **and**
you've done this one-time setup. Until then the consent box simply has nothing
to write to (the privacy-safe default).

### A. Create the destination sheet
1. Make a new Google Sheet (e.g. "AquaPIP responses").
2. From its URL, copy the long id:
   `https://docs.google.com/spreadsheets/d/`**`THIS_LONG_ID`**`/edit`

### B. Create a service account (a robot Google login for the app)
1. Go to <https://console.cloud.google.com>, create/select a project.
2. **APIs & Services → Library** → enable **Google Sheets API** and
   **Google Drive API**.
3. **APIs & Services → Credentials → Create credentials → Service account**.
   Give it a name (e.g. `aquapip-writer`), create, done.
4. Open that service account → **Keys → Add key → Create new key → JSON**.
   A `.json` file downloads. Keep it safe.
5. Open the JSON, copy the `client_email` value
   (looks like `aquapip-writer@…iam.gserviceaccount.com`).
6. In your Google Sheet, click **Share** and share it with that
   `client_email` as **Editor**.

### C. Give the app the credentials
- **On Streamlit Cloud (recommended):** open your app → **⋮ → Settings →
  Secrets**. Paste the contents of `.streamlit/secrets.toml.template`, then
  replace the placeholders with your `sheet_id` and the fields from the JSON
  (copy `private_key` exactly, keeping the `\n` line breaks). Save — the app
  reruns automatically.
- **For local testing:** copy `.streamlit/secrets.toml.template` to
  `.streamlit/secrets.toml` and fill it in the same way. This file is
  git-ignored and must never be committed.

Each consenting submission then appends a row: timestamp, farm name, goals,
indicator codes, farm-profile JSON, and a consent marker.

---

## Keeping it up to date (the maintenance loop)

This is the part that makes the tool "live". **Whenever your database changes:**

1. Save the updated workbook **over** `data/mel_database.xlsx` — **keep that
   exact filename**. (Your file may be called `…v23.xlsx` on your machine, but
   in the repo it must stay `mel_database.xlsx` so the app finds it. Git keeps
   the version history for you, so you don't need versioned filenames here.)
2. In GitHub Desktop: type a commit message (e.g. `Database v23 + method
   summaries`) → **Commit to main** → **Push origin**.
3. Streamlit redeploys within a minute. Refresh the app.

**What updates automatically, with no code change:**
- **Survey options** — the indicator lists and every farm-profile dropdown are
  built from your workbook's `Vocabularies` sheet *and* the values actually used
  in the data. Add a species, a new indicator, a new farm structure → it appears.
- **Recommendations** — new rows, retagged rows, and changed tiers all flow
  straight through.
- **New columns appear by name.** The app reads columns by their **header text**,
  not their position, so:
  - Add **`Protocol Method Summary`** (your AK column) → it renders as the
    “Method summary” block for each protocol (until then it shows
    “method summary pending”).
  - Add **`Protocol Template URL`**, **`Data Sheet URL`**, **`Stats Workbook URL`**
    → the greyed-out “(V2)” download buttons turn into real download links
    (see V2 note below).
  Column letters don't matter — only that the header text matches.

> Keep the data on a sheet named **`in`** and the controlled lists on a sheet
> named **`Vocabularies`** (the app looks for these; it falls back to the first
> sheet if `in` is missing, but keeping the names is safest).

---

## Editing the farmer-facing goal screen

Step 1 of the survey (“What are you interested in measuring?”) is the only
hand-curated layer. It lives at the top of **`mel_logic.py`** in the
`FARMER_GOALS` dictionary. Each goal maps to a list of MEL indicator **codes**
(e.g. `"WQ 1.1.1"`). To add/rename a goal or change which indicators it
pre-selects, edit that dictionary and push. Any indicator **not** attached to a
goal still appears in the step-3 picker, so nothing is unreachable.

---

## V2: attaching downloadable files per protocol

The scaffolding is already in place:
- The empty `protocols/templates`, `protocols/datasheets`, `protocols/workbooks`
  folders are where per-protocol files will live.
- When you're ready, add the three URL columns named above to the workbook and
  put a link (to the committed file, or any hosted location) in the relevant
  cell for a protocol. The download buttons activate for exactly those rows.
No code change needed for V1 → V2; the buttons already read those columns.

---

## Run it on your own computer (optional)

Handy for testing before pushing.
```bash
pip install -r requirements.txt
streamlit run app.py
```
It opens at `http://localhost:8501`.

---

## Troubleshooting

- **“Database not found”** — `data/mel_database.xlsx` is missing or renamed.
  Restore that exact path/filename and push.
- **App won't build on Streamlit Cloud** — check the logs (“Manage app” panel).
  Usually a typo in `requirements.txt` or the repo is still private (make it
  public, or add Streamlit as a collaborator).
- **Responses aren't saving** — confirm the consent box was ticked, the sheet is
  shared with the service-account `client_email` as Editor, and both APIs
  (Sheets + Drive) are enabled. Storage never crashes the app; it just no-ops.
- **An indicator shows “No protocol in the current database yet”** — that's
  expected when no row is tagged with that indicator code (e.g. an indicator in
  the vocab that hasn't been populated yet).
- **A protocol looks mis-ranked** — ranking is: tier (T4→T1), then protocol over
  reference, then farm-profile fit, then lower cost, then lower effort. Adjust an
  entry's `Usability Rating`, `Estimated Cost`, `Sampling Effort`, or tags in the
  workbook to change its position.

---

## Note: the R Shiny alternative

We chose Streamlit because your pipeline is already Python. If you ever want R
instead, the equivalent host is **shinyapps.io** (free tier ~25 active hours per
month, which is fine for a low-traffic survey). It would need the logic
rewritten in R; the Streamlit version has no such monthly cap on Community
Cloud, so it's the better fit here.
