# Reference your live master Google Sheet (read-only) for de-duping

The daily sweep can read your master sheet **"All Closures + Openings"** every
morning to avoid re-surfacing schools (by name **and** by source link) that are
already on the list. The sheet stays **fully private** — the sweep reads it with a
Google *service account* you share it with as **Viewer**. Nothing ever edits it,
and the key lives only in a GitHub secret (never in the website or the browser).

## What you set up once (~10 min)

### 1. Make a service account + key
1. Go to <https://console.cloud.google.com/> → create (or pick) a project.
2. **APIs & Services → Library →** search **"Google Sheets API" → Enable**.
3. **APIs & Services → Credentials → Create credentials → Service account.**
   Give it a name (e.g. `school-tracker-reader`), click **Done**.
4. Open that service account → **Keys → Add key → Create new key → JSON**.
   A `.json` file downloads. Open it in a text editor — you'll paste the whole
   contents in step 3. It contains a `client_email` like
   `school-tracker-reader@your-project.iam.gserviceaccount.com`.

### 2. Share the sheet with it (Viewer)
In the sheet: **Share** → paste that `client_email` → role **Viewer** → Send.
(That's what lets the private sheet be read without making it public.)

### 3. Add two GitHub secrets
In the repo: **Settings → Secrets and variables → Actions → New repository secret.**
Add both (you paste these yourself — I never see them):

| Secret name        | Value                                                        |
|--------------------|-------------------------------------------------------------|
| `GOOGLE_SA_JSON`   | the **entire contents** of the downloaded `.json` key file  |
| `MASTER_SHEET_ID`  | `1q1PM_9goSwlmXy9pZEFUZ8IBVeZ_5LxMQ_qSxiqJdP4`              |

*(The sheet id is the long code between `/d/` and `/edit` in the sheet's URL.)*

### 4. Run it
**Actions → Daily school-tracker sweep → Run workflow.** The log will show
`Google Sheet roster: N names + M source links …`. From then on:
- The sweep skips any candidate whose **name** or **source link** is already on
  the sheet, and writes `data/roster.json`.
- The dashboard reads `data/roster.json` and quietly hides queue candidates
  already on the master sheet (with a "show them" toggle), and warns you if a
  link you paste into **Add by link** is already tracked.

## Notes
- **Read-only, by design.** The scope requested is `spreadsheets.readonly`; the
  service account can't change your sheet even if asked to.
- **Leave the secrets unset** and everything still works — the sweep just
  de-dupes against the workbook snapshot (`All_Closures___Openings.xlsx`) as before.
- The sweep reads **every tab** in the sheet, using the column whose header
  contains "name"/"school" for names and any URLs found in the row for links.
