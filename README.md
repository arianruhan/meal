# Meal Chart Ledger — Live Dashboard

A live dashboard that reads your **original `Meal_Chart_*.xlsx` file, unchanged**
— same layout, same merged cells, same formulas. You keep filling it out
exactly like you do now. GitHub pulls it from OneDrive every 30 minutes and
republishes the dashboard.

```
Your original .xlsx  --(every 30 min, GitHub Action)-->  data.json  -->  GitHub Pages (index.html)
```

## How it reads your file

`scripts/build_data.py` doesn't expect a special format — it reads the sheet
the same way a person would look at it:

- Finds each person by their name in column A, then reads the three rows
  under them (Breakfast, Lunch, Dinner day-counts) and the Personal
  Total / Cost / Deposit / Balance / Khalar / Gas / Electricity cells next to
  their name — by column **header text**, not fixed positions, so it still
  works if you insert or reorder columns.
- Finds the "Bazar List" section and reads every trip row underneath its
  header, until it hits a few blank rows in a row.
- Finds "Meal Rate", "Total Deposit", "Current Balance", "Total due", etc. by
  searching for that exact label text anywhere in the sheet and reading the
  next filled cell to its right — so it's tolerant of the summary block
  shifting up/down if you add more people or bazar entries.

Because it reads the numbers Excel already calculated (your formulas), the
dashboard's totals will always match what you see when you open the file —
no separate math to keep in sync.

**If you ever restructure the sheet significantly** (rename "Bazar List",
remove the Meal Rate cell, etc.), the script will need a matching update —
it mirrors today's file, not a generic parser.

## 1. Put the file on OneDrive

1. Upload your existing `Meal_Chart_*.xlsx` to OneDrive — no edits needed.
2. Right-click it → **Share** → set link permission to **"Anyone with the
   link" → Can view** → **Copy link**.
3. Keep editing the file as you always have (Excel desktop, Excel Online, or
   the OneDrive mobile app all work) — just make sure it's saved before the
   next 30-minute cycle runs.

## 2. Create the GitHub repo

1. Create a new **public** repo (Pages on the free tier needs it public,
   unless you have GitHub Pro/Team).
2. Push everything in this folder to it (`index.html`, `data.json`,
   `scripts/`, `.github/workflows/`, this README).
3. **Settings → Secrets and variables → Actions → New repository secret** —
   add:
   - `ONEDRIVE_SHARE_URL` — the share link from step 1.2.
4. **Settings → Actions → General → Workflow permissions** → select
   **Read and write permissions**, then Save. (Needed so the Action can
   commit the refreshed `data.json` back to the repo.)
5. **Settings → Pages → Build and deployment → Source: Deploy from a branch**
   → Branch: `main`, folder `/ (root)` → Save.
6. Your dashboard will be live at
   `https://<your-username>.github.io/<repo-name>/` within a minute or two.

## 3. First run

**Actions** tab → **Update meal chart data** → **Run workflow**, to pull real
data immediately instead of waiting for the schedule. After that it runs
automatically every 30 minutes and only re-commits `data.json` (triggering a
fresh Pages deploy) when something actually changed.

## Notes

- The OneDrive direct-download trick
  (`api.onedrive.com/v1.0/shares/u!<encoded-link>/root/content`) is a
  long-standing, widely-used convenience API, but it's unofficial rather than
  a documented guarantee. If it ever starts failing, the fallback is
  Microsoft Graph with a proper app registration — more setup, but fully
  supported.
- If your OneDrive is a **SharePoint/OneDrive for Business** account,
  "Anyone with the link" sharing may be disabled by IT policy.
- GitHub's `cron` schedule is best-effort and **disables scheduled workflows
  automatically after 60 days with no repository activity** (any commit
  resets that timer).
- The page also polls `data.json` every 30 minutes in the browser and has a
  **Refresh now** button, so an open tab doesn't need a manual reload.
- If a fetch fails mid-deploy, the dashboard keeps showing the last data it
  loaded and shows a small error banner rather than going blank.
