# NADAC vs. Cost Plus Drugs — Auto-Refreshing Pipeline

Pulls the current CMS NADAC dataset and the live Cost Plus Drugs catalog,
matches them by NDC, and rebuilds:

- `docs/nadac_vs_costplus.xlsx` — the full multi-tab analysis workbook
- `docs/index.html` — the interactive search/scatter explorer

...automatically, on a weekly schedule, with no computer of yours needing to be on.

## What you get at the end

A permanent URL (something like `https://yourusername.github.io/nadac-cpd-pipeline/`)
that always shows last week's freshly-pulled data — and an Excel file sitting in the
same repo that refreshes the same way.

## One-time setup (about 10 minutes)

### 1. Create the repo
- Go to [github.com/new](https://github.com/new)
- Name it anything (e.g. `nadac-cpd-pipeline`)
- Keep it **Public** (required for free GitHub Pages on a free account)
- Don't add a README/gitignore from GitHub's UI — we already have our own files

### 2. Upload these files
Drag-and-drop the whole folder into the repo via GitHub's web UI ("Add file" →
"Upload files"), keeping the folder structure:
```
your-repo/
├── pipeline.py
├── explorer_template.html
├── .github/
│   └── workflows/
│       └── update.yml
└── docs/              (this can start empty — the pipeline fills it in)
```
If GitHub's web uploader flattens your folders, instead clone the repo locally and
copy the files in with `git`, preserving the `.github/workflows/` path exactly —
GitHub Actions only recognizes workflows in that exact location.

### 3. Turn on GitHub Pages
- In your repo: **Settings → Pages**
- Under "Build and deployment", set **Source** to **GitHub Actions**
  (not "Deploy from a branch" — the workflow handles deployment itself)

### 4. Run it for the first time
- Go to the **Actions** tab in your repo
- Click **"Refresh NADAC vs Cost Plus Drugs data"** in the left sidebar
- Click **"Run workflow"** → **Run workflow** (this is the manual trigger;
  after this it also runs automatically every Thursday)
- Wait 1-2 minutes, then refresh the page — you should see a green checkmark

### 5. Find your link
- **Settings → Pages** will now show "Your site is live at https://..."
- That URL serves `docs/index.html` — bookmark it, share it, done

## Checking it worked
- The **Actions** tab shows every run's logs — if something fails, the error
  will be there (most likely cause: CMS changed the NADAC file's URL structure
  or column names — see "If NADAC's download breaks" below)
- Your repo's `docs/` folder will show the actual files it produced, with
  commit timestamps — that's your audit trail of every refresh

## Changing the schedule
Edit the `cron:` line in `.github/workflows/update.yml`. The default
(`0 13 * * 4`) runs every Thursday at 1pm UTC. [crontab.guru](https://crontab.guru)
will translate any cron expression into plain English if you want to change it.
You can also just click "Run workflow" manually anytime, regardless of schedule.

## If NADAC's download breaks
`pipeline.py` tries CMS's Socrata API first (a stable resource ID that shouldn't
need updating), and falls back to scraping the dated CSV link off Medicaid.gov's
NADAC landing page if that fails. Government data portals do occasionally change
structure without much notice. If a run fails:
1. Check the Actions log for which step failed
2. Visit https://www.medicaid.gov/medicaid/prescription-drugs/nadac-national-average-drug-acquisition-cost/index.html
   manually and confirm the CSV link pattern still matches what's in `download_nadac()`
3. The Cost Plus API (`api.costplusdrugs.com/pricelist/cpd`) has been stable
   throughout this project's development, but the same logic applies if it ever changes

## Important caveats (carried over from the original analysis)
- **NADAC** is pharmacy acquisition cost, not what a cash-pay patient is billed.
- **Est. Retail (AWP)** columns use published industry ratios (Brand ≈ NADAC×1.25,
  Generic ≈ NADAC×1.90) — an industry rule of thumb, not a measured retail price.
- **PA Tier** (Low/Moderate/High) is a clinical heuristic based on drug class,
  brand status, and cost — **not** measured prior-authorization or denial data.
- Cost Plus does not carry controlled substances or most cold-chain injectables;
  those will never appear in the matched dataset regardless of how matching is tuned.
- This is a methodology demonstration, not medical or financial advice.

## Running it locally instead
If you'd rather run this on your own machine on demand (no GitHub Actions needed):
```
pip install requests openpyxl
python3 pipeline.py
```
Output lands in the local `docs/` folder.
