# LinkHarvest (Windows)

A desktop app that scans an Excel sheet for links (images, videos, or audio)
and downloads them in parallel, with a modern UI and a live progress bar.

Originally a VBA macro (`Download_All_Row_Images`), rebuilt in Python with:
- A modern GUI (dark theme, `customtkinter`) with sidebar navigation
- **Link Downloader page**: scans an Excel sheet for links, downloads
  images/videos/audio in parallel, custom row/column ranges, live
  progress bar
- **PDF Extractor page**: describe what to extract in plain English
  (e.g. "invoice number, vendor, total"), point it at a folder of PDFs,
  and it uses OpenAI or Gemini (your choice) to pull that data into an
  Excel sheet — accurate, never guesses, multithreaded
- **Data Profiler page**: load a CSV/Excel file, get a live health
  score and donut-chart breakdown of missing values, duplicates, blank
  rows, extra spaces, and special characters, then export a cleaned copy
- Adjustable thread counts and timeouts, with warnings if you move away
  from the tested defaults
- Automatic update checks against GitHub Releases on startup

---

## 1. Quick Start (running from source)

**Requirements:** Windows 10/11, Python 3.10+ installed and added to PATH.

```cmd
git clone https://github.com/<your-username>/linkharvest.git
cd linkharvest
pip install -r requirements.txt
python app.py
```

That's it — the app window will open.

## 2. Using the app

1. **Excel file** — browse to your `.xlsx`/`.xlsm` workbook.
2. **Sheet name** — defaults to `Sheet4` (matches the original macro);
   change it if your data is on a different sheet.
3. **Save folder** — where downloaded files should go.
4. **Rows / Columns** — leave blank for "from row 2 to the last row" and
   "all columns." To scan a single row or column, put the same number in
   both "From" and "To" (e.g. Rows From=5 To=5 scans only row 5).
5. **Advanced** — Threads (default `10`) and Timeout in seconds
   (default `30`). These defaults are the ones from the original script
   and are the most reliable; if you change them, the app will warn you
   that some links may fail or remain un-downloaded.
6. Click **Start Download** and watch the progress bar.

Column A of each row is always used as the file's ID/prefix, e.g.
`1023_image1.jpg`, `1023_image2.jpg`, matching the original macro's
naming convention. The file extension is guessed from the URL, so
video/audio links are saved with the right extension (`.mp4`, `.mp3`,
etc.) instead of always `.jpg`.

## 3. Using the PDF Extractor page

Click **📄 PDF Extractor** in the left sidebar. This page pulls
structured data out of a folder of PDFs using AI, and compiles the
results into an Excel sheet.

1. **AI Engine** — choose **Gemini (Flash tier, free)** or **OpenAI (gpt-5)**.
   Each provider has its own key slot, so switching back and forth
   doesn't overwrite the other's saved key.
2. **Model** — pick from the dropdown or type any model name directly
   (e.g. `gemini-flash-lite-latest` for lower cost/faster, `gpt-5-mini`
   for a cheaper OpenAI option). Both providers rename/retire models
   fairly often, so this field is intentionally free-editable rather
   than locked to a fixed list — check
   [ai.google.dev/gemini-api/docs/models](https://ai.google.dev/gemini-api/docs/models)
   or [platform.openai.com/docs/models](https://platform.openai.com/docs/models)
   for the current lineup if a model stops working.
3. **API Key** — paste the key for whichever engine you selected
   (OpenAI: platform.openai.com/api-keys · Gemini: aistudio.google.com/apikey)
   and click **Save**. It's written to a local `.env` file in the app
   folder and is never sent anywhere except directly to that provider's
   API — `.env` is git-ignored so it won't accidentally get committed.
4. **What to extract** — describe the fields in plain English, e.g.
   `Invoice number, invoice date, vendor name, total amount, payment due date`.
5. **Source Folder** — the folder containing your `.pdf` files.
6. **Output Excel / Sheet Name** — where results get written. If the
   file and sheet already exist, new rows are appended underneath the
   existing data rather than overwriting it.
7. **Concurrent Threads** (1–20, default 10) — how many PDFs to process
   at once. Higher isn't always faster — both providers rate-limit
   requests, so very high thread counts can trigger more retries.
8. Click **Start Extraction**. The log shows real-time per-file status,
   including which fields weren't found in a given document (the model
   is instructed to never guess — a blank cell means it genuinely
   wasn't in the PDF).
9. On completion, a success popup appears and the output folder opens
   automatically.

**Costs money to run:** each PDF sent to OpenAI or Gemini consumes API
credits on your own account with that provider (billed by them, not by
this app). Check your usage dashboard on whichever platform you're using.

**Handles gracefully:** password-protected/encrypted PDFs, corrupted
files, scanned/image-only PDFs with no extractable text, API timeouts
(auto-retries with backoff), and invalid API keys (stops immediately
with a clear message instead of silently failing on every file).

## 4. Using the Data Profiler page

Click **🧹 Data Profiler** in the sidebar. This page loads a CSV/Excel
file, detects common data-quality problems, and lets you export a
cleaned copy.

1. **Data file** — browse to a `.csv`, `.xlsx`, or `.xls` file.
2. Click **Analyze**. This populates:
   - **KPI cards** — Health Score (% of cells that are clean), Total
     Records, Total Columns, Total Anomalies.
   - **Anomaly checklist** — Missing Values, Duplicate Rows, Blank
     Rows, Extra Spaces, Special Characters, and Mixed Data Types,
     each showing how many cases were found. Only checked boxes get
     fixed on export.
   - **Donut chart** — a live visual breakdown; unchecking a box
     removes it from the chart immediately.
3. **Mixed Data Types is never auto-fixed** — it's shown for visibility
   (e.g. a numeric column with a stray text value in it) but is
   ambiguous enough that it's left for you to review manually rather
   than guessing what the "correct" fix is.
4. Click **Clean & Export**, choose where to save (`.xlsx` or `.csv`),
   and the cleaned file is written and the containing folder opens
   automatically.

## 5. Support / donation

The app has no paywall or access restrictions — it's free to use.
There's an optional "Support Development" panel that shows a UPI QR
code so people can leave a voluntary donation if they want to. The
UPI ID lives only in `donation.py`, isn't shown as text anywhere in
the app, and the QR uses a generic "Support Development" label instead
of a real name.

**Please read — important limitation:** UPI is built so that whoever
scans the code sees the bank-verified account holder's real name in
their own payment app before they pay — that's a fraud-prevention
feature of UPI itself, not something this (or any) app can hide or
turn off. This app only controls what's shown inside its own window;
it can't change what the payer's UPI app displays.

## 6. Distributing to other people (no Python required)

This is the "deploy it so anyone can use it" part. LinkHarvest ships
with a GitHub Actions workflow (`.github/workflows/build.yml`) that
automatically builds a standalone `LinkHarvest.exe` and attaches it to
a GitHub Release — nobody downloading it needs Python installed.

**To publish a release:**

```cmd
git tag v1.0.0
git push origin v1.0.0
```

Pushing a tag starting with `v` triggers the workflow. Watch it run
under the repo's **Actions** tab (takes a couple of minutes). When it
finishes, go to **Releases** on your repo — `LinkHarvest.exe` will be
attached and downloadable by anyone with the link, no account needed.

Bump the version and repeat (`git tag v1.0.1` etc.) whenever you push
updates you want to re-release.

**Building locally instead** (e.g. to test before tagging):

```cmd
build_exe.bat
```

This produces `dist\LinkHarvest.exe` with the custom icon baked in.

## 7. Project structure

```
LinkHarvest/
├── app.py                      # GUI (entry point, sidebar + all pages)
├── downloader.py                # link scanning + downloading logic
├── pdf_extractor.py               # PDF text extraction + OpenAI/Gemini + Excel compilation
├── data_profiler.py                # CSV/Excel data-quality detection + cleaning
├── donut_chart.py                   # matplotlib donut chart for the Data Profiler page
├── donation.py                       # generates the optional UPI donation QR
├── updater.py                         # checks GitHub Releases for newer versions
├── version.py                          # current app version (bump before each release)
├── .env                                 # your API key(s) (created by the app, git-ignored)
├── assets/
│   └── icon.ico                          # app icon
├── .github/
│   └── workflows/
│       └── build.yml                      # auto-builds & releases the .exe on tag push
├── requirements.txt
├── build_exe.bat                        # builds a Windows .exe locally
└── .gitignore
```

## 8. Publishing to GitHub

```cmd
cd linkharvest
git init
git add .
git commit -m "Initial commit: LinkHarvest"
git branch -M main
git remote add origin https://github.com/<your-username>/linkharvest.git
git push -u origin main
```

Then follow section 4 above to cut your first release.
