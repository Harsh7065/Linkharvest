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
  rows, extra spaces, and special characters, optionally fill missing
  numeric cells with that column's mean, then export a cleaned copy
- **Excel Editor page**: load a spreadsheet, describe an edit in plain
  English (e.g. "remove cancelled rows and sort by date"), and apply
  it directly to the sheet — with full undo/redo and a save button
- **Dashboard Builder page**: load a CSV/Excel file and get a reference
  dashboard — every column is classified (numeric/categorical/date/ID),
  a chart type is suggested for each with a plain-English reason, a
  live preview renders in-app, and it exports to a real, editable Excel
  workbook with native charts plus an AI-written summary of the logic
- **AI Assistant page**: a Copilot-style helper — attach up to 3 images
  (a table screenshot, a schema, an error message, a sketch) and describe
  what to write; it reads the images and generates SQL, DAX, Python,
  Excel formulas, VBA, or anything else you ask for
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
4. **Fill missing values with column mean** — a separate checkbox below
   the anomaly list. It only ever touches numeric columns (a text column
   has no valid "average" to fall back on), filling each missing numeric
   cell with that column's mean rounded to 2 decimals. Run this before
   checking "Missing Values" if you want numeric gaps filled instead of
   those rows being dropped — any remaining non-numeric gaps are still
   handled by the "Missing Values" checkbox as before.
5. Click **Clean & Export**, choose where to save (`.xlsx` or `.csv`),
   and the cleaned file is written and the containing folder opens
   automatically.

## 5. Using the Dashboard Builder page

Click **📊 Dashboard Builder** in the sidebar. This page turns any data
file into a reference dashboard — a starting point you can restyle in
Excel, not a finished polished report.

1. **Data file** — browse to a `.csv`, `.xlsx`, or `.xls` file.
2. **AI Engine** — reuses whichever key/model you already saved on the
   PDF Extractor page automatically; the AI is only used to write the
   plain-English summary, so this step is optional (uncheck "Also write
   an AI summary..." to skip it entirely and still get the charts).
3. Click **Build Dashboard Plan**. This runs three things:
   - **Column Classification** — every column is tagged as datetime,
     numeric, categorical, identifier, or text, based on its data type
     and how many distinct values it has.
   - **Suggested Charts** — a chart type is proposed per column
     (or column pair) using standard rules of thumb: a date column plus
     a numeric one becomes a line chart (trend over time), a low-cardinality
     category becomes a bar or pie chart, a numeric column on its own becomes
     a histogram (distribution), and two numeric columns together become a
     scatter plot (correlation) — each with a one-line reason shown next to it.
   - **Preview** — a live grid of those charts renders right in the app so
     you can sanity-check the plan before exporting anything.
4. **Dashboard Logic — Summary** — if an AI summary was requested, it
   explains *why* the data was grouped this way and what to watch out for
   (skewed categories, missing data, etc.) in plain English; if skipped or
   it fails, the charts and classification above still stand on their own.
5. Click **Export Dashboard Excel** and choose where to save. Unlike the
   in-app preview, this produces a real, editable `.xlsx` — a "Data" sheet
   with your raw rows, a "Dashboard" sheet with native Excel chart objects
   (restyle/resize/move them like any Excel chart), and a "Summary" sheet
   with the AI's explanation if one was generated.

## 6. Using the AI Assistant page

Click **🤖 AI Assistant** in the sidebar. This is a Copilot-style helper:
attach a few images for context, describe what you want written, and it
generates the logic — in whatever language fits the request.

1. **AI Engine** — reuses whichever key/model you already saved on the
   PDF Extractor page automatically. Use a vision-capable model (the
   suggested defaults both support images).
2. **Attach Images** — up to 3, optional. Screenshots of a spreadsheet,
   a database schema, an error message, a whiteboard sketch of a data
   flow — anything visual that gives the AI context. You can also ask
   without any images if the prompt is self-contained.
3. **What should it write?** — describe the logic you want, e.g.
   *"Write a SQL query that joins these two tables on customer_id"*,
   *"Turn this into a DAX measure"*, *"Write the VBA macro that does
   this"*, or *"Write the pandas code to reproduce this pivot table"*.
   If you don't name a language, it picks the most obviously appropriate
   one from the images and says which it chose.
4. Click **Ask AI Assistant**. The response appears in the box below:
   a short explanation of the approach, then the code in a fenced block,
   then (only if useful) a note on edge cases. Click **Copy** to grab it.
5. Column/table/field names in the generated code are grounded in what's
   actually visible in your images — if something isn't visible or is
   ambiguous, the assistant uses a clearly-marked placeholder instead of
   guessing a name that might not exist.

## 7. Using the Excel Editor page

Click **✏️ Excel Editor** in the sidebar. This page loads a spreadsheet
into the app and lets you edit it by describing what you want in plain
English, instead of clicking through menus.

1. **Spreadsheet** — browse to a `.csv`, `.xlsx`, or `.xls` file and
   click **Load**. If the workbook has multiple sheets, pick which one
   to work on from the **Sheet** dropdown — each sheet has its own
   independent undo history.
2. **AI Engine** — this reuses whichever key/model you already saved
   on the PDF Extractor page automatically; switch engine or key here
   only if you want this page to use something different.
3. **Describe the edit** — type what you want changed, e.g. "remove
   rows where Status is Cancelled, then sort by Date descending",
   "fill missing values in the Score column with the average", or
   "add a column called Total that's Price times Quantity". Click
   **Apply Edit**.
4. The AI never writes or runs arbitrary code — it maps your
   instruction onto a fixed set of safe operations (rename/drop a
   column, filter or sort rows, add a computed column, fill missing
   values, replace values, remove duplicates, change text case, round
   numbers, convert a column's type, and a few more). If it can't map
   your instruction to one of these, it explains why instead of
   guessing.
5. **↶ Undo / ↷ Redo** — every applied edit can be undone (and redone)
   per sheet, so it's safe to experiment.
6. **💾 Save** writes back to the file you loaded; **Save As...** lets
   you save a copy instead (`.xlsx` keeps every sheet, `.csv` only
   works if the workbook has a single sheet).
7. The **Sheet Preview** table shows your data live as edits are
   applied (large sheets show the first 300 rows in the preview, but
   edits and saving always apply to every row).

## 8. Support / donation

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

## 9. Distributing to other people (no Python required)

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

## 10. Project structure

```
LinkHarvest/
├── app.py                      # GUI (entry point, sidebar + all pages)
├── downloader.py                # link scanning + downloading logic
├── pdf_extractor.py               # PDF text extraction + OpenAI/Gemini + Excel compilation
├── data_profiler.py                # CSV/Excel data-quality detection + cleaning
├── sheet_editor.py                   # Excel Editor: AI edit planning + safe operations + undo/redo
├── dashboard_builder.py                # Dashboard Builder: column classification + chart plan + Excel export
├── ai_assistant.py                       # AI Assistant: image+prompt -> SQL/DAX/Python/Excel/VBA logic
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

## 11. Publishing to GitHub

```cmd
cd linkharvest
git init
git add .
git commit -m "Initial commit: LinkHarvest"
git branch -M main
git remote add origin https://github.com/<your-username>/linkharvest.git
git push -u origin main
```

Then follow section 9 above to cut your first release.
