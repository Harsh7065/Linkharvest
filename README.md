# LinkHarvest (Windows)

A desktop app that scans an Excel sheet for links (images, videos, or audio)
and downloads them in parallel, with a modern UI and a live progress bar.

Originally a VBA macro (`Download_All_Row_Images`), rebuilt in Python with:
- A modern GUI (dark theme, `customtkinter`)
- Custom row/column ranges — download everything, or just one specific
  row/column
- Live progress bar with percentage
- Adjustable thread count and timeout (with a warning if you move away
  from the tested defaults)
- File-picker for the Excel file and the save folder

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

## 3. Support / donation

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

## 4. Distributing to other people (no Python required)

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

## 5. Project structure

```
LinkHarvest/
├── app.py                      # GUI (entry point)
├── downloader.py                # scanning + downloading logic
├── donation.py                   # generates the optional UPI donation QR
├── assets/
│   └── icon.ico                  # app icon
├── .github/
│   └── workflows/
│       └── build.yml              # auto-builds & releases the .exe on tag push
├── requirements.txt
├── build_exe.bat                 # builds a Windows .exe locally
└── .gitignore
```

## 6. Publishing to GitHub

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
