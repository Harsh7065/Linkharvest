@echo off
setlocal

echo Installing dependencies...
pip install -r requirements.txt
pip install cython

echo.
echo === Step 1: Compiling backend modules with Cython ===
python setup.py build_ext --inplace --compiler=mingw32
if errorlevel 1 (
    echo Cython build FAILED - stopping before packaging.
    exit /b 1
)

echo.
echo === Step 2: Removing .py sources for the compiled modules ===
REM PyInstaller will otherwise prefer the plain .py over the .pyd if both
REM are present. Move the sources aside instead of deleting, so you don't
REM lose them - build_exe.bat restores them at the end.
if not exist _py_src_backup mkdir _py_src_backup
for %%m in (ai_assistant pdf_extractor dashboard_builder data_profiler donut_chart downloader sheet_editor donation updater) do (
    if exist %%m.py move /Y %%m.py _py_src_backup\ >nul
)

echo.
echo === Step 3: Building LinkHarvest.exe with PyInstaller ===
pyinstaller --noconfirm --onefile --windowed ^
  --icon=assets\icon.ico ^
  --add-data "assets;assets" ^
  --collect-all pdfplumber ^
  --collect-all openai ^
  --collect-all google.generativeai ^
  --collect-all matplotlib ^
  --collect-all openpyxl ^
  --collect-all qrcode ^
  --collect-all requests ^
  --collect-all dotenv ^
  --name LinkHarvest ^
  app.py

echo.
echo === Step 4: Restoring .py sources for local development ===
for %%m in (ai_assistant pdf_extractor dashboard_builder data_profiler donut_chart downloader sheet_editor donation updater) do (
    if exist _py_src_backup\%%m.py move /Y _py_src_backup\%%m.py . >nul
)
rmdir _py_src_backup 2>nul

echo.
echo Build complete. Find your .exe in the "dist" folder.
pause
