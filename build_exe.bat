@echo off
echo Installing dependencies...
pip install -r requirements.txt

echo.
echo Building LinkHarvest.exe ...
pyinstaller --noconfirm --onefile --windowed ^
  --icon=assets\icon.ico ^
  --add-data "assets;assets" ^
  --collect-all pdfplumber ^
  --collect-all openai ^
  --collect-all google.generativeai ^
  --collect-all matplotlib ^
  --name LinkHarvest ^
  app.py

echo.
echo Build complete. Find your .exe in the "dist" folder.
pause
