@echo off
setlocal enabledelayedexpansion

REM release.bat — the ONLY way version.py and the git tag should get set.
REM Usage:  release.bat 1.3.1

if "%~1"=="" (
    echo Usage: release.bat ^<version^>
    echo Example: release.bat 1.3.1
    exit /b 1
)
set VERSION=%~1

echo.
echo === Step 1: Updating version.py to %VERSION% ===
powershell -NoProfile -Command ^
  "(Get-Content version.py) -replace '(?m)^__version__ = \".*\"', '__version__ = \"%VERSION%\"' | Set-Content version.py"

echo version.py now contains:
type version.py
echo.

set /p CONFIRM="Does that look right? (y/n): "
if /i not "%CONFIRM%"=="y" (
    echo Aborted. version.py was changed above but nothing was committed/pushed/tagged.
    echo Fix version.py manually if needed, then re-run this script.
    exit /b 1
)

echo.
echo === Step 2: git status (review before committing) ===
git status
echo.
set /p CONFIRM2="Proceed with commit + push + tag v%VERSION%? (y/n): "
if /i not "%CONFIRM2%"=="y" (
    echo Aborted before committing.
    exit /b 1
)

echo.
echo === Step 3: Commit + push ===
git add -A
git commit -m "Release v%VERSION%"
git push
if errorlevel 1 (
    echo git push FAILED — stopping before tagging so the tag never points at an unpushed commit.
    exit /b 1
)

echo.
echo === Step 4: Tag + push tag ===
git tag v%VERSION%
git push origin v%VERSION%
if errorlevel 1 (
    echo Tag push FAILED. Check the error above — you may need to delete and retry the tag:
    echo   git tag -d v%VERSION%
    exit /b 1
)

echo.
echo === Done ===
echo Tagged and pushed v%VERSION%. version.py matches the tag — this is the fix for the
echo "Update Required" loop that happened when they drifted apart before.
echo.
echo Now go check:
echo   1. GitHub Actions tab — wait for the v%VERSION% build to go green
echo   2. Releases page — confirm LinkHarvest.exe is attached to v%VERSION%
pause