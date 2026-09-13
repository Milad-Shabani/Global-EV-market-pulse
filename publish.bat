@echo off
REM ============================================================
REM Publish this folder as a new public repo on github.com/Milad-Shabani
REM Requires: git, and GitHub CLI (gh) installed + logged in (gh auth login)
REM Optional: Python 3 (to regenerate the workbook/dashboard data before publishing)
REM ============================================================

set REPO_NAME=global-ev-market-pulse
set "REPO_DESC=World EV production, sales, trade, battery supply chain, charging infrastructure and 2035-outlook dashboard, built on the IEA Global EV Outlook 2026. Python-generated Excel workbook + interactive HTML dashboard with a D3 world map."

REM --- adjust this to wherever you unzipped/cloned the project locally
cd /d "C:\Users\MILAD\Desktop\global-ev-market-pulse"

REM --- set your git identity (safe to run every time)
git config --global user.name "Milad Shabani"
git config --global user.email "MILAD.SHABANI6515@GMAIL.COM"

REM --- optional: regenerate the workbook + dashboard data so the repo always
REM     ships fresh figures (skips quietly if Python isn't installed)
where python >nul 2>nul
if %errorlevel%==0 (
    pushd data
    python generate_workbook.py
    popd
)

REM --- init only if not already a repo
if not exist ".git" (
    git init
    git branch -M main
)

REM --- remove any leftover remote from a previous attempt
git remote remove origin 2>nul

git add .
git commit -m "Initial commit: Global EV Market Pulse dashboard"
git branch -M main

gh repo create %REPO_NAME% --public --source=. --remote=origin --push --description "%REPO_DESC%"

gh repo edit Milad-Shabani/%REPO_NAME% --homepage "https://Milad-Shabani.github.io/%REPO_NAME%/" --add-topic electric-vehicles --add-topic ev --add-topic data-visualization --add-topic dashboard --add-topic d3js --add-topic chartjs --add-topic python --add-topic excel --add-topic iea --add-topic market-intelligence

echo.
echo Done. Repo should now be live at:
echo https://github.com/Milad-Shabani/%REPO_NAME%
echo.
echo Next step (one-time): on GitHub, go to Settings ^> Pages ^> Source: "GitHub Actions".
echo The included workflow will then auto-deploy the dashboard to:
echo https://Milad-Shabani.github.io/%REPO_NAME%/
pause
