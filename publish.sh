#!/usr/bin/env bash
#
# publish.sh — one-shot publish of Global EV Market Pulse to GitHub
# ===================================================================
# Usage:
#   ./publish.sh <github-username> <repo-name> [--private]
#
# Example:
#   ./publish.sh Milad-Shabani global-ev-market-pulse
#
# What it does:
#   1. Regenerates the Excel workbook (so the repo always ships fresh data)
#   2. Initialises a git repo (if one doesn't already exist)
#   3. Creates the GitHub repo via the GitHub CLI (if `gh` is installed and
#      authenticated) — or prints the manual steps if it isn't
#   4. Commits everything and pushes to `main`
#
# Requirements:
#   - git
#   - Python 3 with requirements.txt installed (for step 1)
#   - Optional but recommended: GitHub CLI (`gh`), already logged in
#     (`gh auth login`) — otherwise create the empty repo on github.com
#     first and re-run this script, or follow the printed manual steps.

set -euo pipefail

GH_USER="${1:-}"
REPO_NAME="${2:-global-ev-market-pulse}"
VISIBILITY_FLAG="${3:---public}"

if [[ -z "$GH_USER" ]]; then
  echo "Usage: ./publish.sh <github-username> [repo-name] [--private|--public]"
  exit 1
fi

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

echo "▸ Step 1/4 — regenerating the Excel workbook from source data..."
if command -v python3 >/dev/null 2>&1; then
  (cd data && python3 generate_workbook.py)
  if command -v soffice >/dev/null 2>&1; then
    soffice --headless --convert-to xlsx --outdir data data/EV_Global_Outlook_2026.xlsx >/dev/null 2>&1 || \
      echo "  (LibreOffice recalculation skipped/failed — the workbook still opens fine in Excel/Google Sheets)"
  fi
  mkdir -p dashboard/data
  cp data/EV_Global_Outlook_2026.xlsx dashboard/data/EV_Global_Outlook_2026.xlsx
  echo "  ✓ Workbook regenerated and copied into dashboard/data/"
else
  echo "  ! python3 not found — skipping regeneration, using whatever workbook is already on disk."
fi

echo "▸ Step 2/4 — initialising git..."
if [[ ! -d .git ]]; then
  git init -b main
else
  echo "  (git repo already initialised)"
fi

echo "▸ Step 3/4 — creating the GitHub repository..."
REMOTE_URL="https://github.com/${GH_USER}/${REPO_NAME}.git"
if git remote get-url origin >/dev/null 2>&1; then
  echo "  (remote 'origin' already set: $(git remote get-url origin))"
elif command -v gh >/dev/null 2>&1 && gh auth status >/dev/null 2>&1; then
  gh repo create "${GH_USER}/${REPO_NAME}" \
    --source=. \
    --description "World EV production, sales, trade & 2035-outlook dashboard built on the IEA Global EV Outlook 2026" \
    --homepage "https://${GH_USER}.github.io/${REPO_NAME}/" \
    "${VISIBILITY_FLAG}" \
    --remote origin
  echo "  ✓ Repository created: ${REMOTE_URL}"
else
  echo "  ! GitHub CLI ('gh') not found or not authenticated."
  echo "    Create the repo manually at https://github.com/new (name: ${REPO_NAME}), then run:"
  echo "      git remote add origin ${REMOTE_URL}"
  echo "    ...and re-run this script."
  git remote add origin "${REMOTE_URL}" 2>/dev/null || true
fi

echo "▸ Step 4/4 — committing and pushing..."
git add -A
git commit -m "Publish Global EV Market Pulse dashboard (IEA Global EV Outlook 2026)" || \
  echo "  (nothing new to commit)"
git push -u origin main

echo ""
echo "✅ Done. Next steps on GitHub:"
echo "   1. Settings → Pages → Source: 'GitHub Actions' (one-time setup)"
echo "   2. The 'Deploy dashboard to GitHub Pages' workflow will then publish"
echo "      the dashboard automatically on every push to main, at:"
echo "      https://${GH_USER}.github.io/${REPO_NAME}/"
