#!/usr/bin/env bash
# Run this from the project folder whenever you add a new match:
#   ./deploy.sh
#
# It will:
#   1. Rebuild index.html from everything in matches/
#   2. Commit the changes
#   3. Push to GitHub (Vercel auto-deploys on push)
set -e  # stop immediately if any step fails

cd "$(dirname "$0")"   # always run from this script's own folder, no matter where you call it from

echo "── Building site from matches/ ──"
python3 build.py

echo ""
echo "── Checking for changes ──"
if git diff --quiet && git diff --cached --quiet; then
  echo "Nothing changed — index.html is already up to date. Nothing to push."
  exit 0
fi

echo ""
echo "── Committing ──"
git add .
git commit -m "Update site — $(date '+%Y-%m-%d %H:%M')"

echo ""
echo "── Pushing to GitHub (Vercel will auto-deploy) ──"
git push

echo ""
echo "Done! Check the Deployments tab on Vercel in a few seconds."
