#!/bin/sh
# Rebuild everything and push to GitHub Pages.
set -e
cd "$(dirname "$0")"
python3 build.py
cp conferences_2027.html index.html
git add -A
git diff --cached --quiet && { echo "No changes to publish"; exit 0; }
git commit -m "Update conference data $(date +%Y-%m-%d)"
git push
