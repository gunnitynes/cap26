#!/bin/sh
# Stamps a new BUILD into the page so open phones reload into it, then commits and pushes.
#   sh dev/publish.sh "what changed"
cd "$(dirname "$0")/.." || exit 1
sed -i '' "s/const BUILD='[^']*'/const BUILD='$(date +%s)'/" docs/index.html
node dev/test.mjs | tail -1 || exit 1
git add -A && git commit -qm "$1

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" && git push -q && git log --oneline -1
