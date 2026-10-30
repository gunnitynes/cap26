#!/bin/sh
# Stamps a new BUILD into the page so open phones reload into it, then commits and pushes.
#   sh dev/publish.sh "what changed"
cd "$(dirname "$0")/.." || exit 1
sed -i '' "s/^const BUILD='[0-9]*';$/const BUILD='$(date +%s)';/" docs/index.html
node dev/test.mjs | tail -1 || exit 1
# never publish a page whose script does not parse, or whose update check lost its BUILD line
node -e "const h=require('fs').readFileSync('docs/index.html','utf8');const js=h.split('<script>').pop().split('</script>')[0];new Function(js);if(!/^const BUILD='[0-9]+';$/m.test(h))throw new Error('BUILD line missing')" || { echo 'page script broken: not published'; exit 1; }
git add -A && git commit -qm "$1

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>" && git push -q && git log --oneline -1
