#!/bin/sh
# Usage: ./release.sh "commit message"  -- bumps the version so every iPad reloads itself, then pushes.
set -e
cd "$(dirname "$0")"
n=$(grep -o "const BUILD='r[0-9]*'" backstage.html | grep -o '[0-9]*')
next=$((n+1))
sed -i "s/const BUILD='r$n'/const BUILD='r$next'/" backstage.html
printf '{"paused": false, "message": "", "v": "r%s"}' "$next" > status.json
git add -A && git commit -q -m "$1 (r$next)" && git push -q
echo "pushed r$next"
