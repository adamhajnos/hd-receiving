#!/bin/sh
# Usage: ./preview.sh  -- publishes the CURRENT working copy of backstage.html to /preview/ so it can be looked at before a release.
# iPads never load /preview/. It uses the REAL sheet, so do not tap Log/Finish/Send on it.
set -e
cd "$(dirname "$0")"
src="$PWD/backstage.html"; wt=$(mktemp -d)
git fetch -q origin main
git worktree add -q "$wt" origin/main --detach
mkdir -p "$wt/preview"
sed -e "s/const BUILD='r[0-9]*'/const BUILD='preview'/" -e "s|<body\([^>]*\)>|<body\1><div style=\"position:fixed;top:0;left:0;right:0;z-index:99999;background:#FF6B1A;color:#fff;font:700 13px/22px sans-serif;text-align:center;letter-spacing:.06em;pointer-events:none\">PREVIEW · REAL DATA · NOT LIVE</div>|" "$src" > "$wt/preview/backstage.html"
printf '{"paused": false, "message": "", "v": "preview"}' > "$wt/preview/status.json"
cd "$wt"; git add -A preview
if git diff --cached --quiet; then echo "preview already up to date"; else git commit -q -m "Preview build" && git -c push.negotiate=true push -q origin HEAD:main && echo "preview published"; fi
cd - >/dev/null; git worktree remove --force "$wt"
