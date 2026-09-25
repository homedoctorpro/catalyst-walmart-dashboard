#!/usr/bin/env bash
# Push retailers_sync.gs to the Sheet's Apps Script project and cut a new
# version of the SAME web-app deployment, so the /exec URL never changes.
#
#   ./deploy_apps_script.sh "what changed"
#
# One-time setup is in google_apps_script/SETUP.md (clasp login + two ids).
#
# clasp refuses to run inside the Z: share ("Content directory is a symlink"),
# so the files are staged into a local scratch directory first.
set -euo pipefail
REPO="$(cd "$(dirname "$0")" && pwd)"
SRC="$REPO/google_apps_script"
STAGE="${LOCALAPPDATA:-$HOME}/catalyst-clasp"

[ -f "$SRC/.clasp.json" ] || { echo "google_apps_script/.clasp.json is missing — see SETUP.md"; exit 1; }

mkdir -p "$STAGE/src"
cp "$SRC/.clasp.json" "$STAGE/.clasp.json"
cp "$SRC/src/appsscript.json" "$STAGE/src/appsscript.json"
cp "$SRC/retailers_sync.gs" "$STAGE/src/retailers_sync.gs"
[ -f "$SRC/link_picks_oneoff.gs" ] && cp "$SRC/link_picks_oneoff.gs" "$STAGE/src/link_picks_oneoff.gs"

DEPLOYMENT_ID="$(cd "$STAGE" && python -c "import json;print(json.load(open('.clasp.json')).get('deploymentId',''))")"
[ -n "$DEPLOYMENT_ID" ] || { echo "deploymentId missing from .clasp.json — see SETUP.md"; exit 1; }

cd "$STAGE"
echo "→ pushing code"
clasp push -f
echo "→ deploying $DEPLOYMENT_ID"
clasp deploy -i "$DEPLOYMENT_ID" -d "${1:-update $(date -u +%Y-%m-%dT%H:%MZ)}"
echo "✓ live on the existing /exec URL"
