#!/usr/bin/env bash
# Push retailers_sync.gs to the Sheet's Apps Script project and cut a new
# version of the SAME web-app deployment, so the /exec URL never changes.
#
#   ./deploy_apps_script.sh "what changed"
#
# One-time setup is in google_apps_script/SETUP.md (clasp login + two ids).
set -euo pipefail
cd "$(dirname "$0")/google_apps_script"

[ -f .clasp.json ] || { echo "google_apps_script/.clasp.json is missing — see SETUP.md"; exit 1; }
DEPLOYMENT_ID="$(python -c "import json;print(json.load(open('.clasp.json')).get('deploymentId',''))")"
[ -n "$DEPLOYMENT_ID" ] || { echo "deploymentId missing from .clasp.json — see SETUP.md"; exit 1; }

# clasp pushes a directory; keep the single source file in src/ alongside the manifest
cp retailers_sync.gs src/retailers_sync.gs
[ -f link_picks_oneoff.gs ] && cp link_picks_oneoff.gs src/link_picks_oneoff.gs

echo "→ pushing code"
clasp push -f

echo "→ deploying $DEPLOYMENT_ID"
clasp deploy -i "$DEPLOYMENT_ID" -d "${1:-update $(date -u +%Y-%m-%dT%H:%MZ)}"

echo "✓ live on the existing /exec URL"
