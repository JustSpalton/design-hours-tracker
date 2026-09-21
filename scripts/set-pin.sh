#!/usr/bin/env bash
set -euo pipefail

PIN="${1:-}"
PROJECT_ID="design-hours-tracker"

if [[ ! "$PIN" =~ ^[0-9]{4}$ ]]; then
  echo "Usage: bash scripts/set-pin.sh 1234"
  echo "PIN must be exactly 4 digits."
  exit 1
fi

TOKEN="$(gcloud auth print-access-token)"
URL="https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents/tracker/security?updateMask.fieldPaths=pin"

curl --fail --silent --show-error \
  -X PATCH \
  -H "Authorization: Bearer ${TOKEN}" \
  -H "Content-Type: application/json" \
  "$URL" \
  --data "{"fields":{"pin":{"stringValue":"${PIN}"}}}" >/dev/null

echo "Tracker PIN updated."
