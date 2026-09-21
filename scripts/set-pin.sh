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
URL="https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:commit"

BODY=$(cat <<EOF
{
  "writes": [
    {
      "update": {
        "name": "projects/${PROJECT_ID}/databases/(default)/documents/tracker/security",
        "fields": {
          "pin": {
            "stringValue": "${PIN}"
          }
        }
      }
    }
  ]
}
EOF
)

RESPONSE="$(curl --silent --show-error \
  -X POST \
  -H "Authorization: Bearer ${TOKEN}" \
  -H "Content-Type: application/json" \
  "$URL" \
  --data "$BODY")"

if echo "$RESPONSE" | grep -q '"error"'; then
  echo "$RESPONSE"
  exit 1
fi

echo "Tracker PIN updated."
