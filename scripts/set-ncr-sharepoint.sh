#!/usr/bin/env bash
set -euo pipefail

PROJECT_ID="design-hours-tracker"

read -r -p "Paste the SharePoint Excel link: " SHARE_URL
read -r -p "Microsoft Entra Application (client) ID: " CLIENT_ID
read -r -p "Microsoft Entra Directory (tenant) ID [organizations]: " TENANT
TENANT="${TENANT:-organizations}"

if [[ -z "$SHARE_URL" || -z "$CLIENT_ID" ]]; then
  echo "SharePoint link and client ID are required."
  exit 1
fi

TOKEN="$(gcloud auth print-access-token)"
URL="https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:commit"

escape_json() {
  python3 -c 'import json,sys; print(json.dumps(sys.stdin.read().rstrip("\n")))' 
}

SHARE_JSON=$(printf '%s' "$SHARE_URL" | escape_json)
CLIENT_JSON=$(printf '%s' "$CLIENT_ID" | escape_json)
TENANT_JSON=$(printf '%s' "$TENANT" | escape_json)

BODY=$(cat <<EOF
{
  "writes": [
    {
      "update": {
        "name": "projects/${PROJECT_ID}/databases/(default)/documents/tracker/ncrSource",
        "fields": {
          "shareUrl": {"stringValue": ${SHARE_JSON}},
          "clientId": {"stringValue": ${CLIENT_JSON}},
          "tenant": {"stringValue": ${TENANT_JSON}}
        }
      }
    }
  ]
}
EOF
)

RESPONSE="$(curl --silent --show-error -X POST \
  -H "Authorization: Bearer ${TOKEN}" \
  -H "Content-Type: application/json" \
  "$URL" \
  --data "$BODY")"

if echo "$RESPONSE" | grep -q '"error"'; then
  echo "$RESPONSE"
  exit 1
fi

echo "NCR SharePoint source configured."
