#!/bin/bash
# Create a folder for a new shipment.
#
#   scripts/new_shipment.sh "Czech Healthcare" 196652
#
# Makes shipments/<today>_<Client>_<Reference>/ with a README stub.
# Every deliverable for that shipment goes in that folder — one folder per
# shipment, so nothing gets mixed up with another one.
set -euo pipefail

if [ $# -lt 2 ]; then
  echo "usage: $0 \"<Client>\" <ShipmentNumber>" >&2
  exit 1
fi

CLIENT_RAW="$1"
REF_RAW="$2"
slug() { printf '%s' "$1" | tr ' ' '_' | tr -cd 'A-Za-z0-9._-'; }
CLIENT="$(slug "$CLIENT_RAW")"
REF="$(slug "$REF_RAW")"
DATE="$(date +%F)"

ROOT="$(git -C "$(dirname "$0")" rev-parse --show-toplevel)"
DIR="$ROOT/shipments/${DATE}_${CLIENT}_${REF}"

if [ -d "$DIR" ]; then
  echo "Already exists: $DIR"
  exit 0
fi

mkdir -p "$DIR"
cat > "$DIR/README.md" <<README
# ${CLIENT_RAW} ${REF_RAW}

- **Reference:** \`${CLIENT}_${REF}\`
- **Last worked on:** ${DATE}
- **Chat:** _paste this chat's link here_
- **Where it was left:** _in progress_
- **Still open:** _nothing yet_

## Files

- _deliverable goes here: HS_Code_Summary_${CLIENT}_${REF}.xlsx_

## Continue this shipment

Open the chat above, or start a new chat on this repo and say:

> Continue shipment \`${CLIENT}_${REF}\` — files are in \`shipments/${DATE}_${CLIENT}_${REF}/\`.
README

echo "$DIR"
