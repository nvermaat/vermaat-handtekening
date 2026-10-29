#!/usr/bin/env bash
# Zet het webadres van de invoegtoepassing in de bestanden.
# Gebruik:  ./set-host.sh https://www.vermaat-it.nl/handtekening
set -euo pipefail
HOST="${1:?Geef het https adres op}"
HOST="${HOST%/}"
case "$HOST" in https://*) ;; *) echo "Het adres moet met https:// beginnen." >&2; exit 1;; esac
ORIGIN="$(printf '%s' "$HOST" | sed -E 's#^(https://[^/]+).*#\1#')"
DIR="$(cd "$(dirname "$0")" && pwd)"
for f in manifest.xml signature.js launchevent.js; do
  if grep -q '{{HOST}}\|{{ORIGIN}}' "$DIR/$f"; then
    sed -i.bak -e "s#{{HOST}}#$HOST#g" -e "s#{{ORIGIN}}#$ORIGIN#g" "$DIR/$f" && rm -f "$DIR/$f.bak"
    echo "Bijgewerkt: $f"
  else
    echo "$f is al ingesteld, overgeslagen."
  fi
done
echo "Klaar. Upload nu alle bestanden naar $HOST"
