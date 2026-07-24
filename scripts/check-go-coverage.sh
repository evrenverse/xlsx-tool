#!/usr/bin/env bash
set -euo pipefail

profile=${1:-coverage.out}
minimum=${2:-80}
coverage=$(go tool cover -func="$profile" | awk '/^total:/ {gsub(/%/, "", $3); print $3}')

if [[ -z "$coverage" ]]; then
  echo "Unable to read total coverage from $profile" >&2
  exit 1
fi

awk -v actual="$coverage" -v required="$minimum" \
  'BEGIN { if (actual + 0 < required + 0) exit 1 }' || {
  echo "Coverage ${coverage}% is below required ${minimum}%" >&2
  exit 1
}

echo "Coverage ${coverage}% meets required ${minimum}%"
