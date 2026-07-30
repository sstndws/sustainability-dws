#!/usr/bin/env bash
# Render HTML slide deck to landscape PDF via headless Chrome.
set -euo pipefail

HTML="/workspace/docs/slide-deck-source.html"
OUT="/workspace/docs/Sustainability-Dashboard-Slide-Deck-EN.pdf"

google-chrome \
  --headless=new \
  --disable-gpu \
  --no-sandbox \
  --disable-dev-shm-usage \
  --user-data-dir=/tmp/chrome-pdf-profile \
  --run-all-compositor-stages-before-draw \
  --virtual-time-budget=8000 \
  --print-to-pdf="$OUT" \
  "file://$HTML"

echo "Generated: $OUT ($(du -h "$OUT" | cut -f1))"

cp "$OUT" "/workspace/docs/Sustainability-Dashboard-Data-Flow.pdf"
cp "$OUT" "/workspace/docs/Sustainability-Dashboard-Slide-Deck-v2.pdf"
echo "Updated main PDF copies"
