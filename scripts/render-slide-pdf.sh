#!/usr/bin/env bash
# Render HTML slide deck to landscape PDF via headless Chrome.
set -euo pipefail

HTML="/workspace/docs/slide-deck-source.html"
OUT="/workspace/docs/Sustainability-Dashboard-Slide-Deck-v2.pdf"

google-chrome \
  --headless=new \
  --disable-gpu \
  --no-sandbox \
  --disable-dev-shm-usage \
  --run-all-compositor-stages-before-draw \
  --virtual-time-budget=8000 \
  --print-to-pdf="$OUT" \
  "file://$HTML"

echo "Generated: $OUT ($(du -h "$OUT" | cut -f1))"

# Also overwrite main file so PR link gets update
cp "$OUT" "/workspace/docs/Sustainability-Dashboard-Data-Flow.pdf"
echo "Updated: /workspace/docs/Sustainability-Dashboard-Data-Flow.pdf"
