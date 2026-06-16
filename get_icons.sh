#!/bin/bash
mkdir -p icons_svg
BASE="https://unpkg.com/lucide-static@1.20.0/icons"
ICONS=(
  tree-pine building-2 handshake coins link-2 building users footprints
  monitor bar-chart-3 globe banknote satellite-dish user-round percent
  puzzle map-pin shield-check clipboard-list credit-card megaphone scale
  server cpu leaf circle-dollar-sign trending-up badge-check landmark
  network repeat
)
for name in "${ICONS[@]}"; do
  curl -sL "$BASE/$name.svg" -o "icons_svg/$name.svg"
  if grep -q "<svg" "icons_svg/$name.svg"; then echo "OK   $name"; else echo "FAIL $name"; fi
done