#!/usr/bin/env bash
# Weekly NuBrakes Operations Business Review.
#
# Builds everything the Thursday delivery needs, in order:
#   1. compute.py  -> out/metrics.json          (fetches the feeds, computes the week)
#   2. render.py   -> out/NuBrakes_OBR_*.pptx   (the 14-slide deck)
#   3. soffice     -> out/*.pdf                 (print/share copy)
#   4. pdftoppm    -> out/slides/slide-NN.png   (one image per slide, for the page)
#   5. make_page.py-> out/page.html             (the hosted review page)
#
# Pass --asof YYYY-MM-DD (any day inside the target week) to rebuild an earlier week.
#
# READ THE CONSOLE BEFORE SHIPPING. A WARNING means the week is incomplete, a feed
# was stale, or payroll hours are missing. Never publish a deck with a partial week
# without saying so.
set -euo pipefail
cd "$(dirname "$0")"

ASOF=()
if [ "${1:-}" = "--asof" ] && [ -n "${2:-}" ]; then ASOF=(--asof "$2"); fi

need() {
  command -v "$1" >/dev/null 2>&1 || {
    echo "ERROR: '$1' not found. $2" >&2; exit 1; }
}
need python3 "Install Python 3.9 or newer."
need soffice  "Install LibreOffice (provides soffice) to export the PDF and slide images."
need pdftoppm "Install poppler-utils (provides pdftoppm) to render the slide images."

python3 -c "import pptx, matplotlib" 2>/dev/null || {
  echo "ERROR: missing Python packages. Run: pip install python-pptx matplotlib openpyxl" >&2
  exit 1; }

echo "--- 1/5 compute"
python3 scripts/compute.py --out out --data data "${ASOF[@]}"

echo "--- 2/5 deck"
python3 scripts/render.py --metrics out/metrics.json --out out

PPTX="$(ls -1t out/NuBrakes_OBR_*.pptx | head -1)"
STEM="$(basename "${PPTX%.pptx}")"

echo "--- 3/5 pdf"
rm -f "out/$STEM.pdf"
soffice --headless --norestore --convert-to pdf "$PPTX" --outdir out >/dev/null 2>&1
[ -f "out/$STEM.pdf" ] || { echo "ERROR: PDF export produced nothing." >&2; exit 1; }

echo "--- 4/5 slide images"
rm -rf out/slides && mkdir -p out/slides
pdftoppm -png -r 110 "out/$STEM.pdf" out/slides/slide
# pdftoppm pads to the page count (slide-1 or slide-01); normalise to two digits.
( cd out/slides
  for f in slide-*.png; do
    n="${f#slide-}"; n="${n%.png}"; n="$((10#$n))"
    printf -v want "slide-%02d.png" "$n"
    [ "$f" != "$want" ] && mv "$f" "$want" || true
  done )

echo "--- 5/5 page"
python3 scripts/make_page.py . out

echo
echo "deck   $PPTX"
echo "pdf    out/$STEM.pdf"
echo "slides $(ls -1 out/slides/slide-*.png | wc -l | tr -d ' ') images in out/slides/"
echo "page   out/page.html"
