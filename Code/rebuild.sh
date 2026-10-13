#!/bin/bash
#
# rebuild.sh  -  rebuild the website from the workbook, then open it.
#
# USAGE (in Terminal, from anywhere):
#   bash "$HOME/Documents/Zooplankton Key/Code/rebuild.sh"              # build everything
#   bash "$HOME/Documents/Zooplankton Key/Code/rebuild.sh" --limit 5    # quick test: first 5 species
#   bash "$HOME/Documents/Zooplankton Key/Code/rebuild.sh" --only bosmina-longirostris
#
# The site is written to the "docs" folder, which is what GitHub Pages publishes.
# Anything typed after the script name (--limit, --only, --content ...) is passed
# on to build_site.py.
#
# --redirects leaves a small forwarding page at every OLD web address (the
# html/Organisms/... and html/choices/... paths), so citations and bookmarks from the
# old site still work. Keep it on.
#
# While this is a TEST copy, the build adds a "do not index" tag so it does not turn
# up in search results beside the live UNH site. Remove --noindex from the last line
# when this site is meant to replace the old one.

ZK="$HOME/Documents/Zooplankton Key"
PY="$HOME/miniconda3/envs/video/bin/python"

# the workbook: in data/ (tidy, recommended) or still at the top level
WB="$ZK/data/cfbkey_workbook.xlsx"
[ -f "$WB" ] || WB="$ZK/cfbkey_workbook.xlsx"

if [ ! -x "$PY" ]; then
  echo "Could not find Python at $PY"
  echo "Open Terminal, run:  conda activate video  then  which python  and put that path in rebuild.sh"
  exit 1
fi
if [ ! -f "$WB" ]; then
  echo "Could not find the workbook. Looked in:"
  echo "  $ZK/data/cfbkey_workbook.xlsx"
  echo "  $ZK/cfbkey_workbook.xlsx"
  exit 1
fi

"$PY" "$ZK/Code/build_site.py" \
  --workbook "$WB" \
  --old-site "$ZK/cfbkey_local" \
  --content "$ZK/content" \
  --out "$ZK/docs" \
  --redirects \
  --noindex "$@" && open "$ZK/docs/index.html"
