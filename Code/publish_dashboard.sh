#!/usr/bin/env bash
# publish_dashboard.sh
#
# Builds the food web dashboard from the key's data and writes it into the
# CyanoCounter repository, which is where it gets published.
#
#   bash "$ZK/Code/publish_dashboard.sh"
#
# The dashboard is BUILT here, in the key project, because that is where the trait
# tables live. It is PUBLISHED on the counter site, because that is where the people
# counting are, and because two pages on the same site can hand data to each other
# without a file in between.
#
# After this runs, commit and push the counter repo:
#   cd "$COUNTER" && git add dashboard.html && git commit -m "Update the dashboard" && git push

set -euo pipefail

ZK="${ZK:-$HOME/Documents/Zooplankton Key}"
COUNTER="${COUNTER:-$HOME/Documents/GitHub/cyanocounter}"

PY="$(command -v python3 || command -v python)"

# the name-difference list, used when it exists (Dolichospermum -> Anabaena and so on)
ALIASES=""
[ -f "$ZK/data/phyco_aliases.csv" ] && ALIASES="$ZK/data/phyco_aliases.csv"

if [ ! -d "$COUNTER" ]; then
  echo "ERROR: the counter repository is not at:"
  echo "   $COUNTER"
  echo "Set it for this run, for example:"
  echo "   COUNTER=\"\$HOME/Documents/GitHub/cyanocounter\" bash \"\$ZK/Code/publish_dashboard.sh\""
  exit 1
fi

for f in "data/cfbkey_workbook.xlsx" "data/phyto_traits.xlsx"; do
  if [ ! -f "$ZK/$f" ]; then
    echo "ERROR: missing $ZK/$f"
    [ "$f" = "data/phyto_traits.xlsx" ] && echo "   Run build_phyto_traits.py first."
    exit 1
  fi
done

echo "Building the dashboard from the key's trait tables..."
"$PY" "$ZK/Code/build_dashboard.py" \
  --zooplankton  "$ZK/data/cfbkey_workbook.xlsx" \
  --phytoplankton "$ZK/data/phyto_traits.xlsx" \
  --rules        "$ZK/data/foodweb_rules.csv" \
  ${ALIASES:+--aliases "$ALIASES"} \
  --out          "$COUNTER/dashboard.html"

echo
echo "Written to: $COUNTER/dashboard.html"
echo
echo "Next, publish it:"
echo "   cd \"$COUNTER\""
echo "   git add dashboard.html"
echo "   git commit -m \"Update the food web dashboard\""
echo "   git push"
