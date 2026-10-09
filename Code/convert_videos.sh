#!/bin/bash
#
# convert_videos.sh  (v2)
#
# Batch-converts all .mov files in a folder (and its subfolders) to
# web-friendly H.264 MP4, preserving the original folder structure.
#
# USAGE:
#   bash convert_videos.sh /path/to/input_folder /path/to/output_folder
#
# OPTIONAL: add --webm as a third argument to also generate a WebM
# fallback copy of each file (slower, roughly doubles processing time).
#
#   bash convert_videos.sh /path/to/input_folder /path/to/output_folder --webm
#
# WHAT IT DOES:
#   - Finds every .mov file under the input folder, including subfolders
#   - Converts each to H.264/MP4 (universally playable, no plugin needed)
#   - Mirrors the input folder's structure in the output folder
#   - Skips files that have already been converted (safe to re-run)
#   - Deletes partial output if a conversion fails, so a re-run retries it
#   - Prints a summary of successes/failures at the end
#
# CHANGES FROM v1:
#   - ffmpeg now runs with -nostdin. Without it, ffmpeg swallowed the
#     file list from find and the loop stopped after the first file.
#   - Added -pix_fmt yuv420p so every clip plays in every browser.
#   - Failed conversions remove their partial .mp4.
#   - Trailing slashes on the folder arguments are handled.
#
# REQUIREMENTS:
#   ffmpeg must be installed and on your PATH. If you installed it with
#   conda, activate that environment first (e.g. conda activate video).
#

set -uo pipefail

INPUT_DIR="${1:-}"
OUTPUT_DIR="${2:-}"
MAKE_WEBM="${3:-}"

if [[ -z "$INPUT_DIR" || -z "$OUTPUT_DIR" ]]; then
  echo "Usage: $0 <input_folder> <output_folder> [--webm]"
  exit 1
fi

# Strip any trailing slashes so relative-path math works
INPUT_DIR="${INPUT_DIR%/}"
OUTPUT_DIR="${OUTPUT_DIR%/}"

if ! command -v ffmpeg &> /dev/null; then
  echo "ERROR: ffmpeg is not installed or not on your PATH."
  echo "If you installed it with conda, run: conda activate video"
  exit 1
fi

if [[ ! -d "$INPUT_DIR" ]]; then
  echo "ERROR: Input folder not found: $INPUT_DIR"
  exit 1
fi

mkdir -p "$OUTPUT_DIR"

SUCCESS_COUNT=0
SKIP_COUNT=0
FAIL_COUNT=0
FAILED_FILES=()

echo "Scanning for .mov files in: $INPUT_DIR"
echo "Output will be written to: $OUTPUT_DIR"
[[ "$MAKE_WEBM" == "--webm" ]] && echo "Also generating WebM fallback copies."
echo "----------------------------------------------"

# Find every .mov file (case-insensitive), preserving relative paths.
# Hidden macOS resource files (._name.mov) are ignored.
while IFS= read -r -d '' infile; do
  # Compute the file's path relative to INPUT_DIR
  relpath="${infile#"$INPUT_DIR"/}"
  reldir=$(dirname "$relpath")
  base=$(basename "$relpath")
  base_noext="${base%.*}"

  outdir="$OUTPUT_DIR/$reldir"
  mkdir -p "$outdir"

  outfile_mp4="$outdir/${base_noext}.mp4"
  outfile_webm="$outdir/${base_noext}.webm"

  # Skip if already converted
  if [[ -f "$outfile_mp4" ]]; then
    echo "SKIP (already exists): $relpath"
    SKIP_COUNT=$((SKIP_COUNT + 1))
    continue
  fi

  echo "Converting: $relpath"

  if ffmpeg -nostdin -y -i "$infile" \
      -c:v libx264 -crf 23 -preset medium -pix_fmt yuv420p \
      -c:a aac -b:a 128k \
      -movflags +faststart \
      -vf "scale='min(1280,iw)':-2" \
      "$outfile_mp4" \
      -loglevel error; then

    SUCCESS_COUNT=$((SUCCESS_COUNT + 1))

    if [[ "$MAKE_WEBM" == "--webm" ]]; then
      echo "  Also generating WebM fallback..."
      ffmpeg -nostdin -y -i "$infile" \
        -c:v libvpx-vp9 -crf 32 -b:v 0 -pix_fmt yuv420p \
        -c:a libopus -b:a 128k \
        -vf "scale='min(1280,iw)':-2" \
        "$outfile_webm" \
        -loglevel error \
        || { echo "  WebM failed for: $relpath"; rm -f "$outfile_webm"; }
    fi
  else
    echo "  FAILED: $relpath"
    rm -f "$outfile_mp4"
    FAIL_COUNT=$((FAIL_COUNT + 1))
    FAILED_FILES+=("$relpath")
  fi

done < <(find "$INPUT_DIR" -type f -iname "*.mov" ! -name "._*" -print0)

echo "----------------------------------------------"
echo "Done."
echo "  Converted: $SUCCESS_COUNT"
echo "  Skipped (already existed): $SKIP_COUNT"
echo "  Failed: $FAIL_COUNT"

if [[ $FAIL_COUNT -gt 0 ]]; then
  echo ""
  echo "Files that failed to convert:"
  for f in "${FAILED_FILES[@]}"; do
    echo "  - $f"
  done
fi
