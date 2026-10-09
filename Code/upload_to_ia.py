#!/usr/bin/env python3
"""
upload_to_ia.py

Uploads the converted CFB Key MP4s to ONE Internet Archive item, keeping the
folder structure (media/..., html/...), and writes a manifest CSV with one row
per video. That manifest becomes the starting point for the 'videos' sheet of
the species spreadsheet.

SAFE BY DEFAULT: without --go this is a DRY RUN. It scans, computes checksums,
writes the manifest, and prints what it WOULD upload. Nothing is sent.

USAGE (run inside the conda 'video' environment):

  # 1. Dry run (sends nothing)
  python upload_to_ia.py --src "$ZK/cfbkey_mp4" --identifier cfbkey-videos

  # 2. Rehearsal: 3 files into a throwaway item, then check with `ia list`
  python upload_to_ia.py --src "$ZK/cfbkey_mp4" --identifier cfbkey-test-003 --limit 3 --go

  # 3. The real thing (safe to re-run; files already uploaded are skipped)
  python upload_to_ia.py --src "$ZK/cfbkey_mp4" --identifier cfbkey-videos --go

OPTIONS:
  --manifest PATH   where to write the manifest CSV (default: video_manifest.csv
                    in the folder you run the script from)
  --limit N         only handle the first N files (for rehearsals)
  --exclude PATTERN skip files whose path matches (repeatable; case-insensitive;
                    wildcards allowed), e.g. --exclude "*dontuse*"
  --title TEXT      override the item title, e.g. --title "TEST ONLY - ignore"
  --no-derive       tell Internet Archive not to make thumbnails/derivative
                    formats (faster, less clutter; the original MP4s are what
                    the website uses either way)

Requires: the 'internetarchive' package (already installed) and a configured
login (ia configure). ffprobe is used if available to record duration/audio.
"""

import argparse
import csv
import fnmatch
import hashlib
import json
import re
import subprocess
import sys
from pathlib import Path

# ----------------------------------------------------------------------------
# EDIT THESE before the real upload. Item METADATA can be changed later with
# `ia metadata --modify`, but the identifier (--identifier) can never change.
# ----------------------------------------------------------------------------
COLLECTION = "opensource_movies"
MEDIATYPE = "movies"
TITLE = "CFB Zooplankton Key: video clips"
DESCRIPTION = (
    "Video clips supporting \"An Image-Based Key to the Zooplankton of North "
    "America,\" produced by the University of New Hampshire Center for "
    "Freshwater Biology (https://cfb.unh.edu/cfbkey). Clips were converted "
    "from the original QuickTime (.mov) files to H.264 MP4. Folders mirror the "
    "media folders of the original website."
)
CREATOR = "Center for Freshwater Biology, University of New Hampshire"  # confirm credit line with Amanda
SUBJECTS = ["zooplankton", "freshwater", "plankton", "taxonomy", "Cladocera", "Copepoda", "Rotifera"]
LICENSE_URL = ""  # e.g. a Creative Commons URL once decided; leave "" to omit for now
# ----------------------------------------------------------------------------

CHUNK = 10  # files per upload call, so an interruption loses little


def clean_remote(rel: str) -> str:
    """Remote file path: keep folders, replace spaces/odd characters with _."""
    s = rel.replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9._/\-]", "_", s)


def md5_of(path: Path) -> str:
    h = hashlib.md5()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def probe(path: Path):
    """Return (duration_seconds or '', has_audio True/False/'')."""
    try:
        out = subprocess.run(
            ["ffprobe", "-v", "error",
             "-show_entries", "format=duration:stream=codec_type",
             "-of", "json", str(path)],
            capture_output=True, text=True, timeout=60)
        data = json.loads(out.stdout)
        dur = round(float(data["format"]["duration"]), 2)
        audio = any(s.get("codec_type") == "audio" for s in data.get("streams", []))
        return dur, audio
    except Exception:
        return "", ""


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", required=True, help="folder holding the converted MP4s (e.g. cfbkey_mp4)")
    ap.add_argument("--identifier", required=True, help="Internet Archive item identifier")
    ap.add_argument("--manifest", default="video_manifest.csv")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--exclude", action="append", default=[], metavar="PATTERN")
    ap.add_argument("--title", default=None, help="override the item title (use for rehearsals)")
    ap.add_argument("--no-derive", action="store_true")
    ap.add_argument("--go", action="store_true", help="actually upload (otherwise dry run)")
    args = ap.parse_args()

    src = Path(args.src).expanduser().resolve()
    if not src.is_dir():
        sys.exit(f"ERROR: source folder not found: {src}")

    files = sorted(p for p in src.rglob("*")
                   if p.is_file() and p.suffix.lower() == ".mp4" and not p.name.startswith("._"))
    if args.exclude:
        before = len(files)
        files = [p for p in files
                 if not any(fnmatch.fnmatch(p.relative_to(src).as_posix().lower(), pat.lower())
                            for pat in args.exclude)]
        print(f"Excluded {before - len(files)} file(s) by --exclude.")
    if args.limit:
        files = files[:args.limit]
    if not files:
        sys.exit("ERROR: no .mp4 files found under the source folder.")

    # Build rows, and make sure cleaned remote names don't collide
    rows, seen = [], {}
    for p in files:
        rel = p.relative_to(src).as_posix()
        remote = clean_remote(rel)
        if remote in seen:
            sys.exit(f"ERROR: two files would get the same remote name:\n  {seen[remote]}\n  {rel}")
        seen[remote] = rel
        dur, audio = probe(p)
        rows.append({
            "video_id": clean_remote(rel[:-4]),   # path without .mp4
            "local_path": rel,
            "remote_path": remote,
            "url": f"https://archive.org/download/{args.identifier}/{remote}",
            "size_bytes": p.stat().st_size,
            "duration_s": dur,
            "has_audio": audio,
            "md5": md5_of(p),
            "status": "",
        })

    # What is already in the item? (offline or new item: assume nothing)
    existing, item_exists = {}, False
    try:
        from internetarchive import get_item
        item = get_item(args.identifier)
        item_exists = bool(item.exists)
        if item_exists:
            existing = {f["name"]: f.get("md5") for f in item.files}
    except Exception as e:
        print(f"(Could not check archive.org: {e}. Assuming nothing is uploaded yet.)")

    todo = []
    for r in rows:
        if existing.get(r["remote_path"]) == r["md5"]:
            r["status"] = "already uploaded"
        else:
            r["status"] = "to upload"
            todo.append(r)

    total_mb = sum(r["size_bytes"] for r in todo) / 1e6
    print(f"Source:      {src}")
    print(f"Identifier:  {args.identifier}  ({'exists' if item_exists else 'new item'})")
    print(f"Found:       {len(rows)} MP4 files")
    print(f"Already up:  {len(rows) - len(todo)}")
    print(f"To upload:   {len(todo)} files, {total_mb:.1f} MB")
    print("Examples (local path  ->  URL):")
    for r in todo[:3]:
        print(f"  {r['local_path']}\n    -> {r['url']}")

    def write_manifest():
        with open(args.manifest, "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)

    write_manifest()
    print(f"Manifest written: {Path(args.manifest).resolve()}")

    if not args.go:
        print("\nDRY RUN: nothing was uploaded. Add --go to upload.")
        return
    if not todo:
        print("\nNothing to upload.")
        return

    from internetarchive import upload

    metadata = None
    if not item_exists:
        metadata = {
            "mediatype": MEDIATYPE,
            "collection": COLLECTION,
            "title": args.title or TITLE,
            "description": DESCRIPTION,
            "creator": CREATOR,
            "subject": SUBJECTS,
        }
        if LICENSE_URL:
            metadata["licenseurl"] = LICENSE_URL

    failed = []
    for i in range(0, len(todo), CHUNK):
        chunk = todo[i:i + CHUNK]
        print(f"\nUploading files {i + 1}-{i + len(chunk)} of {len(todo)} ...")
        try:
            resps = upload(
                args.identifier,
                files={r["remote_path"]: str(src / r["local_path"]) for r in chunk},
                metadata=metadata,
                queue_derive=not args.no_derive,
                retries=5, retries_sleep=10,
                verbose=True)
            bad = [x for x in resps if x is not None and getattr(x, "status_code", 200) != 200]
            if bad:
                raise RuntimeError(f"{len(bad)} file(s) returned an error status")
            for r in chunk:
                r["status"] = "uploaded"
            metadata = None  # item exists now; only send metadata once
        except Exception as e:
            print(f"  FAILED this batch: {e}")
            for r in chunk:
                r["status"] = "FAILED"
                failed.append(r["local_path"])
        write_manifest()

    print("\n----------------------------------------------")
    print(f"Uploaded: {sum(r['status'] == 'uploaded' for r in rows)}   "
          f"Failed: {len(failed)}   "
          f"Already there: {sum(r['status'] == 'already uploaded' for r in rows)}")
    if failed:
        print("Failed files (re-run the same command to retry them):")
        for f in failed:
            print("  -", f)
        sys.exit(1)
    print(f"Check later with:  ia list {args.identifier} | wc -l")
    print("(Archive.org takes a while to list new files; the count may lag.)")


if __name__ == "__main__":
    main()
