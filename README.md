# CFB Zooplankton Key — rebuild (Version 6.0, in progress)

A rebuild of [An Image-Based Key to the Zooplankton of North America](https://cfb.unh.edu/cfbkey/html/begin.html),
the taxonomic key published by the University of New Hampshire Center for Freshwater Biology.

**Preview of this copy: https://jvthaney.github.io/cfbkey-test/**

This is a **test copy for review**, not the live key. The live key is still the UNH
site linked above and is unchanged. Pages here carry a "do not index" tag so this
copy does not appear in search results beside the real one.

## What is different from the old site

- Every page is generated from one spreadsheet, so a fact is edited once, in one place,
  instead of in hand-written HTML on hundreds of pages.
- The QuickTime videos are converted to MP4 and play in the browser with no plugin.
  They are hosted at the Internet Archive ([cfbkey-videos](https://archive.org/details/cfbkey-videos)).
- The pages work on phones and tablets, which the old fixed-width pages could not.
- Accessibility work toward WCAG 2.2 AA: real headings, keyboard access to everything that
  used to need a mouse, text alternatives for pictures, and no Flash or plugins.
  **Still in progress:** most pictures carry draft text alternatives awaiting review.

## How to read this repository

| Folder | What it is |
| --- | --- |
| `docs/` | **The website.** Generated: never edit by hand, it is rewritten on every build. This is the folder GitHub Pages publishes. |
| `data/` | The workbook (`cfbkey_workbook.xlsx`) that holds the species text, taxonomy, picture lists, video descriptions, the whole identification key and the wording shown in the footer, plus `video_manifest.csv`, the web address of each video. Dated backups are written to `data/workbook_backups/` and are not kept here. |
| `content/` | The ordinary text pages (Home, Instructions, Purpose, History, References, Contact, Copyright, Barcoding, Video introduction) as plain text files, and the pictures used on them. |
| `Code/` | The scripts that read the workbook and write `docs/`. |
| `site_src/` | Optional: editable copies of the page templates and stylesheet. Present only if someone has run `build_site.py --init`. |
| `guide/` | The editing guide (`cfbkey-editing-guide.docx`). |

## Files that are not in this repository

Three things are too large to keep here. Whoever maintains the key needs to know where
they are, because the workbook refers to them.

| What | Where it lives | Why it matters |
| --- | --- | --- |
| **`cfbkey_local`** (about 700 MB) | Jess Trout-Haney's Mac, and the project OneDrive folder shared by Amanda McQuaid | The original pictures, and the only complete copy of the old site. The `file` column of the workbook's `images` sheet points into this folder. Needed to **add or replace** any picture. |
| **The `.mov` originals** (about 545 MB) | The same OneDrive folder. `cfbkey_mp4` holds the converted MP4s. | The source footage, and the safety net behind the published videos. |
| **The videos as published** (101 clips) | [archive.org/details/cfbkey-videos](https://archive.org/details/cfbkey-videos), under the project account `cfbzoopkey@gmail.com` | What the site actually plays. Each address is listed in `data/video_manifest.csv`. |

A rebuild does **not** need `cfbkey_local`. If it is absent the build says so and reuses the
pictures already in `docs/`, so text, taxonomy, alt text and video descriptions can all be
changed on any computer with this repository. Only adding or replacing a picture needs the
originals.

**Before anyone leaves the project:** make sure at least two people can reach the OneDrive
folder and the Internet Archive account, and that `cfbkey_local` exists somewhere other than
one laptop. Without it, new species photographs cannot be added to the key.

## Making a change

Content changes are spreadsheet or text-file edits; no programming:

1. Edit `data/cfbkey_workbook.xlsx` (species text, taxonomy, alt text, video descriptions,
   the key) or a file in `content/` (ordinary page text).
2. Rebuild: `bash Code/rebuild.sh`
3. Read the short report it prints, then check the preview it opens.
4. Commit and push. GitHub Pages republishes within about a minute.

The editing guide in `guide/` explains each job in detail: which sheet to use for what, how to
read the build report, how to publish, and where the original files live.

The footer wording, the site title and the citation are in the workbook's `settings` sheet, so
a typo there is a spreadsheet edit like any other. Links can be written inside those values as
`[the words](https://the-address)`.

## Still to do

- Review the draft text alternatives for about 1,100 pictures and 315 detail figures.
- Write descriptions for the video clips that do not have one.
- Label the picture-only choices in the identification key.
- Settle the licence and credit wording with Jim Haney, Amanda McQuaid and UNH.
- Switch on visitor counting. GoatCounter is already wired in; it is suppressed while this
  review copy is built with `--noindex`, and starts when that option is removed.
- Write the Version 6.0 paragraph on the History of the Key page.
- Decide where the key will live permanently, and keep the old web addresses working.

## Who to ask

- **Dr. James F. Haney** (`jfhaney@unh.edu`) is the key's original author and principal
  investigator; questions about content, taxonomy and permissions go to him.
- **Amanda McQuaid** holds the project files and the OneDrive folder.
- **Jess Trout-Haney** did the 2026 rebuild and runs the build scripts.

## Credit

The key is the work of many people at the UNH Center for Freshwater Biology; see the
Contributors list on the home page and the History of the Key page. Images remain the
property of the organisations and individuals credited under Purpose and General
Information. Please read the Copyright page before reusing anything here.
