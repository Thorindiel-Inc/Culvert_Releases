# Culvert macro releases

Published Excel VBA modules that drive the MIDAS Civil NX API, and the
update feed for `midas-macro-updater.bas`.

**Generated — do not edit here.** These files are produced from a private
source repo by `scripts/make_manifest.py --publish`. Anything changed
directly in this repo is overwritten on the next publish.

## What the updater fetches

| path | purpose |
|---|---|
| `manifest.txt` | one `id\|version` line per module |
| `<id>.bas` | the full text of that module |

The updater compares each module's `SCRIPT_VERSION` against the manifest
and offers to replace whatever is out of date. Modules identify
themselves by a `SCRIPT_ID` constant rather than by their name in the VBA
project, because these files carry no `Attribute VB_Name` — the name in a
workbook is whatever the person pasting it typed.

## No credentials here

These modules contain **no MAPI key**. Each reads it at runtime from
`INPUT!K20` in the host workbook. That is deliberate: this repo is public,
and a MIDAS API key grants write and delete on whatever model Civil NX
currently has open.

The publisher refuses to write a tree containing anything key-shaped, so a
key cannot reach this repo by accident.

## Using them

1. Open the workbook's VBA editor (Alt+F11).
2. Paste a module into a standard module.
3. Put your MAPI key in `INPUT!K20` (Civil NX: `Tools > API > API Setting`).
4. Run `CheckMidasMacroUpdates()` from `macro-updater` to keep them current.

Installing updates needs *Trust access to the VBA project object model*
(File > Options > Trust Center > Trust Center Settings > Macro Settings).
It is off by default and set per machine. Without it the updater still
reports what is stale; it just cannot install.
