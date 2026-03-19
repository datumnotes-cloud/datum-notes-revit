# Datum Notes pyRevit Extension

This extension adds a **Datum Notes** tab with a **Project Tools** panel and a **Meeting Notes** button.

## What it does

- Reads rooms from the active Revit model and groups them by level.
- Lets you add timestamped room notes.
- Tracks checklist completion for notes in this session.
- Persists notes to a project JSON file next to the Revit model.
- Exports a print-ready HTML checklist grouped by room.
- Provides a link to learn about uploading to Datum Notes.

## Folder output

- `DatumNotes.extension/`
- `Datum Notes.tab/Project Tools.panel/Meeting Notes.pushbutton/script.py`
- `Datum Notes.tab/Project Tools.panel/Meeting Notes.pushbutton/notes_ui.xaml`
- `Datum Notes.tab/Project Tools.panel/Meeting Notes.pushbutton/bundle.yaml`

## Install in pyRevit

1. Put `DatumNotes.extension` under your pyRevit extensions folder.
2. Run `pyrevit reload` (or reload from pyRevit settings UI).
3. Open Revit and click **Datum Notes > Project Tools > Meeting Notes**.

## Notes

- Designed for IronPython (Python 2.7 compatible).
- No external dependencies required.
