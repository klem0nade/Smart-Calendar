# Smart Calendar — Schedule v2.70

Source snapshot of the Google Sheets **Smart Calendar** labeled **Schedule v2.70**, together with its container-bound Google Apps Script project.

## Repository layout

- `spreadsheet/Schedule-v2.70.xlsx` — Excel export of the Google Sheet, including formulas, formatting, and worksheets supported by the export format.
- `apps-script/Untitled.gs` — original Apps Script source file from the bound project.
- `apps-script/appsscript.json` — Apps Script manifest reflecting the project's visible settings.

## Restoring or updating the Apps Script

This public snapshot intentionally excludes the private Google Sheet ID and Apps Script project ID. To use Google's `clasp` command-line tool, create a local `.clasp.json` containing your own project binding and keep it out of source control.

The spreadsheet file is a point-in-time export. Future edits made in Google Sheets or Apps Script will not automatically appear in this repository.
