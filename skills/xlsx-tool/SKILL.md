---
name: xlsx-tool
description: Inspect, search, read, edit, batch-update, and validate Microsoft Excel XLSX workbooks with the xlsx-tool CLI. Use when an agent must understand or modify an .xlsx file while keeping reads bounded, producing machine-readable JSON, preserving workbook structure, and verifying every write.
---

# XLSX Tool

Use `xlsx-tool` instead of writing a one-off spreadsheet script.

## Prepare

1. Run `command -v xlsx-tool`.
2. If missing and this repository is available, run
   `go install ./cmd/xlsx-tool`. Otherwise run
   `go install github.com/evrenverse/xlsx-tool/cmd/xlsx-tool@latest`.
3. Run `xlsx-tool --version`.
4. Run `xlsx-tool capabilities --json` and `xlsx-tool doctor --json`.
5. Use `xlsx-tool schema <name>` when constructing structured input.
6. Work on a copy unless the user explicitly wants an in-place edit.

## Inspect before editing

1. Run `xlsx-tool info <file> --json`.
2. Use `find --json`, a bounded `read --range`, or `read --cells`.
3. Avoid dumping complete large sheets into context.

## Edit and verify

1. Prefer one `batch` operation for multiple changes.
2. Treat a leading `=` as a formula.
3. Read every changed cell back with `--json`.
4. If a baseline is available and LibreOffice is installed, run `validate`.
5. Report the output path and any validation limitation.

Never overwrite the only copy of an important workbook. Do not claim complete
format preservation without checking the resulting file.

Read [references/cli.md](references/cli.md) when command syntax or exit behavior
is needed.
