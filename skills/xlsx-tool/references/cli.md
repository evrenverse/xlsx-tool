# CLI reference

Use `xlsx-tool <command> --help` as the authoritative option reference.

## Agent discovery

```bash
xlsx-tool capabilities --json
xlsx-tool doctor --json
xlsx-tool doctor --json --require libreoffice
xlsx-tool schema
xlsx-tool schema batch
xlsx-tool version --json
```

Use the bundled schema as the authoritative batch shape. The contract reports
hard input limits and whether each command mutates files.

## Discovery and reading

```bash
xlsx-tool sheets workbook.xlsx --json
xlsx-tool info workbook.xlsx --json
xlsx-tool find workbook.xlsx "Total" --json
xlsx-tool read workbook.xlsx Summary --range A1:D20 --json
xlsx-tool read workbook.xlsx Summary --cells B5,C5 --json
```

`find` returns exit code 1 when nothing matches. Large multi-sheet reads require
`--confirm-large`.

## Editing

```bash
xlsx-tool write workbook.xlsx Summary C5 "1250.00" --json
xlsx-tool write workbook.xlsx Summary C6 "=SUM(C2:C5)" --json
xlsx-tool batch workbook.xlsx changes.json --json
```

Inspect `xlsx-tool batch --help` and `xlsx-tool schema batch` before
constructing the JSON payload. Writes replace the target atomically after a
successful save.

## Validation

```bash
xlsx-tool validate result.xlsx --baseline original.xlsx --json
```

Validation requires LibreOffice on `PATH`, recalculates the result in place,
and compares formulas and fillable cells to the baseline.
