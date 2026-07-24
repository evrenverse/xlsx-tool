# Agent contract

Contract version `1` makes the CLI self-describing:

```bash
xlsx-tool capabilities --json
xlsx-tool doctor --json
xlsx-tool schema
xlsx-tool schema batch
xlsx-tool version --json
```

`capabilities` reports commands, mutation behavior, addressing conventions,
optional dependencies, safety properties, schemas, and resource limits.
`doctor` performs read-only runtime checks and can require LibreOffice with
`--require libreoffice`. `schema` prints bundled JSON Schema documents without
network access.

## Stable conventions

- Contract and schema versions are independent from the binary version.
- JSON data is written to stdout; diagnostics are written to stderr.
- Workbook sheet names are explicit and cell addresses use Excel A1 notation.
- Success exits `0`. Search with no match and other command failures exit `1`;
  diagnostics explain the category on stderr.
- Agents should inspect, make a backup, perform a bounded operation, and read
  every changed cell back.

Backward-incompatible contract changes require a new contract major version.
Additive JSON fields and new schemas are allowed within version `1`.

The executable schemas are embedded from `internal/commands/schemas`. The
portable runtime guidance is in `skills/xlsx-tool`.
