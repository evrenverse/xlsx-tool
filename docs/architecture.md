# Architecture

`xlsx-tool` is a local, single-binary Go CLI. It has no server component,
telemetry, or runtime network dependency.

```text
agent or developer
        |
        v
Cobra command layer  -----> JSON stdout / diagnostics stderr
        |
        v
validation + bounded addressing
        |
        v
Excelize workbook model
        |
        v
same-directory temporary file -> reopen/check -> fsync -> atomic rename
```

`cmd/xlsx-tool` owns process startup and version injection. `internal/commands`
owns the public command contract, validation, resource limits, and atomic
publication. `internal/format` owns stable output rendering. `evals` exercises
the compiled CLI as an agent would; unit and fuzz tests cover internal behavior.

Only `recalc` and `validate` cross a process boundary. They invoke a locally
installed LibreOffice executable with a fresh temporary user profile. The
workbook is copied into that workspace, recalculated, and published only after
success; `validate` then checks the result against a baseline.

## Design constraints

- One tool and one document format.
- Machine-readable discovery before document access.
- Bounded reads and input sizes.
- Stable JSON; human diagnostics never corrupt stdout.
- Mutations are crash-safe and independently verifiable.
- No organization-specific paths, services, or sibling-tool dependencies.
