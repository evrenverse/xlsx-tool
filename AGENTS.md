# Agent setup

This repository is an independent public Go CLI. Do not install unrelated
document tools when the task only concerns XLSX files.

## Install for use

1. Inspect `README.md`, `SECURITY.md`, and `go.mod`.
2. Prefer a signed/tagged release. If building is appropriate, run:

   ```bash
   go install github.com/evrenverse/xlsx-tool/cmd/xlsx-tool@latest
   xlsx-tool --version
   ```

3. For a source checkout, run `make check` before using the build.
4. Install the project skill only when the user wants it:

   ```bash
   mkdir -p .agents/skills
   cp -R skills/xlsx-tool .agents/skills/xlsx-tool
   ```

5. Run `xlsx-tool capabilities --json` and `xlsx-tool doctor --json` before
   choosing a workflow. Use `xlsx-tool schema <name>` instead of guessing batch
   structures.
6. Use scoped reads and JSON output. Make a backup, edit, then read changed
   cells back.

## Contributing

Keep stdout machine-readable, stderr diagnostic, and file writes atomic.
Do not add telemetry, remote services, secrets, organization-specific
workflows, or runtime dependencies on sibling repositories.

Run `make check` before submitting changes. User-visible contract changes must
update schemas, tests, the eval, README, skill, and changelog together.
