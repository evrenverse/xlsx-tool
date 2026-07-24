# Contributing

Issues and focused pull requests are welcome.

## Development

1. Install Go 1.25 or newer. `make check` also needs LibreOffice (`soffice` or
   `libreoffice` on `PATH`, package `libreoffice-calc`); the recalculation
   tests skip without it and the coverage gate then falls below its floor.
2. Create a branch and add tests for behavior changes.
3. Run the same local gate used by CI:

   ```bash
   make install
   make check
   make audit
   ```

4. Update `README.md`, the bundled skill, and `CHANGELOG.md` when user-facing
   behavior changes.
5. If runtime dependencies change, regenerate and review
   `third_party_licenses/` as described in `THIRD_PARTY_NOTICES.md`.

Preserve the CLI contract: stable JSON fields, bounded output, useful exit
codes, diagnostics on stderr, and atomic writes. Fixtures must be synthetic or
redistributable and must not contain personal, customer, or company data.

## Pull requests

Keep changes focused and explain the user or agent workflow they improve. Add
tests for behavior, error paths, and limits. A contract change also needs
updated schemas and eval coverage. Prefer conventional commit subjects such as
`feat:`, `fix:`, `docs:`, `test:`, or `chore:`.

By participating, you agree to the [Code of Conduct](CODE_OF_CONDUCT.md).
