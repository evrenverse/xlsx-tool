# xlsx-tool

[![CI](https://github.com/evrenverse/xlsx-tool/actions/workflows/ci.yml/badge.svg)](https://github.com/evrenverse/xlsx-tool/actions/workflows/ci.yml)
[![CodeQL](https://github.com/evrenverse/xlsx-tool/actions/workflows/codeql.yml/badge.svg)](https://github.com/evrenverse/xlsx-tool/actions/workflows/codeql.yml)
[![OpenSSF Scorecard](https://api.scorecard.dev/projects/github.com/evrenverse/xlsx-tool/badge)](https://scorecard.dev/viewer/?uri=github.com/evrenverse/xlsx-tool)
[![Go Reference](https://pkg.go.dev/badge/github.com/evrenverse/xlsx-tool.svg)](https://pkg.go.dev/github.com/evrenverse/xlsx-tool)
[![License: MIT](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

An agent-friendly CLI for inspecting and editing `.xlsx` workbooks. It offers
small, composable commands, predictable exit codes, JSON output, bounded reads,
and crash-safe writes.

The repository is also a skills-only Codex plugin. Its portable Agent Skill is
under [`skills/xlsx-tool`](skills/xlsx-tool).

## Why agent-native

This is not a thin wrapper branded for AI. The executable exposes a versioned,
machine-readable contract; keeps document reads bounded; separates JSON stdout
from diagnostics; limits workbook expansion and batch input; publishes writes
atomically; and ships a synthetic locate-edit-verify eval that drives the real
CLI. An agent can discover capabilities without opening a user document:

```bash
xlsx-tool capabilities --json
xlsx-tool doctor --json
xlsx-tool schema batch
xlsx-tool version --json
```

## Give this repository to an agent

Copy this prompt:

> Inspect https://github.com/evrenverse/xlsx-tool and follow `AGENTS.md`.
> Install `xlsx-tool`, verify it with `xlsx-tool --version`, install the bundled
> `xlsx-tool` skill for this project, then use the CLI rather than writing a
> one-off spreadsheet script.

The instructions deliberately ask the agent to inspect the repository before
running anything.

## Install

### Go

Requires Go 1.25 or newer:

```bash
go install github.com/evrenverse/xlsx-tool/cmd/xlsx-tool@latest
xlsx-tool --version
```

### Release binary

Download the archive for your operating system from
[GitHub Releases](https://github.com/evrenverse/xlsx-tool/releases), verify it
against `checksums.txt`, and place `xlsx-tool` on `PATH`.

### From a clone

```bash
go install ./cmd/xlsx-tool
go test ./...
```

## Install the Agent Skill

For a project-local installation that works with Agent Skills compatible
clients:

```bash
mkdir -p .agents/skills
cp -R skills/xlsx-tool .agents/skills/xlsx-tool
```

Codex can also load the repository as a skills-only plugin because
`.codex-plugin/plugin.json` is included. Claude Code users can copy the same
folder to `.claude/skills/xlsx-tool`.

The skill does not hide a second implementation. It teaches an agent how to
install and use this CLI safely.

## Quick start

```bash
# Discover structure without dumping the entire workbook
xlsx-tool info report.xlsx --json
xlsx-tool sheets report.xlsx --json

# Locate and inspect a value
xlsx-tool find report.xlsx "Grand total" --json
xlsx-tool read report.xlsx Summary --cells B5,C5 --json

# Make one change or apply a JSON batch
xlsx-tool write report.xlsx Summary C5 "1250.00" --json
xlsx-tool batch report.xlsx changes.json --json

# Verify the result
xlsx-tool read report.xlsx Summary --cells C5 --json
```

Run `xlsx-tool <command> --help` for all options. Commands are:

| Command | Purpose |
| --- | --- |
| `sheets` | List sheet names |
| `info` | Summarize workbook and sheet dimensions |
| `find` | Locate values or formulas |
| `read` | Read a sheet, range, or bounded cell list |
| `write` | Set one value or formula |
| `batch` | Apply multiple writes from JSON |
| `validate` | Recalculate and compare a result to a baseline |
| `capabilities` | Print the versioned agent contract |
| `doctor` | Check runtime and optional dependencies |
| `schema` | List or print bundled JSON Schemas |
| `version` | Print tool and contract versions |

`validate` requires LibreOffice (`soffice` or `libreoffice`) on `PATH`.
LibreOffice runs with a new private profile for every invocation. Core
inspection and editing commands do not require LibreOffice.

## Safe automation contract

- JSON goes to stdout; diagnostics go to stderr.
- No-match searches return exit code `1`.
- Large unscoped reads require an explicit confirmation flag.
- Writes use a same-directory temporary file, `fsync`, and atomic rename.
- Formula recalculation never uses a shared LibreOffice profile or
  `LD_PRELOAD`.
- Back up important documents and inspect diffs before distributing them.

Excelize necessarily rewrites parts of the OOXML package when saving. The test
suite covers the workbook features used by this CLI, but no library can
promise lossless preservation of every vendor extension.

## Development

```bash
make install
make check
```

The quality gate includes formatting, vet, race tests, coverage, build, and the
real CLI agent eval. CI additionally tests Linux with and without LibreOffice,
known vulnerabilities, CodeQL, and dependency changes.

## Project documentation

- [Architecture](docs/architecture.md)
- [Agent contract](docs/agent-contract.md)
- [Threat model](docs/threat-model.md)
- [Compatibility](docs/compatibility.md)
- [Release process](docs/release-process.md)
- [Contributing](CONTRIBUTING.md), [security](SECURITY.md), and
  [support](SUPPORT.md)

## License

[MIT](LICENSE). Binary archives also include the dependency license materials
listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
