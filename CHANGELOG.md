# Changelog

All notable changes are documented here. This project follows Semantic
Versioning and keeps an [Unreleased] section.

## [Unreleased]

## [0.1.0] - 2026-07-24

- Prepare the independent public release.
- Add an Agent Skill and Codex plugin manifest.
- Isolate LibreOffice recalculation behind a private temporary profile.
- Update Excelize and transitive dependencies.
- Add a versioned, machine-readable Agent Contract with capabilities, doctor,
  version, and bundled JSON Schema commands.
- Add a `recalc` command that refreshes formula caches in place and reports
  only whether recalculation succeeded, leaving content judgement to
  `validate`.
- Add explicit workbook expansion and batch limits, fuzz tests, race-tested
  coverage enforcement, and a real CLI locate-edit-verify eval.
- Add CI, CodeQL, dependency review, OpenSSF Scorecard, SBOMs, and provenance
  attestations with commit-pinned GitHub Actions.
- Add architecture, threat-model, compatibility, release, support, governance,
  and public-publishing documentation.

[Unreleased]: https://github.com/evrenverse/xlsx-tool/compare/v0.1.0...HEAD
[0.1.0]: https://github.com/evrenverse/xlsx-tool/releases/tag/v0.1.0
