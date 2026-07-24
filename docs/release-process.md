# Release process

Releases are automated from signed or annotated `vMAJOR.MINOR.PATCH` tags.

1. Make `main` green. Promote the Unreleased changelog section to the new
   version with its release date, and point the `[Unreleased]` link at
   `compare/vPREVIOUS...HEAD` once a tag exists.
2. Run `make release-check`.
3. Create and push the release tag.
4. GitHub Actions reruns checks and GoReleaser builds Linux, macOS, and Windows
   archives for amd64 and arm64.
5. The workflow publishes checksums, an SPDX SBOM, generated release notes,
   and GitHub/Sigstore provenance and SBOM attestations.
6. Verify an archive with `checksums.txt` and
   `gh attestation verify <archive> --repo evrenverse/xlsx-tool`.

Never reuse or move a published tag. Fix a bad release with a new patch version.
Repository rules should require CI, CodeQL, dependency review, linear history,
and review of workflow or CODEOWNERS changes.
