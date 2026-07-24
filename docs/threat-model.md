# Threat model

## Assets

- Workbook confidentiality and integrity.
- Files accessible to the invoking user.
- Predictable agent output and host availability.

## Trust boundaries

The input workbook is untrusted. Excelize and the Go ZIP/XML stack are inside
the parser trust boundary. The local filesystem is trusted only to the
permissions of the current process. LibreOffice joins the trusted computing
base only for `validate`.

## Main threats and controls

| Threat | Control |
| --- | --- |
| ZIP/XML resource exhaustion | Compressed and expanded-size limits; bounded JSON batches and reads |
| Partial or corrupt output | Same-directory temporary file, reopen check, sync, atomic rename |
| Context flooding | Scoped read commands and explicit confirmation for large output |
| Shared LibreOffice state or macro side effects | Private temporary profile and isolated copy |
| Dependency compromise | Lockfile, Dependabot, dependency review, CodeQL, govulncheck, pinned Actions |
| Release substitution | Checksums, SBOM, GitHub artifact provenance attestations |
| Accidental data disclosure | No telemetry/network calls; synthetic fixtures and issue policy |

## Out of scope

The tool is not a malware sandbox, a formula engine, or a guarantee that every
vendor-specific OOXML extension survives rewriting. Run hostile documents in
an operating-system sandbox and visually review high-value output.

Report vulnerabilities through [SECURITY.md](../SECURITY.md).
