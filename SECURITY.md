# Security policy

## Supported versions

Security fixes are provided for the latest released version.

## Report a vulnerability

Use GitHub's private vulnerability reporting for this repository. Do not open
a public issue containing exploit details or sensitive documents. Include the
affected version, impact, reproduction steps, and any suggested mitigation.

## Trust model

`xlsx-tool` processes untrusted ZIP/XML-based office files. Run it with the
least filesystem access practical and keep dependencies current. The CLI does
not use telemetry or network access at runtime.

Most writes modify the input path. They are crash-safe, but they are not a
substitute for a backup. `validate` invokes the locally installed LibreOffice
binary on a private temporary copy and then replaces the result atomically.
LibreOffice remains part of that command's trusted computing base.

Never pass secrets as spreadsheet data unless the output file is meant to
contain them. Review generated workbooks before sharing.
