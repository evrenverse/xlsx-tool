# Compatibility

| Area | Supported |
| --- | --- |
| Go | 1.25 and newer |
| Operating systems | Linux, macOS, Windows |
| Architecture releases | amd64 and arm64 |
| Document format | `.xlsx` OOXML workbooks |
| Optional integration | LibreOffice for `validate` |

CI runs tests and builds on all three operating systems. The primary Linux job
also runs the race detector, coverage gate, agent eval, and a real LibreOffice
integration path.

The JSON agent contract is versioned separately from releases. Excelize may
normalize OOXML package details on save. Password-protected files, legacy
`.xls`, cloud spreadsheets, and perfect preservation of arbitrary vendor
extensions are not promised.
