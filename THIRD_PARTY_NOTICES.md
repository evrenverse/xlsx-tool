# Third-party notices

`xlsx-tool` is statically linked with open-source Go dependencies. Their
required license and copyright materials are included under
`third_party_licenses/` and are shipped in every binary archive.

The set is generated from `go.mod` for Linux with
[`google/go-licenses`](https://github.com/google/go-licenses):

```bash
go-licenses save ./cmd/xlsx-tool \
  --save_path=<target> \
  --ignore github.com/evrenverse/xlsx-tool
go-licenses check ./cmd/xlsx-tool \
  --ignore github.com/evrenverse/xlsx-tool
```

Regenerate and review the directory whenever runtime dependencies change.
The SPDX SBOM attached to each release provides the exact version inventory;
the copied license texts are the redistribution materials, not a substitute
for that machine-readable inventory.
