# Agent evals

The eval executes the same bounded workflow an autonomous agent should use:

1. discover workbook structure;
2. locate a label without dumping the workbook;
3. read the target cell;
4. mutate it through the CLI;
5. read it again and verify the stored value.

All input workbooks are generated synthetically in a temporary directory.
No customer or employer documents are used.

```bash
go build -o /tmp/xlsx-tool-eval ./cmd/xlsx-tool
XLSX_TOOL_BIN=/tmp/xlsx-tool-eval go run ./evals
```

The runner writes one JSON result to stdout and exits nonzero on a failed
assertion.
