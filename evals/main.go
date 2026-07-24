package main

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

type evalResult struct {
	SchemaVersion string `json:"schema_version"`
	Tool          string `json:"tool"`
	Eval          string `json:"eval"`
	Passed        bool   `json:"passed"`
	Commands      int    `json:"commands"`
	DurationMS    int64  `json:"duration_ms"`
}

func main() {
	start := time.Now()
	binary := os.Getenv("XLSX_TOOL_BIN")
	if binary == "" {
		binary = "xlsx-tool"
	}
	if _, err := exec.LookPath(binary); err != nil {
		fatalf("xlsx-tool binary not found; set XLSX_TOOL_BIN: %v", err)
	}

	root, err := os.MkdirTemp("", "xlsx-tool-eval-*")
	if err != nil {
		fatalf("create temp directory: %v", err)
	}
	defer os.RemoveAll(root)
	workbook := filepath.Join(root, "report.xlsx")
	createFixture(workbook)

	commands := [][]string{
		{"info", workbook, "--json"},
		{"find", workbook, "Grand total", "--json"},
		{"read", workbook, "Summary", "--cells", "B1", "--json"},
		{"write", workbook, "Summary", "B1", "1250", "--json"},
		{"read", workbook, "Summary", "--cells", "B1", "--json"},
	}
	var final []byte
	for _, args := range commands {
		output, runErr := exec.Command(binary, args...).CombinedOutput()
		if runErr != nil {
			fatalf("%s failed: %v\n%s", strings.Join(args, " "), runErr, output)
		}
		final = output
	}
	var payload struct {
		Cells map[string]struct {
			Value any `json:"value"`
		} `json:"cells"`
	}
	if err := json.Unmarshal(final, &payload); err != nil {
		fatalf("decode final read: %v\n%s", err, final)
	}
	cell, ok := payload.Cells["B1"]
	if !ok || fmt.Sprint(cell.Value) != "1250" {
		fatalf("final B1 value is %v, want 1250", cell.Value)
	}

	result := evalResult{
		SchemaVersion: "1",
		Tool:          "xlsx-tool",
		Eval:          "locate-edit-verify",
		Passed:        true,
		Commands:      len(commands),
		DurationMS:    time.Since(start).Milliseconds(),
	}
	encoder := json.NewEncoder(os.Stdout)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(result); err != nil {
		fatalf("encode result: %v", err)
	}
}

func createFixture(path string) {
	workbook := excelize.NewFile()
	defer workbook.Close()
	if err := workbook.SetSheetName("Sheet1", "Summary"); err != nil {
		fatalf("rename sheet: %v", err)
	}
	if err := workbook.SetCellValue("Summary", "A1", "Grand total"); err != nil {
		fatalf("write label: %v", err)
	}
	if err := workbook.SetCellValue("Summary", "B1", 1000); err != nil {
		fatalf("write value: %v", err)
	}
	if err := workbook.SaveAs(path); err != nil {
		fatalf("save fixture: %v", err)
	}
}

func fatalf(format string, args ...any) {
	fmt.Fprintf(os.Stderr, "eval failed: "+format+"\n", args...)
	os.Exit(1)
}
