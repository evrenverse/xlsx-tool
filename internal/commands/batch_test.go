package commands

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

func TestBatchRejectsOversizedJSON(t *testing.T) {
	input := strings.NewReader(strings.Repeat(" ", maxBatchJSONBytes+1))
	var output bytes.Buffer
	err := runBatchFromReader("unused.xlsx", input, &output)
	if err == nil || !strings.Contains(err.Error(), "JSON input exceeds") {
		t.Fatalf("expected size-limit error, got %v", err)
	}
}

func TestBatchCommand(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	jsonPath := filepath.Join(t.TempDir(), "changes.json")
	jsonContent := `[
		{"sheet": "Sheet1", "cell": "A5", "value": "Gamma"},
		{"sheet": "Sheet1", "cell": "B5", "value": 30},
		{"sheet": "Data", "cell": "A3", "value": 2}
	]`
	os.WriteFile(jsonPath, []byte(jsonContent), 0644)

	var stdout bytes.Buffer
	err := runBatch(xlsxPath, jsonPath, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5")
	if val != "Gamma" {
		t.Errorf("expected 'Gamma', got %q", val)
	}

	val = testutil.ReadCellValue(t, xlsxPath, "Sheet1", "B5")
	if val != "30" {
		t.Errorf("expected '30', got %q", val)
	}

	val = testutil.ReadCellValue(t, xlsxPath, "Data", "A3")
	if val != "2" {
		t.Errorf("expected '2', got %q", val)
	}

	output := stdout.String()
	if !strings.Contains(output, "3 cells") {
		t.Errorf("expected '3 cells' in output, got:\n%s", output)
	}
}

func TestBatchCommand_WithFormula(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	jsonPath := filepath.Join(t.TempDir(), "changes.json")
	jsonContent := `[
		{"sheet": "Sheet1", "cell": "C1", "value": "Computed"},
		{"sheet": "Sheet1", "cell": "C2", "formula": "=A2&B2"}
	]`
	os.WriteFile(jsonPath, []byte(jsonContent), 0644)

	var stdout bytes.Buffer
	err := runBatch(xlsxPath, jsonPath, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "C2")
	if formula != "A2&B2" {
		t.Errorf("expected formula 'A2&B2', got %q", formula)
	}
}

func TestBatchCommand_PreservesExisting(t *testing.T) {
	xlsxPath := testutil.CreateFormattedFixture(t)

	jsonPath := filepath.Join(t.TempDir(), "changes.json")
	jsonContent := `[{"sheet": "Sheet1", "cell": "B3", "value": "new"}]`
	os.WriteFile(jsonPath, []byte(jsonContent), 0644)

	var stdout bytes.Buffer
	err := runBatch(xlsxPath, jsonPath, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	font := testutil.GetCellStyleFont(t, xlsxPath, "Sheet1", "A1")
	if font == nil || !font.Bold {
		t.Error("formatting on A1 was lost after batch write")
	}

	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A2")
	if val != "Data" {
		t.Errorf("existing cell corrupted: expected 'Data', got %q", val)
	}
}

func TestBatchCommand_InvalidJSON(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	jsonPath := filepath.Join(t.TempDir(), "bad.json")
	os.WriteFile(jsonPath, []byte(`[{"invalid json`), 0644)

	var stdout bytes.Buffer
	err := runBatch(xlsxPath, jsonPath, &stdout)
	if err == nil {
		t.Fatal("expected error for invalid JSON")
	}
}

func TestBatchCommand_SheetNotFound(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	jsonPath := filepath.Join(t.TempDir(), "changes.json")
	jsonContent := `[{"sheet": "NonExistent", "cell": "A1", "value": "test"}]`
	os.WriteFile(jsonPath, []byte(jsonContent), 0644)

	var stdout bytes.Buffer
	err := runBatch(xlsxPath, jsonPath, &stdout)
	if err == nil {
		t.Fatal("expected error for nonexistent sheet")
	}
}

func TestBatchCommand_LiteralEqualsString(t *testing.T) {
	// An explicit "value" field is NEVER interpreted as a formula — this is
	// the documented escape hatch for literal strings starting with "=".
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "=not a formula"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5")
	if val != "=not a formula" {
		t.Errorf("expected literal '=not a formula', got %q", val)
	}
	formula := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "A5")
	if formula != "" {
		t.Errorf("expected no formula for explicit value field, got %q", formula)
	}
}

func TestBatchCommand_Formula_OutsideDimension(t *testing.T) {
	// F10 is outside the fixture's stored dimension (A1:B4); batch must
	// still extend it so later full-sheet reads see the new cell.
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "F10", "formula": "=SUM(B2:B3)"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "F10")
	if formula != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)', got %q", formula)
	}

	// batch no longer recalcs — no cached value yet; validate owns that.
	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "F10")
	if val != "" {
		t.Errorf("expected no cached value before validate recalcs, got %q", val)
	}
}

func TestBatchCommand_MissingValueAndFormula(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[
		{"sheet": "Sheet1", "cell": "A5", "value": "ok"},
		{"sheet": "Sheet1", "cell": "B9"}
	]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err == nil {
		t.Fatal("expected error for change with neither value nor formula")
	}
	if !strings.Contains(err.Error(), "B9") || !strings.Contains(err.Error(), "1") {
		t.Errorf("expected error naming entry index and cell B9, got: %v", err)
	}

	// Nothing must have been applied, including the valid first entry
	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5")
	if val != "" {
		t.Errorf("no change may be applied when validation fails, got A5=%q", val)
	}
}

func TestBatchCommand_DegenerateFormula(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "formula": "="}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err == nil {
		t.Fatal("expected error for degenerate formula '='")
	}
	if !strings.Contains(err.Error(), "formula") {
		t.Errorf("expected 'formula' in error, got: %v", err)
	}
}

// --- Structured (--json) result tests ---

func TestBatchCommand_JSON(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[
		{"sheet": "Sheet1", "cell": "A5", "value": "Gamma"},
		{"sheet": "Sheet1", "cell": "C5", "formula": "=SUM(B2:B3)"}
	]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be pure valid JSON: %v\n%s", err, stdout.String())
	}

	written, ok := result["written"].([]any)
	if !ok || len(written) != 2 {
		t.Fatalf("expected 2 written entries, got %v", result["written"])
	}

	first := written[0].(map[string]any)
	if first["sheet"] != "Sheet1" || first["cell"] != "A5" || first["value"] != "Gamma" {
		t.Errorf("expected {Sheet1, A5, Gamma}, got %v", first)
	}

	second := written[1].(map[string]any)
	if second["formula"] != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)' (no leading =), got %v", second["formula"])
	}

	if _, ok := result["skipped_locked"].([]any); !ok {
		t.Errorf("expected 'skipped_locked' array, got %v", result["skipped_locked"])
	}
}

func TestBatchCommand_JSON_SkippedLocked(t *testing.T) {
	xlsxPath := testutil.CreateProtectedFixture(t)

	stdinContent := `[
		{"sheet": "Sheet1", "cell": "A2", "value": 99},
		{"sheet": "Sheet2", "cell": "A2", "value": "ok"}
	]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	skipped, ok := result["skipped_locked"].([]any)
	if !ok || len(skipped) != 1 {
		t.Fatalf("expected 1 skipped cell, got %v", result["skipped_locked"])
	}
	entry := skipped[0].(map[string]any)
	if entry["sheet"] != "Sheet1" || entry["cell"] != "A2" {
		t.Errorf("expected {Sheet1, A2}, got %v", entry)
	}

	written := result["written"].([]any)
	if len(written) != 1 {
		t.Fatalf("expected 1 written entry (unprotected Sheet2), got %d", len(written))
	}
	if written[0].(map[string]any)["sheet"] != "Sheet2" {
		t.Errorf("expected written entry on Sheet2, got %v", written[0])
	}
}

func TestBatchCommand_JSON_SkippedLocked_FormulaCell(t *testing.T) {
	xlsxPath := testutil.CreateProtectedFormulaFixture(t)

	// A2 is a locked formula cell (=A1*2); A1 is a plain locked cell.
	stdinContent := `[{"sheet": "Sheet1", "cell": "A2", "value": 99}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	skipped, ok := result["skipped_locked"].([]any)
	if !ok || len(skipped) != 1 {
		t.Fatalf("expected 1 skipped cell, got %v", result["skipped_locked"])
	}
	cell := skipped[0].(map[string]any)

	if _, ok := cell["formula"]; !ok {
		t.Errorf("expected formula field on skipped_locked formula cell, got %v", cell)
	}

	hint, _ := cell["hint"].(string)
	if !strings.Contains(hint, "computes automatically") {
		t.Errorf("expected instructive hint containing 'computes automatically', got %q", hint)
	}
}

func TestBatchCommand_JSON_SkippedLocked_PlainCell(t *testing.T) {
	xlsxPath := testutil.CreateProtectedFormulaFixture(t)

	// A1 is a plain locked (non-formula) cell.
	stdinContent := `[{"sheet": "Sheet1", "cell": "A1", "value": 99}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	skipped, ok := result["skipped_locked"].([]any)
	if !ok || len(skipped) != 1 {
		t.Fatalf("expected 1 skipped cell, got %v", result["skipped_locked"])
	}
	cell := skipped[0].(map[string]any)

	if _, ok := cell["formula"]; ok {
		t.Errorf("expected no formula field on plain locked cell, got %v", cell)
	}

	hint, _ := cell["hint"].(string)
	if !strings.Contains(hint, "do not treat as a blocker") {
		t.Errorf("expected hint containing 'do not treat as a blocker', got %q", hint)
	}
}

func TestBatchCommand_Stdin(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "FromStdin"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5")
	if val != "FromStdin" {
		t.Errorf("expected 'FromStdin', got %q", val)
	}
}

func TestBatchCommand_InvalidCellRef(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[
		{"sheet": "Sheet1", "cell": "A5", "value": "ok"},
		{"sheet": "Sheet1", "cell": "9Z", "value": 1}
	]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err == nil {
		t.Fatal("expected error for invalid cell reference")
	}
	if !strings.Contains(err.Error(), "9Z") {
		t.Errorf("expected error naming the offending ref, got: %v", err)
	}

	// Nothing applied
	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "" {
		t.Errorf("no change may be applied when validation fails, got A5=%q", val)
	}
}

func TestBatchCommand_NonScalarValue(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	for _, payload := range []string{
		`[{"sheet": "Sheet1", "cell": "A5", "value": {"nested": 1}}]`,
		`[{"sheet": "Sheet1", "cell": "A5", "value": [1, 2]}]`,
	} {
		stdinReader := strings.NewReader(payload)
		var stdout bytes.Buffer
		err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
		if err == nil {
			t.Errorf("expected error for non-scalar value %s", payload)
			continue
		}
		if !strings.Contains(err.Error(), "scalar") {
			t.Errorf("expected 'scalar' in error, got: %v", err)
		}
	}

	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "" {
		t.Errorf("non-scalar value must not be written, got %q", val)
	}
}

func TestBatchCommand_BothValueAndFormula(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "x", "formula": "=B2"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err == nil {
		t.Fatal("expected error for entry with BOTH value and formula")
	}
	if !strings.Contains(err.Error(), "not both") {
		t.Errorf("expected 'not both' in error, got: %v", err)
	}
}

func TestBatchCommand_TrailingGarbage(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "ok"}] trailing garbage`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReader(xlsxPath, stdinReader, &stdout)
	if err == nil {
		t.Fatal("expected error for trailing data after the JSON array")
	}
	if !strings.Contains(err.Error(), "trailing") {
		t.Errorf("expected 'trailing' in error, got: %v", err)
	}

	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "" {
		t.Errorf("nothing may be written on trailing-garbage input, got %q", val)
	}
}
