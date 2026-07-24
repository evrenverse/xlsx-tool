package commands

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

// --- Plain text (default) tests ---

func TestReadCommand_Plain_FullSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	// Should have header line
	if !strings.Contains(output, "Sheet: Sheet1") {
		t.Errorf("expected 'Sheet: Sheet1' header, got:\n%s", output)
	}

	// Should contain cell values
	if !strings.Contains(output, `"Name"`) {
		t.Errorf("expected '\"Name\"' in output, got:\n%s", output)
	}
	if !strings.Contains(output, `"Alpha"`) {
		t.Errorf("expected '\"Alpha\"' in output, got:\n%s", output)
	}

	// Numbers should be without quotes
	if !strings.Contains(output, "10") {
		t.Errorf("expected '10' in output, got:\n%s", output)
	}

	// Formulas should have = prefix
	if !strings.Contains(output, "=SUM(B2:B3)") {
		t.Errorf("expected '=SUM(B2:B3)' in output, got:\n%s", output)
	}

	// Should NOT be valid JSON
	var js any
	if json.Unmarshal(stdout.Bytes(), &js) == nil {
		t.Error("plain text output should NOT be valid JSON")
	}
}

func TestReadCommand_MissingSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{}, &stdout)
	if err == nil {
		t.Fatal("expected error when sheet name is empty")
	}
	if !strings.Contains(err.Error(), "not found") {
		t.Errorf("expected 'not found' in error, got: %v", err)
	}
}

func TestReadCommand_Plain_Range(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cellRange: "A1:B2"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `"Alpha"`) {
		t.Errorf("expected 'Alpha' in output")
	}
	if strings.Contains(output, `"Beta"`) {
		t.Errorf("should not contain 'Beta' (outside range)")
	}
}

func TestReadCommand_Plain_SingleCell(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A2"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if output != `"Alpha"` {
		t.Errorf("expected '\"Alpha\"', got %q", output)
	}
}

func TestReadCommand_Plain_Formula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "B4"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if output != "=SUM(B2:B3)" {
		t.Errorf("expected '=SUM(B2:B3)', got %q", output)
	}
}

func TestReadCommand_Plain_FillColor(t *testing.T) {
	path := testutil.CreateFillFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	// Row 2: empty cells with yellow fill
	if !strings.Contains(output, "[fill=#FFFF00]") {
		t.Errorf("expected '[fill=#FFFF00]' for empty yellow cells, got:\n%s", output)
	}

	// Row 3: value with green fill
	if !strings.Contains(output, `"Fest" [fill=#00FF00]`) {
		t.Errorf("expected '\"Fest\" [fill=#00FF00]' for green cell, got:\n%s", output)
	}

	// Row 3: formula without fill
	if !strings.Contains(output, "=A3") {
		t.Errorf("expected '=A3' for formula cell, got:\n%s", output)
	}
}

func TestReadCommand_Plain_FillColor_SingleCell(t *testing.T) {
	path := testutil.CreateFillFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A2"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if output != "[fill=#FFFF00]" {
		t.Errorf("expected '[fill=#FFFF00]', got %q", output)
	}
}

func TestReadCommand_Plain_RowNumbers(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	// Row numbers should be in brackets
	if !strings.Contains(output, "[1]") {
		t.Errorf("expected '[1]' row number, got:\n%s", output)
	}
	if !strings.Contains(output, "[4]") {
		t.Errorf("expected '[4]' row number, got:\n%s", output)
	}
}

func TestReadCommand_Plain_EmptyCell(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "C1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if output != "(empty)" {
		t.Errorf("expected '(empty)' for empty cell, got %q", output)
	}
}

// --- JSON (--json flag) tests ---

func TestReadCommand_JSON_FullSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if result["sheet"] != "Sheet1" {
		t.Errorf("expected sheet 'Sheet1', got %v", result["sheet"])
	}

	rows, ok := result["rows"].([]any)
	if !ok || len(rows) == 0 {
		t.Fatalf("expected rows array, got %v", result["rows"])
	}

	// First row should contain "Name"
	firstRow := rows[0].(map[string]any)
	cells := firstRow["cells"].(map[string]any)
	aCell := cells["A"].(map[string]any)
	if aCell["value"] != "Name" {
		t.Errorf("expected A1='Name', got %v", aCell["value"])
	}
}

func TestReadCommand_JSON_SingleCell(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A2", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v", err)
	}

	if result["value"] != "Alpha" {
		t.Errorf("expected value 'Alpha', got %v", result["value"])
	}
}

func TestReadCommand_JSON_Formula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "B4", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v", err)
	}

	if result["formula"] != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)', got %v", result["formula"])
	}
	if _, hasValue := result["value"]; hasValue {
		t.Errorf("formula cells should not have 'value' key")
	}
}

func TestReadCommand_JSON_RangeFormula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cellRange: "A4:B4", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `"formula"`) {
		t.Errorf("expected 'formula' key for formula cells, got:\n%s", output)
	}
}

func TestReadCommand_JSON_FillColor(t *testing.T) {
	path := testutil.CreateFillFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	rows := result["rows"].([]any)

	// Row 1: headers without fill
	row1 := rows[0].(map[string]any)["cells"].(map[string]any)
	a1 := row1["A"].(map[string]any)
	if _, hasFill := a1["fill"]; hasFill {
		t.Errorf("A1 (header) should not have fill, got %v", a1["fill"])
	}

	// Row 2: empty cells with yellow fill
	row2 := rows[1].(map[string]any)["cells"].(map[string]any)
	a2 := row2["A"].(map[string]any)
	if a2["fill"] != "#FFFF00" {
		t.Errorf("A2 (empty+yellow) expected fill '#FFFF00', got %v", a2["fill"])
	}
	if a2["value"] != nil {
		t.Errorf("A2 should have null value, got %v", a2["value"])
	}

	// Row 3: value with green fill
	row3 := rows[2].(map[string]any)["cells"].(map[string]any)
	a3 := row3["A"].(map[string]any)
	if a3["fill"] != "#00FF00" {
		t.Errorf("A3 (value+green) expected fill '#00FF00', got %v", a3["fill"])
	}
	if a3["value"] != "Fest" {
		t.Errorf("A3 expected value 'Fest', got %v", a3["value"])
	}

	// B3: formula without fill
	b3 := row3["B"].(map[string]any)
	if b3["formula"] != "A3" {
		t.Errorf("B3 expected formula 'A3', got %v", b3["formula"])
	}
	if _, hasFill := b3["fill"]; hasFill {
		t.Errorf("B3 (formula, no fill) should not have fill")
	}
}

func TestReadCommand_JSON_FillColor_SingleCell(t *testing.T) {
	path := testutil.CreateFillFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A2", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v", err)
	}

	if result["fill"] != "#FFFF00" {
		t.Errorf("A2 single cell expected fill '#FFFF00', got %v", result["fill"])
	}
}

func TestReadCommand_SheetNotFound(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "NonExistent", readOptions{}, &stdout)
	if err == nil {
		t.Fatal("expected error for nonexistent sheet")
	}
	if !strings.Contains(err.Error(), "not found") {
		t.Errorf("expected 'not found' in error, got: %v", err)
	}
}

// --- Scattered cells (--cells flag) tests ---

func TestReadCommand_Plain_Cells(t *testing.T) {
	// B4 (=SUM(B2:B3)) needs a cached computed value — batch/write no longer
	// recalc, so the fixture carries a pre-patched cache (see validate,
	// which recalcs via LibreOffice).
	path := testutil.CreateSimpleFixtureWithCachedTotal(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "B4,A2,C1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	want := "B4: =SUM(B2:B3) → 30\n" +
		"A2: \"Alpha\"\n" +
		"C1: (empty)\n"
	if stdout.String() != want {
		t.Errorf("expected output:\n%s\ngot:\n%s", want, stdout.String())
	}
}

func TestReadCommand_Plain_Cells_OrderPreserved(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A3,A1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	want := "A3: \"Beta\"\n" +
		"A1: \"Name\"\n"
	if stdout.String() != want {
		t.Errorf("expected input order preserved:\n%s\ngot:\n%s", want, stdout.String())
	}
}

func TestReadCommand_Plain_Cells_Whitespace(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: " A2 ,  B2 "}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	want := "A2: \"Alpha\"\n" +
		"B2: 10\n"
	if stdout.String() != want {
		t.Errorf("expected whitespace-tolerant parsing:\n%s\ngot:\n%s", want, stdout.String())
	}
}

func TestReadCommand_JSON_Cells(t *testing.T) {
	// B4 (=SUM(B2:B3)) needs a cached computed value — batch/write no longer
	// recalc, so the fixture carries a pre-patched cache (see validate,
	// which recalcs via LibreOffice).
	path := testutil.CreateSimpleFixtureWithCachedTotal(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "B4,A2,C1", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if result["sheet"] != "Sheet1" {
		t.Errorf("expected sheet 'Sheet1', got %v", result["sheet"])
	}

	cells, ok := result["cells"].(map[string]any)
	if !ok {
		t.Fatalf("expected 'cells' object, got %v", result["cells"])
	}
	if len(cells) != 3 {
		t.Errorf("expected 3 cells, got %d: %v", len(cells), cells)
	}

	b4, ok := cells["B4"].(map[string]any)
	if !ok {
		t.Fatalf("expected B4 cell object, got %v", cells["B4"])
	}
	if b4["formula"] != "SUM(B2:B3)" {
		t.Errorf("expected B4 formula 'SUM(B2:B3)', got %v", b4["formula"])
	}
	// Cached numeric formula results are now emitted as a typed JSON number.
	if num, ok := b4["value"].(float64); !ok || num != 30 {
		t.Errorf("expected B4 cached value 30 (number), got %T (%v)", b4["value"], b4["value"])
	}

	a2, ok := cells["A2"].(map[string]any)
	if !ok {
		t.Fatalf("expected A2 cell object, got %v", cells["A2"])
	}
	if a2["value"] != "Alpha" {
		t.Errorf("expected A2 value 'Alpha', got %v", a2["value"])
	}

	c1, present := cells["C1"]
	if !present {
		t.Error("expected C1 key to be present")
	}
	if c1 != nil {
		t.Errorf("expected C1 to be null (empty cell), got %v", c1)
	}
}

func TestReadCommand_Cells_InvalidRef(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A1,12,B2"}, &stdout)
	if err == nil {
		t.Fatal("expected error for invalid cell reference")
	}
	if !strings.Contains(err.Error(), `"12"`) {
		t.Errorf("expected error to name the offending ref '12', got: %v", err)
	}
}

func TestReadCommand_Cells_EmptyList(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: " , "}, &stdout)
	if err == nil {
		t.Fatal("expected error for empty cell list")
	}
	if !strings.Contains(err.Error(), "no cell references") {
		t.Errorf("expected 'no cell references' in error, got: %v", err)
	}
}

func TestReadCommand_Cells_MutuallyExclusiveWithCell(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A1,B2", cell: "A1"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --cells with --cell")
	}
	if !strings.Contains(err.Error(), "--cells cannot be combined") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

func TestReadCommand_Cell_MutuallyExclusiveWithRange(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A1", cellRange: "A1:B2"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --cell with --range")
	}
	if !strings.Contains(err.Error(), "--cell cannot be combined with --range") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

func TestReadCommand_Cells_MutuallyExclusiveWithRange(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A1,B2", cellRange: "A1:B2"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --cells with --range")
	}
	if !strings.Contains(err.Error(), "--cells cannot be combined") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

func TestReadCommand_Plain_Cells_DedupRefs(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A2,A2,B2,A2"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	want := "A2: \"Alpha\"\n" +
		"B2: 10\n"
	if stdout.String() != want {
		t.Errorf("expected duplicate refs collapsed to first occurrence:\n%s\ngot:\n%s",
			want, stdout.String())
	}
}

func TestReadCommand_JSON_Cells_DedupRefs(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cells: "A2,A2,B2", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	cells, ok := result["cells"].(map[string]any)
	if !ok {
		t.Fatalf("expected 'cells' object, got %v", result["cells"])
	}
	if len(cells) != 2 {
		t.Errorf("expected 2 cells after dedup, got %d: %v", len(cells), cells)
	}
}

// --- Multi-sheet (--sheets / --all-sheets) tests ---

func TestReadCommand_Plain_Sheets(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1,Data"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	idxSheet1 := strings.Index(output, "Sheet: Sheet1")
	idxData := strings.Index(output, "Sheet: Data")
	if idxSheet1 == -1 || idxData == -1 {
		t.Fatalf("expected both sheet headers, got:\n%s", output)
	}
	if idxSheet1 > idxData {
		t.Errorf("expected requested order (Sheet1 before Data), got:\n%s", output)
	}

	// Blocks separated by a blank line
	if !strings.Contains(output, "\n\nSheet: Data") {
		t.Errorf("expected blank line between sheet blocks, got:\n%s", output)
	}

	// Content from both sheets
	if !strings.Contains(output, `"Alpha"`) {
		t.Errorf("expected Sheet1 content 'Alpha', got:\n%s", output)
	}
	if !strings.Contains(output, "99.5") {
		t.Errorf("expected Data content '99.5', got:\n%s", output)
	}
}

func TestReadCommand_Plain_Sheets_RequestedOrder(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Data,Sheet1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	idxData := strings.Index(output, "Sheet: Data")
	idxSheet1 := strings.Index(output, "Sheet: Sheet1")
	if idxData == -1 || idxSheet1 == -1 {
		t.Fatalf("expected both sheet headers, got:\n%s", output)
	}
	if idxData > idxSheet1 {
		t.Errorf("expected requested order (Data before Sheet1), got:\n%s", output)
	}
}

func TestReadCommand_JSON_Sheets(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1,Data", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be a single valid JSON document: %v\n%s", err, stdout.String())
	}

	sheets, ok := result["sheets"].([]any)
	if !ok {
		t.Fatalf("expected 'sheets' array, got %v", result["sheets"])
	}
	if len(sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(sheets))
	}

	first := sheets[0].(map[string]any)
	second := sheets[1].(map[string]any)
	if first["sheet"] != "Sheet1" {
		t.Errorf("expected first sheet 'Sheet1', got %v", first["sheet"])
	}
	if second["sheet"] != "Data" {
		t.Errorf("expected second sheet 'Data', got %v", second["sheet"])
	}

	// Each entry is a full SheetJSON with rows
	rows, ok := first["rows"].([]any)
	if !ok || len(rows) == 0 {
		t.Fatalf("expected rows in first sheet, got %v", first["rows"])
	}
	row1 := rows[0].(map[string]any)["cells"].(map[string]any)
	a1 := row1["A"].(map[string]any)
	if a1["value"] != "Name" {
		t.Errorf("expected Sheet1!A1='Name', got %v", a1["value"])
	}
}

func TestReadCommand_JSON_AllSheets(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{allSheets: true, jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	sheets, ok := result["sheets"].([]any)
	if !ok {
		t.Fatalf("expected 'sheets' array, got %v", result["sheets"])
	}
	if len(sheets) != 2 {
		t.Fatalf("expected all 2 sheets, got %d", len(sheets))
	}

	// Workbook order: Sheet1 created first, Data second
	if sheets[0].(map[string]any)["sheet"] != "Sheet1" {
		t.Errorf("expected workbook order with 'Sheet1' first, got %v",
			sheets[0].(map[string]any)["sheet"])
	}
	if sheets[1].(map[string]any)["sheet"] != "Data" {
		t.Errorf("expected 'Data' second, got %v", sheets[1].(map[string]any)["sheet"])
	}
}

func TestReadCommand_Sheets_Whitespace(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: " Sheet1 , Data ", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}
	sheets := result["sheets"].([]any)
	if len(sheets) != 2 {
		t.Fatalf("expected 2 sheets with whitespace-tolerant parsing, got %d", len(sheets))
	}
}

func TestReadCommand_Sheets_WithRange(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1,Data", cellRange: "A1:A1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `"Name"`) {
		t.Errorf("expected Sheet1!A1='Name', got:\n%s", output)
	}
	if !strings.Contains(output, `"ID"`) {
		t.Errorf("expected Data!A1='ID', got:\n%s", output)
	}
	if strings.Contains(output, `"Alpha"`) {
		t.Errorf("should not contain 'Alpha' (outside A1:A1 range), got:\n%s", output)
	}
}

func TestReadCommand_Sheets_UnknownSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1,Nope"}, &stdout)
	if err == nil {
		t.Fatal("expected error for unknown sheet in --sheets")
	}
	if !strings.Contains(err.Error(), "not found") {
		t.Errorf("expected 'not found' in error, got: %v", err)
	}
	if !strings.Contains(err.Error(), "Available sheets") {
		t.Errorf("expected available-sheets listing in error, got: %v", err)
	}
}

func TestReadCommand_Sheets_EmptyList(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: " , "}, &stdout)
	if err == nil {
		t.Fatal("expected error for empty sheet list")
	}
	if !strings.Contains(err.Error(), "no sheet names") {
		t.Errorf("expected 'no sheet names' in error, got: %v", err)
	}
}

func TestReadCommand_Sheets_MutuallyExclusiveWithAllSheets(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1", allSheets: true}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --sheets with --all-sheets")
	}
	if !strings.Contains(err.Error(), "--sheets cannot be combined with --all-sheets") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

func TestReadCommand_Sheets_MutuallyExclusiveWithPositionalSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{sheets: "Data"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --sheets with positional sheet")
	}
	if !strings.Contains(err.Error(), "positional sheet argument") {
		t.Errorf("expected positional-sheet error, got: %v", err)
	}
}

func TestReadCommand_Sheets_MutuallyExclusiveWithCell(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{sheets: "Sheet1", cell: "A1"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --sheets with --cell")
	}
	if !strings.Contains(err.Error(), "--sheets cannot be combined with --cell or --cells") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

func TestReadCommand_AllSheets_MutuallyExclusiveWithCells(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "", readOptions{allSheets: true, cells: "A1,B2"}, &stdout)
	if err == nil {
		t.Fatal("expected error when combining --all-sheets with --cells")
	}
	if !strings.Contains(err.Error(), "--all-sheets cannot be combined with --cell or --cells") {
		t.Errorf("expected mutual-exclusion error, got: %v", err)
	}
}

// --- Error value marking tests ---

func TestRead_MarksErrorValues(t *testing.T) {
	path := testutil.CreateCachedErrorFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	// A3 has formula =A1/A2 with cached error #DIV/0!
	if !strings.Contains(output, "[ERROR]") {
		t.Errorf("expected '[ERROR]' marker in output, got:\n%s", output)
	}
	if !strings.Contains(output, "#DIV/0!") {
		t.Errorf("expected '#DIV/0!' in output, got:\n%s", output)
	}
}

func TestRead_MarksErrorValues_SingleCell(t *testing.T) {
	path := testutil.CreateCachedErrorFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A3"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())

	if !strings.Contains(output, "=A1/A2") {
		t.Errorf("expected formula '=A1/A2' in output, got: %q", output)
	}
	if !strings.Contains(output, "#DIV/0!") {
		t.Errorf("expected '#DIV/0!' in output, got: %q", output)
	}
	if !strings.Contains(output, "[ERROR]") {
		t.Errorf("expected '[ERROR]' marker in output, got: %q", output)
	}
}

func TestRead_JSON_ErrorField(t *testing.T) {
	path := testutil.CreateCachedErrorFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A3", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if result["formula"] != "A1/A2" {
		t.Errorf("expected formula 'A1/A2', got %v", result["formula"])
	}
	if result["error"] != "#DIV/0!" {
		t.Errorf("expected error '#DIV/0!', got %v", result["error"])
	}
	if result["value"] != nil {
		t.Errorf("expected value nil for error cell, got %v", result["value"])
	}
}

func TestRead_JSON_ErrorField_Range(t *testing.T) {
	path := testutil.CreateCachedErrorFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	rows := result["rows"].([]any)
	// A3 is row 3 (index 2)
	row3 := rows[2].(map[string]any)
	cells := row3["cells"].(map[string]any)
	a3 := cells["A"].(map[string]any)

	if a3["formula"] != "A1/A2" {
		t.Errorf("expected formula 'A1/A2', got %v", a3["formula"])
	}
	if a3["error"] != "#DIV/0!" {
		t.Errorf("expected error '#DIV/0!', got %v", a3["error"])
	}
	if a3["value"] != nil {
		t.Errorf("expected value nil for error cell, got %v", a3["value"])
	}
}

func TestReadCommand_InvalidSingleCellRef(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "GARBAGE"}, &stdout)
	if err == nil {
		t.Fatal("expected error for invalid --cell reference, got '(empty)' behavior")
	}
	if !strings.Contains(err.Error(), "GARBAGE") {
		t.Errorf("expected error naming the offending ref, got: %v", err)
	}
}

func TestReadCommand_ReversedRange_Normalized(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cellRange: "B2:A1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `"Name"`) {
		t.Errorf("reversed range must show A1 ('Name'), got:\n%s", output)
	}
	if !strings.Contains(output, "10") {
		t.Errorf("reversed range must show B2 (10), got:\n%s", output)
	}
	if !strings.Contains(output, "A1:B2") {
		t.Errorf("range header must be normalized to A1:B2, got:\n%s", output)
	}
}

// --- Raw value + formatted display tests (J6) ---

func TestRead_JSON_RawValueAndDisplay_Percent(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A1", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	// value must be the raw NUMBER (0.225), not the formatted "22.50%" string.
	num, ok := result["value"].(float64)
	if !ok {
		t.Fatalf("expected numeric value, got %T (%v)", result["value"], result["value"])
	}
	if num != 0.225 {
		t.Errorf("expected value 0.225, got %v", num)
	}
	if result["display"] != "22.50%" {
		t.Errorf("expected display '22.50%%', got %v", result["display"])
	}
}

func TestRead_Plain_RawValueAndDisplay_Percent(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A1"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	// Plain output must contain BOTH the raw value and the bracketed display.
	if !strings.Contains(output, "0.225") {
		t.Errorf("expected raw value '0.225' in plain output, got: %q", output)
	}
	if !strings.Contains(output, `[display "22.50%"]`) {
		t.Errorf("expected bracketed display, got: %q", output)
	}
}

func TestRead_JSON_PlainNumber_NoDisplay(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A3", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if num, ok := result["value"].(float64); !ok || num != 42 {
		t.Errorf("expected numeric value 42, got %T (%v)", result["value"], result["value"])
	}
	// formatted == raw → no redundant display field.
	if _, has := result["display"]; has {
		t.Errorf("plain number must not carry a display field, got %v", result["display"])
	}
}

func TestRead_Plain_PlainNumber_NoDisplaySuffix(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A3"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if output != "42" {
		t.Errorf("expected bare '42' with no [display] suffix, got: %q", output)
	}
}

func TestRead_JSON_TextCell_NotCoercedToNumber(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A4", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	// "00123" is a text cell — it must stay a string, not become 123.
	if result["value"] != "00123" {
		t.Errorf("expected string value '00123', got %T (%v)", result["value"], result["value"])
	}
	if _, has := result["display"]; has {
		t.Errorf("text cell must not carry a display field, got %v", result["display"])
	}
}

func TestRead_JSON_FormulaWithFormattedCachedValue(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A5", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if result["formula"] != "A1+A2" {
		t.Errorf("expected formula 'A1+A2', got %v", result["formula"])
	}
	if num, ok := result["value"].(float64); !ok || num != 0.24 {
		t.Errorf("expected cached numeric value 0.24, got %T (%v)", result["value"], result["value"])
	}
	if result["display"] != "24.00%" {
		t.Errorf("expected display '24.00%%', got %v", result["display"])
	}
}

func TestRead_Plain_FormulaWithFormattedCachedValue(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A5"}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := strings.TrimSpace(stdout.String())
	if !strings.Contains(output, "=A1+A2 → 0.24") {
		t.Errorf("expected '=A1+A2 → 0.24', got: %q", output)
	}
	if !strings.Contains(output, `[display "24.00%"]`) {
		t.Errorf("expected bracketed display for formula, got: %q", output)
	}
}

func TestRead_JSON_LargeIntegerPreservesExactDigits(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{cell: "A6", jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// The integer 9007199254740993 (2^53 + 1) is not representable as a float64
	// (it would round to …992). It must be preserved verbatim as a json.Number,
	// and emitted UNQUOTED in the JSON text (i.e. a JSON number, not a string).
	raw := stdout.String()
	if !strings.Contains(raw, `"value": 9007199254740993`) {
		t.Errorf("expected unquoted exact integer in JSON, got:\n%s", raw)
	}
	if strings.Contains(raw, `"value": "9007199254740993"`) {
		t.Errorf("value must be a JSON number, not a quoted string:\n%s", raw)
	}

	// Decode with UseNumber so the value is not coerced through float64.
	dec := json.NewDecoder(bytes.NewReader(stdout.Bytes()))
	dec.UseNumber()
	var result map[string]any
	if err := dec.Decode(&result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, raw)
	}
	num, ok := result["value"].(json.Number)
	if !ok {
		t.Fatalf("expected json.Number value, got %T (%v)", result["value"], result["value"])
	}
	if num.String() != "9007199254740993" {
		t.Errorf("expected exact '9007199254740993', got %q", num.String())
	}
}
