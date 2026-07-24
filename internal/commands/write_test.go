package commands

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

func TestWriteCommand_String(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "A5", "NewValue", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A5")
	if val != "NewValue" {
		t.Errorf("expected 'NewValue', got %q", val)
	}

	output := stdout.String()
	if !strings.Contains(output, "A5") {
		t.Errorf("expected cell ref in output, got:\n%s", output)
	}
}

func TestWriteCommand_Number(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "B5", "42.5", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "B5")
	if val != "42.5" {
		t.Errorf("expected '42.5', got %q", val)
	}
}

func TestWriteCommand_PreservesExistingCells(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "C1", "Extra", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A1")
	if val != "Name" {
		t.Errorf("existing cell A1 corrupted: expected 'Name', got %q", val)
	}
	val = testutil.ReadCellValue(t, path, "Sheet1", "B2")
	if val != "10" {
		t.Errorf("existing cell B2 corrupted: expected '10', got %q", val)
	}
}

func TestWriteCommand_PreservesFormatting(t *testing.T) {
	path := testutil.CreateFormattedFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "B3", "written", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	font := testutil.GetCellStyleFont(t, path, "Sheet1", "A1")
	if font == nil || !font.Bold {
		t.Error("formatting on A1 was lost after writing B3")
	}
}

func TestWriteCommand_PreservesFormulas(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "A5", "Extra", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula := testutil.ReadCellFormula(t, path, "Sheet1", "B4")
	if formula != "SUM(B2:B3)" {
		t.Errorf("formula in B4 was corrupted: expected 'SUM(B2:B3)', got %q", formula)
	}
}

func TestWriteCommand_SheetNotFound(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "NonExistent", "A1", "test", &stdout)
	if err == nil {
		t.Fatal("expected error for nonexistent sheet")
	}
}

// --- Formula writing tests ---

func TestWriteCommand_Formula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "A4", "=SUM(B2:B3)", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula := testutil.ReadCellFormula(t, path, "Sheet1", "A4")
	if formula != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)', got %q", formula)
	}

	// write no longer recalcs — the cache stays untouched (here: absent,
	// since excelize never computes on write) until 'validate' recalcs.
	val := testutil.ReadCellValue(t, path, "Sheet1", "A4")
	if val != "" {
		t.Errorf("expected no cached value before validate recalcs, got %q", val)
	}

	if !strings.Contains(stdout.String(), "=SUM(B2:B3)") {
		t.Errorf("expected formula in output, got:\n%s", stdout.String())
	}
}

func TestWriteCommand_Formula_OutsideDimension(t *testing.T) {
	// Fixture's stored dimension is A1:B4; F10 lies outside. The write must
	// still extend the dimension so later full-sheet reads see the new cell.
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "F10", "=SUM(B2:B3)", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	formula := testutil.ReadCellFormula(t, path, "Sheet1", "F10")
	if formula != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)', got %q", formula)
	}

	// write no longer recalcs — no cached value yet, regardless of dimension.
	val := testutil.ReadCellValue(t, path, "Sheet1", "F10")
	if val != "" {
		t.Errorf("expected no cached value before validate recalcs, got %q", val)
	}
}

func TestWriteCommand_ExtendsStoredDimension(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "F10", "outside", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("reopen: %v", err)
	}
	defer f.Close()

	dim, err := f.GetSheetDimension("Sheet1")
	if err != nil {
		t.Fatalf("get dimension: %v", err)
	}
	if dim != "A1:F10" {
		t.Errorf("expected stored dimension extended to 'A1:F10', got %q", dim)
	}
}

func TestWriteCommand_DegenerateFormula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	for _, value := range []string{"=", "=   ", "==B1"} {
		var stdout bytes.Buffer
		err := runWrite(path, "Sheet1", "A5", value, &stdout)
		if err == nil {
			t.Errorf("expected error for degenerate formula %q", value)
			continue
		}
		if !strings.Contains(err.Error(), "formula") {
			t.Errorf("expected 'formula' in error for %q, got: %v", value, err)
		}
	}

	// Nothing must have been written
	val := testutil.ReadCellValue(t, path, "Sheet1", "A5")
	if val != "" {
		t.Errorf("degenerate formula must not write anything, got %q", val)
	}
}

func TestWriteCommand_SkipsLockedOnLegacyPasswordProtectedSheet(t *testing.T) {
	path := testutil.CreateLegacyPasswordProtectedFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "A2", "99", &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if !strings.Contains(stdout.String(), "Skipped") {
		t.Errorf("expected skip notice for legacy-password-protected sheet, got:\n%s",
			stdout.String())
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A2")
	if val != "42" {
		t.Errorf("locked cell on legacy-protected sheet must stay untouched: expected '42', got %q",
			val)
	}
}

// --- Structured (--json) result tests ---

func TestWriteCommand_JSON_Value(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A5", "NewValue", writeOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Pure JSON, no prose before/after
	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be pure valid JSON: %v\n%s", err, stdout.String())
	}

	written, ok := result["written"].([]any)
	if !ok || len(written) != 1 {
		t.Fatalf("expected 1 written entry, got %v", result["written"])
	}
	entry := written[0].(map[string]any)
	if entry["sheet"] != "Sheet1" || entry["cell"] != "A5" || entry["value"] != "NewValue" {
		t.Errorf("expected {Sheet1, A5, NewValue}, got %v", entry)
	}

	skipped, ok := result["skipped_locked"].([]any)
	if !ok {
		t.Fatalf("expected 'skipped_locked' to be an empty array (not null), got %v",
			result["skipped_locked"])
	}
	if len(skipped) != 0 {
		t.Errorf("expected 0 skipped cells, got %d", len(skipped))
	}
}

func TestWriteCommand_JSON_Formula(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "C5", "=SUM(B2:B3)", writeOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	written := result["written"].([]any)
	if len(written) != 1 {
		t.Fatalf("expected 1 written entry, got %d", len(written))
	}
	entry := written[0].(map[string]any)
	if entry["formula"] != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)' (no leading =), got %v", entry["formula"])
	}
	if _, hasValue := entry["value"]; hasValue {
		t.Errorf("formula write entry should not have 'value' key, got %v", entry)
	}
}

func TestWriteCommand_JSON_SkippedLocked(t *testing.T) {
	path := testutil.CreateProtectedFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A2", "99", writeOptions{jsonOut: true}, &stdout)
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
	if len(written) != 0 {
		t.Errorf("expected 0 written entries, got %d", len(written))
	}

	// Cell must be untouched
	val := testutil.ReadCellValue(t, path, "Sheet1", "A2")
	if val != "42" {
		t.Errorf("locked cell must stay untouched: expected '42', got %q", val)
	}
}

func TestWriteCommand_JSON_SkippedLocked_FormulaCell(t *testing.T) {
	path := testutil.CreateProtectedFormulaFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A2", "99", writeOptions{jsonOut: true}, &stdout)
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

func TestWriteCommand_JSON_SkippedLocked_PlainCell(t *testing.T) {
	path := testutil.CreateProtectedFormulaFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A1", "99", writeOptions{jsonOut: true}, &stdout)
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

func TestWriteCommand_InvalidCellRef(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWrite(path, "Sheet1", "NOPE", "1", &stdout)
	if err == nil {
		t.Fatal("expected error for invalid cell reference")
	}
	if !strings.Contains(err.Error(), "NOPE") {
		t.Errorf("expected error naming the offending ref, got: %v", err)
	}
}

func TestWriteCommand_InvalidCellRef_OnProtectedSheet(t *testing.T) {
	// Invalid refs must error BEFORE the protection check — previously the
	// fail-safe lock detection misreported them as skipped_locked.
	path := testutil.CreateProtectedFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "NOPE", "1", writeOptions{jsonOut: true}, &stdout)
	if err == nil {
		t.Fatalf("expected error for invalid ref on protected sheet, got output:\n%s",
			stdout.String())
	}
	if !strings.Contains(err.Error(), "NOPE") {
		t.Errorf("expected error naming the offending ref, got: %v", err)
	}
	if strings.Contains(stdout.String(), "skipped_locked") {
		t.Errorf("invalid ref must not be reported as skipped_locked, got:\n%s", stdout.String())
	}
}

func TestWriteCommand_StringFlag_NoCoercion(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A5", "00123", writeOptions{asString: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A5")
	if val != "00123" {
		t.Errorf("--string must preserve leading zeros, got %q", val)
	}
}

func TestWriteCommand_StringFlag_NoFormulaInterpretation(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A5", "=SUM(B2:B3)", writeOptions{asString: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A5")
	if val != "=SUM(B2:B3)" {
		t.Errorf("--string must store leading-= literally, got %q", val)
	}
	formula := testutil.ReadCellFormula(t, path, "Sheet1", "A5")
	if formula != "" {
		t.Errorf("--string must not create a formula, got %q", formula)
	}
}

func TestWriteCommand_EchoShowsStoredValue(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// "1e5" is coerced to float 100000 — the echo must show the STORED value.
	var stdout bytes.Buffer
	if err := runWrite(path, "Sheet1", "A5", "1e5", &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(stdout.String(), "100000") {
		t.Errorf("echo must show stored value 100000, got: %q", stdout.String())
	}
	if !strings.Contains(stdout.String(), "float") {
		t.Errorf("echo must name the stored type, got: %q", stdout.String())
	}

	// "t" is coerced to boolean true.
	stdout.Reset()
	if err := runWrite(path, "Sheet1", "A6", "t", &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(stdout.String(), "true") || !strings.Contains(stdout.String(), "boolean") {
		t.Errorf("echo must show stored boolean true, got: %q", stdout.String())
	}

	// plain string keeps quotes
	stdout.Reset()
	if err := runWrite(path, "Sheet1", "A7", "hello", &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(stdout.String(), `"hello" (string)`) {
		t.Errorf("echo must show quoted string with type, got: %q", stdout.String())
	}
}

// --- parseValue unit tests ---

func TestParseValueGermanDecimalComma(t *testing.T) {
	cases := []struct {
		in   string
		want any
	}{
		{"23,45", 23.45},
		{"1.234,56", 1234.56},
		{"1.234.567,89", 1234567.89},
		{" 15,0 ", 15.0},
		{"-3,5", -3.5},
	}
	for _, c := range cases {
		got := parseValue(c.in)
		if got != c.want {
			t.Errorf("parseValue(%q) = %v (%T), want %v", c.in, got, got, c.want)
		}
	}
}

func TestParseValueAsciiUnchanged(t *testing.T) {
	if got := parseValue("1.234"); got != 1.234 {
		t.Errorf("ASCII dot-decimal must stay: got %v", got)
	}
	if got := parseValue("42"); got != int64(42) {
		t.Errorf("int parsing must stay: got %v (%T)", got, got)
	}
	if got := parseValue("true"); got != true {
		t.Errorf("bool parsing must stay: got %v", got)
	}
}

func TestParseValueNonNumericStaysString(t *testing.T) {
	for _, s := range []string{"5,94 %", "24,75 €", "abc", "1,2,3", ""} {
		if got := parseValue(s); got != s {
			t.Errorf("parseValue(%q) = %v, want unchanged string", s, got)
		}
	}
}

// --- near-numeric warning tests ---

func TestLooksNumericButUnparseable(t *testing.T) {
	positives := []string{"5,94 %", "24,75 €", "1,2,3", "12 345,6"}
	for _, s := range positives {
		if !looksNumericButUnparseable(s) {
			t.Errorf("expected warning for %q", s)
		}
	}
	negatives := []string{"abc", "23,45", "42", "", "Pos. 90", "=SUM(A1:A2)"}
	for _, s := range negatives {
		if looksNumericButUnparseable(s) {
			t.Errorf("did not expect warning for %q", s)
		}
	}
}

func TestWriteCommand_WarnsNearNumericStoredAsText(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// "5,94 %" stays text (parseValue won't coerce units) — must warn.
	var stdout bytes.Buffer
	if err := runWrite(path, "Sheet1", "A5", "5,94 %", &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if val := testutil.ReadCellValue(t, path, "Sheet1", "A5"); val != "5,94 %" {
		t.Errorf("near-numeric value must stay text, got %q", val)
	}
	out := stdout.String()
	if !strings.Contains(out, "Warning:") || !strings.Contains(out, "TEXT") {
		t.Errorf("expected near-numeric TEXT warning, got:\n%s", out)
	}

	// "23,45" coerces fine (German comma) — must NOT warn.
	stdout.Reset()
	if err := runWrite(path, "Sheet1", "A6", "23,45", &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if strings.Contains(stdout.String(), "looks numeric") {
		t.Errorf("coercible value must not warn, got:\n%s", stdout.String())
	}
}

func TestWriteCommand_JSON_WarningsAlwaysPresent(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// A plain string write produces no warning, but the array must still be there.
	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A5", "NewValue", writeOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}
	warnings, ok := result["warnings"].([]any)
	if !ok {
		t.Fatalf("expected 'warnings' to be an empty array (not null), got %v", result["warnings"])
	}
	if len(warnings) != 0 {
		t.Errorf("expected 0 warnings, got %d", len(warnings))
	}

	// A near-numeric write fills the array.
	stdout.Reset()
	err = runWriteOpts(path, "Sheet1", "A6", "5,94 %", writeOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}
	warnings, ok = result["warnings"].([]any)
	if !ok || len(warnings) != 1 {
		t.Fatalf("expected 1 warning, got %v", result["warnings"])
	}
	nw := warnings[0].(map[string]any)
	if nw["cell"] != "A6" || nw["value"] != "5,94 %" {
		t.Errorf("expected {A6, '5,94 %%'}, got %v", nw)
	}
}

func TestWriteCommand_StringFlag_SuppressesNearNumericWarning(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// --string is the documented silencer: no coercion, no warning.
	var stdout bytes.Buffer
	err := runWriteOpts(path, "Sheet1", "A5", "5,94 %", writeOptions{asString: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if strings.Contains(stdout.String(), "looks numeric") {
		t.Errorf("--string must suppress the near-numeric warning, got:\n%s", stdout.String())
	}
}

// --- batch string-value coercion tests ---

func TestBatchCommand_CoercesGermanCommaString(t *testing.T) {
	// A batch "value" string must be coerced like `write` does: a German
	// decimal-comma number is stored as the NUMBER, not as text.
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "1.234,56"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	if err := runBatchFromReader(xlsxPath, stdinReader, &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Stored as the number 1234.56 (excelize renders without trailing zeros).
	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "1234.56" {
		t.Errorf("expected coerced number '1234.56', got %q", val)
	}
	// No formula should have been created.
	if f := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "A5"); f != "" {
		t.Errorf("value coercion must not create a formula, got %q", f)
	}
	// Echo shows the coerced number, not the raw German string.
	if !strings.Contains(stdout.String(), "1234.56") {
		t.Errorf("echo must show coerced number, got:\n%s", stdout.String())
	}
}

func TestBatchCommand_LeadingEqualsStringStaysLiteral(t *testing.T) {
	// A "=" -prefixed value field is never a formula and parseValue returns it
	// unchanged — it must remain literal text after coercion routing.
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "=abc"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	if err := runBatchFromReader(xlsxPath, stdinReader, &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "=abc" {
		t.Errorf("expected literal '=abc', got %q", val)
	}
	if f := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "A5"); f != "" {
		t.Errorf("value field must never be a formula, got %q", f)
	}
}

func TestBatchCommand_NearNumericStringWarnsAndStaysText(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "5,94 %"}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	// Stays text (units not coerced).
	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "5,94 %" {
		t.Errorf("near-numeric value must stay text, got %q", val)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}
	warnings, ok := result["warnings"].([]any)
	if !ok || len(warnings) != 1 {
		t.Fatalf("expected 1 warning in batch JSON, got %v", result["warnings"])
	}
	nw := warnings[0].(map[string]any)
	if nw["cell"] != "A5" || nw["value"] != "5,94 %" {
		t.Errorf("expected {A5, '5,94 %%'}, got %v", nw)
	}
	// The written value must echo as the (text) value.
	written := result["written"].([]any)
	if len(written) != 1 || written[0].(map[string]any)["value"] != "5,94 %" {
		t.Errorf("expected written value '5,94 %%', got %v", result["written"])
	}
}

func TestBatchCommand_JSON_WarningsAlwaysPresent(t *testing.T) {
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": "Gamma"}]`
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
	warnings, ok := result["warnings"].([]any)
	if !ok {
		t.Fatalf("expected 'warnings' empty array (not null), got %v", result["warnings"])
	}
	if len(warnings) != 0 {
		t.Errorf("expected 0 warnings for a plain string, got %d", len(warnings))
	}
}

func TestBatchCommand_JSONNumberStoredUnchangedNoWarning(t *testing.T) {
	// A JSON number (float64) bypasses string coercion entirely: it is stored
	// as the number unchanged and must NOT produce a near-numeric warning.
	xlsxPath := testutil.CreateSimpleFixture(t)

	stdinContent := `[{"sheet": "Sheet1", "cell": "A5", "value": 1234.56}]`
	stdinReader := strings.NewReader(stdinContent)

	var stdout bytes.Buffer
	err := runBatchFromReaderOpts(xlsxPath, stdinReader, batchOptions{jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	if val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A5"); val != "1234.56" {
		t.Errorf("expected number '1234.56' stored unchanged, got %q", val)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}
	warnings, ok := result["warnings"].([]any)
	if !ok || len(warnings) != 0 {
		t.Errorf("JSON number must not warn, got %v", result["warnings"])
	}
	written := result["written"].([]any)
	if len(written) != 1 || written[0].(map[string]any)["value"] != 1234.56 {
		t.Errorf("expected written value 1234.56 (number), got %v", result["written"])
	}
}
