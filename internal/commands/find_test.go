package commands

import (
	"bytes"
	"encoding/json"
	"errors"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

// --- Plain text (default) tests ---

func TestFindCommand_Plain_SubstringCaseInsensitive(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "alpha", "", findOptions{max: 50}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `Sheet1!A2: "Alpha"`) {
		t.Errorf("expected case-insensitive substring hit 'Sheet1!A2: \"Alpha\"', got:\n%s", output)
	}
}

func TestFindCommand_Plain_SearchesAllSheets(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "Amount", "", findOptions{max: 50}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `Data!B1: "Amount"`) {
		t.Errorf("expected hit on second sheet 'Data!B1: \"Amount\"', got:\n%s", output)
	}
}

func TestFindCommand_Plain_SheetNarrowing(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "Alpha", "Data", findOptions{max: 50}, &stdout)
	if !errors.Is(err, ErrNoMatches) {
		t.Fatalf("expected ErrNoMatches for zero matches, got: %v", err)
	}

	output := stdout.String()
	if strings.Contains(output, "Sheet1!A2") {
		t.Errorf("sheet narrowing to 'Data' must not search Sheet1, got:\n%s", output)
	}
	if !strings.Contains(output, `No matches found for "Alpha".`) {
		t.Errorf("expected exact zero-match message, got:\n%s", output)
	}
}

func TestFindCommand_Exact(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// Case mismatch must not match in exact mode
	var stdout bytes.Buffer
	err := runFind(path, "alpha", "", findOptions{exact: true, max: 50}, &stdout)
	if !errors.Is(err, ErrNoMatches) {
		t.Fatalf("expected ErrNoMatches, got: %v", err)
	}
	if strings.Contains(stdout.String(), "A2") {
		t.Errorf("exact mode must be case-sensitive, got:\n%s", stdout.String())
	}

	// Substring must not match in exact mode
	stdout.Reset()
	err = runFind(path, "Alph", "", findOptions{exact: true, max: 50}, &stdout)
	if !errors.Is(err, ErrNoMatches) {
		t.Fatalf("expected ErrNoMatches, got: %v", err)
	}
	if strings.Contains(stdout.String(), "A2") {
		t.Errorf("exact mode must be full-cell match, got:\n%s", stdout.String())
	}

	// Full-cell exact match works
	stdout.Reset()
	err = runFind(path, "Alpha", "", findOptions{exact: true, max: 50}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(stdout.String(), `Sheet1!A2: "Alpha"`) {
		t.Errorf("expected exact full-cell match, got:\n%s", stdout.String())
	}
}

func TestFindCommand_Formulas(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// Without --formulas, "SUM" matches nothing (B4's cached value is empty)
	var stdout bytes.Buffer
	err := runFind(path, "SUM", "", findOptions{max: 50}, &stdout)
	if !errors.Is(err, ErrNoMatches) {
		t.Fatalf("expected ErrNoMatches, got: %v", err)
	}
	if strings.Contains(stdout.String(), "B4") {
		t.Errorf("without --formulas formula text must not match, got:\n%s", stdout.String())
	}

	// With --formulas, the formula string is searched too
	stdout.Reset()
	err = runFind(path, "SUM", "", findOptions{formulas: true, max: 50}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	if !strings.Contains(stdout.String(), "Sheet1!B4") {
		t.Errorf("expected formula hit 'Sheet1!B4', got:\n%s", stdout.String())
	}
}

func TestFindCommand_Plain_MaxTruncation(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	// "a" matches 6 cells: Name, Value, Alpha, Beta, Total (Sheet1) + Amount (Data)
	var stdout bytes.Buffer
	err := runFind(path, "a", "", findOptions{max: 2}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, "(showing first 2 of 6 matches — raise --max)") {
		t.Errorf("expected explicit truncation notice, got:\n%s", output)
	}
	hitLines := 0
	for line := range strings.SplitSeq(strings.TrimSpace(output), "\n") {
		if strings.Contains(line, "!") {
			hitLines++
		}
	}
	if hitLines != 2 {
		t.Errorf("expected exactly 2 match lines, got %d:\n%s", hitLines, output)
	}
}

// --- JSON (--json flag) tests ---

func TestFindCommand_JSON(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "Alpha", "", findOptions{max: 50, jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	matches, ok := result["matches"].([]any)
	if !ok || len(matches) != 1 {
		t.Fatalf("expected 1 match, got %v", result["matches"])
	}
	m := matches[0].(map[string]any)
	if m["sheet"] != "Sheet1" || m["cell"] != "A2" || m["value"] != "Alpha" {
		t.Errorf("expected {Sheet1, A2, Alpha}, got %v", m)
	}
	if m["kind"] != "cell" {
		t.Errorf("expected kind 'cell', got %v", m["kind"])
	}
	if result["total"] != float64(1) {
		t.Errorf("expected total 1, got %v", result["total"])
	}
	if result["truncated"] != false {
		t.Errorf("expected truncated false, got %v", result["truncated"])
	}
}

func TestFindCommand_JSON_Truncated(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "a", "", findOptions{max: 2, jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	matches := result["matches"].([]any)
	if len(matches) != 2 {
		t.Errorf("expected 2 matches shown, got %d", len(matches))
	}
	if result["total"] != float64(6) {
		t.Errorf("expected total 6, got %v", result["total"])
	}
	if result["truncated"] != true {
		t.Errorf("expected truncated true, got %v", result["truncated"])
	}
}

func TestFindCommand_JSON_FormulaField(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "SUM", "", findOptions{formulas: true, max: 50, jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	matches := result["matches"].([]any)
	if len(matches) != 1 {
		t.Fatalf("expected 1 match, got %d", len(matches))
	}
	m := matches[0].(map[string]any)
	if m["cell"] != "B4" {
		t.Errorf("expected cell B4, got %v", m["cell"])
	}
	if m["formula"] != "SUM(B2:B3)" {
		t.Errorf("expected formula 'SUM(B2:B3)', got %v", m["formula"])
	}
}

func TestFindCommand_JSON_FormattedCell_SurfacesDisplay(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	// find matches against the FORMATTED value ("22.50%"); the result must keep
	// the raw value (0.225) AND surface the formatted display the user searched.
	var stdout bytes.Buffer
	err := runFind(path, "22.50%", "Sheet1", findOptions{max: 50, jsonOut: true}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	matches := result["matches"].([]any)
	if len(matches) != 1 {
		t.Fatalf("expected 1 match for '22.50%%', got %d:\n%s", len(matches), stdout.String())
	}
	m := matches[0].(map[string]any)
	if m["cell"] != "A1" {
		t.Errorf("expected cell A1, got %v", m["cell"])
	}
	// find's JSON value stays the raw stored string (0.225), with display added.
	if m["value"] != "0.225" {
		t.Errorf("expected raw value '0.225', got %v", m["value"])
	}
	if m["display"] != "22.50%" {
		t.Errorf("expected display '22.50%%', got %v", m["display"])
	}
}

func TestFindCommand_Plain_FormattedCell_ShowsDisplay(t *testing.T) {
	path := testutil.CreateNumberFormatFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "22.50%", "Sheet1", findOptions{max: 50}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, "0.225") {
		t.Errorf("expected raw value '0.225' in plain find output, got:\n%s", output)
	}
	if !strings.Contains(output, `[display "22.50%"]`) {
		t.Errorf("expected bracketed display in plain find output, got:\n%s", output)
	}
}

func TestFindCommand_JSON_NoMatches(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "zzz-no-such-text", "", findOptions{max: 50, jsonOut: true}, &stdout)
	if !errors.Is(err, ErrNoMatches) {
		t.Fatalf("expected ErrNoMatches for zero matches, got: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	matches, ok := result["matches"].([]any)
	if !ok {
		t.Fatalf("expected 'matches' to be an empty array (not null), got %v", result["matches"])
	}
	if len(matches) != 0 {
		t.Errorf("expected 0 matches, got %d", len(matches))
	}
	if result["total"] != float64(0) {
		t.Errorf("expected total 0, got %v", result["total"])
	}
}

// --- Error cases ---

func TestFindCommand_UnknownSheet(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "Alpha", "Nope", findOptions{max: 50}, &stdout)
	if err == nil {
		t.Fatal("expected error for unknown sheet")
	}
	if !strings.Contains(err.Error(), "not found") {
		t.Errorf("expected 'not found' in error, got: %v", err)
	}
	if !strings.Contains(err.Error(), "Available sheets") {
		t.Errorf("expected available-sheets listing, got: %v", err)
	}
}

func TestFindCommand_EmptyText(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "", "", findOptions{max: 50}, &stdout)
	if err == nil {
		t.Fatal("expected error for empty search text")
	}
	if !strings.Contains(err.Error(), "search text") {
		t.Errorf("expected 'search text' in error, got: %v", err)
	}
}

func TestFindCommand_InvalidMax(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runFind(path, "Alpha", "", findOptions{max: 0}, &stdout)
	if err == nil {
		t.Fatal("expected error for --max < 1")
	}
	if !strings.Contains(err.Error(), "--max") {
		t.Errorf("expected '--max' in error, got: %v", err)
	}
}
