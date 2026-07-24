package commands

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

func TestInfoCommand_Plain(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, false, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	if !strings.Contains(output, "File: simple.xlsx") {
		t.Errorf("expected 'File: simple.xlsx', got:\n%s", output)
	}

	if !strings.Contains(output, "Sheets: 2") {
		t.Errorf("expected 'Sheets: 2', got:\n%s", output)
	}

	if !strings.Contains(output, "- Sheet1") {
		t.Errorf("expected '- Sheet1' in output, got:\n%s", output)
	}

	if !strings.Contains(output, "- Data") {
		t.Errorf("expected '- Data' in output, got:\n%s", output)
	}

	// Dimensions should be in parentheses
	if !strings.Contains(output, "(A1:B4)") {
		t.Errorf("expected dimension '(A1:B4)' for Sheet1, got:\n%s", output)
	}

	// Should NOT be valid JSON
	var js any
	if json.Unmarshal(stdout.Bytes(), &js) == nil {
		t.Error("plain text output should NOT be valid JSON")
	}
}

func TestInfoCommand_JSON(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, true, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result InfoJSON
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if result.File != "simple.xlsx" {
		t.Errorf("expected file 'simple.xlsx', got %q", result.File)
	}

	if len(result.Sheets) != 2 {
		t.Fatalf("expected 2 sheets, got %d", len(result.Sheets))
	}

	sheetNames := map[string]bool{}
	for _, s := range result.Sheets {
		sheetNames[s.Name] = true
	}

	if !sheetNames["Sheet1"] {
		t.Errorf("expected 'Sheet1' in sheets")
	}
	if !sheetNames["Data"] {
		t.Errorf("expected 'Data' in sheets")
	}
}

func TestInfo_ShowsFormulaCount(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, false, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	// Sheet1 has SUM(B2:B3) in B4 → 1 formula
	if !strings.Contains(output, "Formulas: 1") {
		t.Errorf("expected 'Formulas: 1' for Sheet1, got:\n%s", output)
	}

	// Data sheet has no formulas → 0
	if !strings.Contains(output, "Formulas: 0") {
		t.Errorf("expected 'Formulas: 0' for Data sheet, got:\n%s", output)
	}
}

func TestInfo_ShowsProtectedSheet(t *testing.T) {
	path := testutil.CreateProtectedFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, false, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	if !strings.Contains(output, "Protected") {
		t.Errorf("expected 'Protected' in output for protected sheet, got:\n%s", output)
	}

	// Sheet2 is not protected — its line should not contain "Protected"
	lines := strings.Split(output, "\n")
	for _, line := range lines {
		if strings.Contains(line, "Sheet2") && strings.Contains(line, "Protected") {
			t.Errorf("Sheet2 should not be protected, got:\n%s", line)
		}
	}
}

func TestInfo_JSON_IncludesNewFields(t *testing.T) {
	path := testutil.CreateProtectedFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, true, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result InfoJSON
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	// Sheet1 is protected
	if len(result.Sheets) < 1 {
		t.Fatal("expected at least 1 sheet")
	}
	sheet1 := result.Sheets[0]
	if !sheet1.Protected {
		t.Errorf("expected Sheet1 to be protected")
	}

	// Sheet2 is not protected
	if len(result.Sheets) < 2 {
		t.Fatal("expected at least 2 sheets")
	}
	sheet2 := result.Sheets[1]
	if sheet2.Protected {
		t.Errorf("expected Sheet2 to not be protected")
	}

	// Formulas field should exist (even if 0)
	if sheet1.Formulas != 0 {
		t.Errorf("expected 0 formulas for Sheet1, got %d", sheet1.Formulas)
	}
}

func TestInfo_JSON_NamedRanges(t *testing.T) {
	path := testutil.CreateNamedRangeFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, true, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result InfoJSON
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("output should be valid JSON: %v\n%s", err, stdout.String())
	}

	if len(result.NamedRanges) != 1 {
		t.Fatalf("expected 1 named range, got %d", len(result.NamedRanges))
	}

	nr := result.NamedRanges[0]
	if nr.Name != "TaxRate" {
		t.Errorf("expected name 'TaxRate', got %q", nr.Name)
	}
	if nr.RefersTo != "Sheet1!$B$1" {
		t.Errorf("expected refers_to 'Sheet1!$B$1', got %q", nr.RefersTo)
	}
}

func TestInfo_Plain_NamedRanges(t *testing.T) {
	path := testutil.CreateNamedRangeFixture(t)

	var stdout bytes.Buffer
	err := runInfo(path, false, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()

	if !strings.Contains(output, "Named Ranges: 1") {
		t.Errorf("expected 'Named Ranges: 1', got:\n%s", output)
	}
	if !strings.Contains(output, "TaxRate -> Sheet1!$B$1") {
		t.Errorf("expected 'TaxRate -> Sheet1!$B$1', got:\n%s", output)
	}
}

func TestIsSheetProtected_LegacyPasswordOnly(t *testing.T) {
	path := testutil.CreateLegacyPasswordProtectedFixture(t)

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open fixture: %v", err)
	}
	defer f.Close()

	if !isSheetProtected(f, "Sheet1") {
		t.Error("sheet protected via legacy password attribute only must be detected as protected")
	}
}
