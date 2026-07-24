package commands

import (
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

func TestScanFormulaErrors_NoErrors(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	f, err := excelize.OpenFile(path, excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("open file: %v", err)
	}
	defer f.Close()

	// The fixture's formula cell has no cached value yet (excelize never
	// computes on write, and neither batch nor write recalc anymore) —
	// scanFormulaErrors must not treat an absent/empty cache as an error.
	errors := scanFormulaErrors(f)
	if len(errors) != 0 {
		t.Errorf("expected 0 formula errors, got %d: %+v", len(errors), errors)
	}
}

func TestScanFormulaErrors_DetectsDivByZero(t *testing.T) {
	// scanFormulaErrors reads only the CACHED value (prerequisite: the file
	// has already been recalced by LibreOffice/xlsx-recalc — see runValidate).
	// It never evaluates a formula itself, so the cached-error fixture
	// (simulating what xlsx-recalc would have written) is used directly.
	path := testutil.CreateCachedErrorFixture(t)

	f, err := excelize.OpenFile(path, excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatalf("open file: %v", err)
	}
	defer f.Close()

	errors := scanFormulaErrors(f)
	if len(errors) != 1 {
		t.Fatalf("expected 1 formula error, got %d: %+v", len(errors), errors)
	}

	fe := errors[0]
	if fe.Sheet != "Sheet1" {
		t.Errorf("expected sheet 'Sheet1', got %q", fe.Sheet)
	}
	if fe.Cell != "A3" {
		t.Errorf("expected cell 'A3', got %q", fe.Cell)
	}
	if fe.Error != "#DIV/0!" {
		t.Errorf("expected error '#DIV/0!', got %q", fe.Error)
	}
	if fe.Formula != "A1/A2" {
		t.Errorf("expected formula 'A1/A2' (no leading =), got %q", fe.Formula)
	}
}
