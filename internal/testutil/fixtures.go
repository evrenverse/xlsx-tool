package testutil

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// CreateSimpleFixture creates a small test xlsx with known content.
func CreateSimpleFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "simple.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Value")
	f.SetCellValue("Sheet1", "A2", "Alpha")
	f.SetCellValue("Sheet1", "B2", 10)
	f.SetCellValue("Sheet1", "A3", "Beta")
	f.SetCellValue("Sheet1", "B3", 20)
	f.SetCellValue("Sheet1", "A4", "Total")
	f.SetCellFormula("Sheet1", "B4", "SUM(B2:B3)")

	// Dimension must be set explicitly for programmatic fixtures.
	// Real Excel files always have correct dimensions.
	f.SetSheetDimension("Sheet1", "A1:B4")

	idx, _ := f.NewSheet("Data")
	_ = idx
	f.SetCellValue("Data", "A1", "ID")
	f.SetCellValue("Data", "B1", "Amount")
	f.SetCellValue("Data", "A2", 1)
	f.SetCellValue("Data", "B2", 99.5)
	f.SetSheetDimension("Data", "A1:B2")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateSimpleFixtureWithCachedTotal mirrors CreateSimpleFixture but with
// B4's formula pre-patched to carry a cached result of 30, as if a recalc
// (LibreOffice/xlsx-recalc via `validate`, or a real spreadsheet app) had
// already run. batch/write no longer recalc in-process, so tests that need
// a real cached formula value to exercise read use this instead of relying
// on write's removed recalc side effect.
func CreateSimpleFixtureWithCachedTotal(t *testing.T) string {
	t.Helper()
	basePath := CreateSimpleFixture(t)

	path := filepath.Join(t.TempDir(), "simple_cached.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="B4" t="str"><f>SUM(B2:B3)</f></c>`,
		`<c r="B4"><f>SUM(B2:B3)</f><v>30</v></c>`,
	)
	return path
}

// CreateFormattedFixture creates a fixture with formatting to verify preservation.
func CreateFormattedFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "formatted.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	boldStyle, _ := f.NewStyle(&excelize.Style{
		Font: &excelize.Font{Bold: true, Size: 14, Color: "FF0000"},
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FFFF00"}},
	})

	f.SetCellValue("Sheet1", "A1", "Header")
	f.SetCellStyle("Sheet1", "A1", "A1", boldStyle)
	f.SetCellValue("Sheet1", "A2", "Data")
	f.SetCellValue("Sheet1", "B2", 42)
	f.SetSheetDimension("Sheet1", "A1:B2")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateNumberFormatFixture creates a fixture exercising number formats so the
// raw stored value and the formatted display value differ.
// Layout:
//
//	A1: 0.225 with format "0.00%"   → display "22.50%"
//	A2: 0.015 with format "0.00%"   → display "1.50%"
//	A3: 42    (plain number, no number format → display == raw)
//	A4: "00123" (text → display == raw, must NOT be coerced to a number)
//	A5: =A1+A2 with format "0.00%" and cached value 0.24 → display "24.00%"
//	A6: 9007199254740993 (integer just beyond 2^53 — must stay exact, not …992)
//
// Excelize does not compute or cache formula results for programmatic files, so
// the A5 cached <v> is injected via raw-XML patching (like the error fixtures).
func CreateNumberFormatFixture(t *testing.T) string {
	t.Helper()

	f := excelize.NewFile()
	defer f.Close()

	pctStyle, err := f.NewStyle(&excelize.Style{CustomNumFmt: strPtr("0.00%")})
	if err != nil {
		t.Fatalf("create percent style: %v", err)
	}

	f.SetCellValue("Sheet1", "A1", 0.225)
	f.SetCellStyle("Sheet1", "A1", "A1", pctStyle)
	f.SetCellValue("Sheet1", "A2", 0.015)
	f.SetCellStyle("Sheet1", "A2", "A2", pctStyle)
	f.SetCellValue("Sheet1", "A3", 42)
	f.SetCellValue("Sheet1", "A4", "00123")
	f.SetCellFormula("Sheet1", "A5", "A1+A2")
	f.SetCellStyle("Sheet1", "A5", "A5", pctStyle)
	// Integer one past 2^53 — float64 cannot represent it exactly.
	f.SetCellValue("Sheet1", "A6", int64(9007199254740993))
	f.SetSheetDimension("Sheet1", "A1:A6")

	basePath := filepath.Join(t.TempDir(), "base_numfmt.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatalf("create fixture: %v", err)
	}

	// Inject the cached formula result (0.24) and drop t="str" so A5 is treated
	// as a numeric formula cell. The style attribute (s="N") is preserved by
	// matching only the type/value portion.
	path := filepath.Join(t.TempDir(), "numfmt.xlsx")
	patchXlsxCell(t, basePath, path,
		` t="str"><f>A1+A2</f></c>`,
		`><f>A1+A2</f><v>0.24</v></c>`,
	)
	return path
}

func strPtr(s string) *string { return &s }

// CreateFillFixture creates a fixture with fill colors to test fill detection.
// Layout:
//
//	A1: "Feld"  (header, no fill)
//	A2: ""      (empty, yellow fill → fill target)
//	A3: "Fest"  (value, green fill)
//	B1: "Wert"  (header, no fill)
//	B2: ""      (empty, yellow fill → fill target)
//	B3: =A3     (formula, no fill)
func CreateFillFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "fill.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	yellowFill, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"FFFF00"}},
	})
	greenFill, _ := f.NewStyle(&excelize.Style{
		Fill: excelize.Fill{Type: "pattern", Pattern: 1, Color: []string{"00FF00"}},
	})

	f.SetCellValue("Sheet1", "A1", "Feld")
	f.SetCellValue("Sheet1", "B1", "Wert")

	// Empty cells with yellow fill (fill targets)
	// SetCellValue with "" ensures the cell exists in the XML dimension.
	f.SetCellValue("Sheet1", "A2", "")
	f.SetCellStyle("Sheet1", "A2", "A2", yellowFill)
	f.SetCellValue("Sheet1", "B2", "")
	f.SetCellStyle("Sheet1", "B2", "B2", yellowFill)

	// Value with green fill
	f.SetCellValue("Sheet1", "A3", "Fest")
	f.SetCellStyle("Sheet1", "A3", "A3", greenFill)

	// Formula without fill
	f.SetCellFormula("Sheet1", "B3", "A3")

	f.SetSheetDimension("Sheet1", "A1:B3")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateChainedFormulaFixture creates a fixture with chained formula dependencies.
// Layout:
//
//	A1: 10
//	A2: 20
//	A3: =SUM(A1:A2)   (depends on A1, A2)
//	A4: =A3*2         (depends on A3, which depends on A1, A2)
func CreateChainedFormulaFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "chained.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 20)
	f.SetCellFormula("Sheet1", "A3", "SUM(A1:A2)")
	f.SetCellFormula("Sheet1", "A4", "A3*2")
	f.SetSheetDimension("Sheet1", "A1:A4")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateErrorFixture creates a fixture with a division-by-zero formula error.
// Layout:
//
//	A1: 10
//	A2: 0
//	A3: =A1/A2  (produces #DIV/0! after recalc)
func CreateErrorFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellFormula("Sheet1", "A3", "A1/A2")
	f.SetSheetDimension("Sheet1", "A1:A3")

	path := filepath.Join(t.TempDir(), "errors.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

// CreateMultiErrorFixture creates a fixture with two different formula error types.
// Layout:
//
//	A1: "hello"  (text)
//	A2: 0
//	A3: =A2/A2        (produces #DIV/0! — zero divided by zero)
//	A4: =VALUE(A1)    (produces #VALUE! — can't parse "hello" as number)
func CreateMultiErrorFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "hello")
	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellFormula("Sheet1", "A3", "A2/A2")
	f.SetCellFormula("Sheet1", "A4", "VALUE(A1)")
	f.SetSheetDimension("Sheet1", "A1:A4")

	path := filepath.Join(t.TempDir(), "multi_error.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	return path
}

// CreateProtectedFixture creates a fixture with a protected sheet.
// Layout:
//
//	Sheet1: A1="Data", A2=42 (protected)
//	Sheet2: A1="Open" (not protected)
func CreateProtectedFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "protected.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Data")
	f.SetCellValue("Sheet1", "A2", 42)
	f.SetSheetDimension("Sheet1", "A1:A2")

	if err := f.ProtectSheet("Sheet1", &excelize.SheetProtectionOptions{
		Password:          "test",
		SelectLockedCells: true,
	}); err != nil {
		t.Fatalf("protect sheet: %v", err)
	}

	idx, _ := f.NewSheet("Sheet2")
	_ = idx
	f.SetCellValue("Sheet2", "A1", "Open")
	f.SetSheetDimension("Sheet2", "A1:A1")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateProtectedFormulaFixture creates a fixture with a protected sheet that
// contains both a plain locked cell and a locked formula cell.
// Layout:
//
//	Sheet1: A1=10 (plain locked), A2==A1*2 (formula locked, computes 20), protected
//	Sheet2: A1="Open" (not protected)
func CreateProtectedFormulaFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "protected_formula.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellFormula("Sheet1", "A2", "A1*2")
	f.SetSheetDimension("Sheet1", "A1:A2")

	if err := f.ProtectSheet("Sheet1", &excelize.SheetProtectionOptions{
		Password:          "test",
		SelectLockedCells: true,
	}); err != nil {
		t.Fatalf("protect sheet: %v", err)
	}

	idx, _ := f.NewSheet("Sheet2")
	_ = idx
	f.SetCellValue("Sheet2", "A1", "Open")
	f.SetSheetDimension("Sheet2", "A1:A1")

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateDimensionLieFixture creates a single-sheet fixture whose stored
// dimension attribute is patched to dimRef — lying in either direction
// (oversized like "A1:XFD1048576" or undersized like "A1").
//
// Real content/extent (A1:B4):
//
//	A1="Name", B1="Value", A2="Alpha", B2=10, A3="Beta", B3=20,
//	A4="Total", B4==SUM(B2:B3)
func CreateDimensionLieFixture(t *testing.T, dimRef string) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Name")
	f.SetCellValue("Sheet1", "B1", "Value")
	f.SetCellValue("Sheet1", "A2", "Alpha")
	f.SetCellValue("Sheet1", "B2", 10)
	f.SetCellValue("Sheet1", "A3", "Beta")
	f.SetCellValue("Sheet1", "B3", 20)
	f.SetCellValue("Sheet1", "A4", "Total")
	f.SetCellFormula("Sheet1", "B4", "SUM(B2:B3)")
	f.SetSheetDimension("Sheet1", "A1:B4")

	basePath := filepath.Join(t.TempDir(), "base_dim.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(t.TempDir(), "dim_lie.xlsx")
	patchXlsxCell(t, basePath, path,
		`<dimension ref="A1:B4"`,
		`<dimension ref="`+dimRef+`"`,
	)
	return path
}

// CreateLegacyPasswordProtectedFixture creates a fixture whose sheet is
// protected via the legacy password hash attribute with EVERY option
// attribute explicitly blocked ("1") — as produced by some older tools.
// Excelize's GetSheetProtection negates the option attributes, so this
// variant yields all-false option booleans and an empty AlgorithmName:
// invisible to a heuristic that only checks those fields.
//
// Excelize cannot create this variant natively, so the raw XML is patched.
//
// Layout:
//
//	Sheet1: A1="Data", A2=42 (locked by Excel default, sheet protected)
func CreateLegacyPasswordProtectedFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Data")
	f.SetCellValue("Sheet1", "A2", 42)
	f.SetSheetDimension("Sheet1", "A1:A2")

	basePath := filepath.Join(t.TempDir(), "base_legacy.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatal(err)
	}

	path := filepath.Join(t.TempDir(), "legacy_protected.xlsx")
	patchXlsxCell(t, basePath, path,
		"</sheetData>",
		`</sheetData><sheetProtection password="83AF" sheet="1"`+
			` autoFilter="1" deleteColumns="1" deleteRows="1" formatCells="1"`+
			` formatColumns="1" formatRows="1" insertColumns="1"`+
			` insertHyperlinks="1" insertRows="1" objects="1" pivotTables="1"`+
			` scenarios="1" selectLockedCells="1" selectUnlockedCells="1" sort="1"/>`,
	)
	return path
}

// CreateNamedRangeFixture creates a fixture with defined names (named ranges).
// Layout:
//
//	Sheet1: A1="Rate", B1=0.19
//	Named ranges: TaxRate -> Sheet1!$B$1
func CreateNamedRangeFixture(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	path := filepath.Join(dir, "named.xlsx")

	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "Rate")
	f.SetCellValue("Sheet1", "B1", 0.19)
	f.SetSheetDimension("Sheet1", "A1:B1")

	if err := f.SetDefinedName(&excelize.DefinedName{
		Name:     "TaxRate",
		RefersTo: "Sheet1!$B$1",
	}); err != nil {
		t.Fatalf("set defined name: %v", err)
	}

	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateCachedErrorFixture creates a fixture where a formula has a cached error value.
// This simulates a file that was previously opened in Excel or recalculated,
// where the formula result was an error.
//
// Excelize does not natively support creating error-type cells (t="e"),
// so we patch the raw XML after creation.
//
// Layout:
//
//	A1: 10
//	A2: 0
//	A3: =A1/A2 with cached value "#DIV/0!" (cell type "e")
func CreateCachedErrorFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", 10)
	f.SetCellValue("Sheet1", "A2", 0)
	f.SetCellFormula("Sheet1", "A3", "A1/A2")
	f.SetSheetDimension("Sheet1", "A1:A3")

	basePath := filepath.Join(t.TempDir(), "base_error.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatal(err)
	}

	// Patch XML: change A3 from t="str" (formula) to t="e" (error) with cached value
	path := filepath.Join(t.TempDir(), "cached_error.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="A3" t="str"><f>A1/A2</f></c>`,
		`<c r="A3" t="e"><f>A1/A2</f><v>#DIV/0!</v></c>`,
	)
	return path
}

// CreateValidateBaselineFixture creates a clean baseline workbook for the
// validate command: a column of numbers feeding a live sum formula.
// Layout (Sheet1):
//
//	B1: 2
//	B2: 3
//	B3: =B1+B2   (computes 5, no error)
//
// Programmatic fixtures carry no cached formula result — that is fine: this
// fixture is only ever used as the read-only baseline (or as a clean,
// error-free produced file), and validate reads baseline formulas by
// presence/literal-value only, never by evaluating them.
func CreateValidateBaselineFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "B1", 2)
	f.SetCellValue("Sheet1", "B2", 3)
	f.SetCellFormula("Sheet1", "B3", "B1+B2")
	f.SetSheetDimension("Sheet1", "A1:B3")

	path := filepath.Join(t.TempDir(), "validate_baseline.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidateProducedNewErrFixture mirrors the baseline's B3 address but an
// edit turned B1 into text and the B3 formula needs a number, so B3 recalcs to
// a genuine #VALUE! (Excel's VALUE("x") raises #VALUE!). validate reads only
// the CACHED value of a formula cell (LibreOffice recalcs before scanning),
// so the cached error is injected via raw XML patching — the same technique
// CreateCachedErrorFixture uses — to simulate what xlsx-recalc would persist.
// Layout (Sheet1):
//
//	B1: "x"          (text — was 2 in baseline)
//	B2: 3
//	B3: =VALUE(B1) with cached value "#VALUE!" (cell type "e")
func CreateValidateProducedNewErrFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "B1", "x")
	f.SetCellValue("Sheet1", "B2", 3)
	f.SetCellFormula("Sheet1", "B3", "VALUE(B1)")
	f.SetSheetDimension("Sheet1", "A1:B3")

	basePath := filepath.Join(t.TempDir(), "base_validate_produced_newerr.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatalf("create fixture: %v", err)
	}

	path := filepath.Join(t.TempDir(), "validate_produced_newerr.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="B3" t="str"><f>VALUE(B1)</f></c>`,
		`<c r="B3" t="e"><f>VALUE(B1)</f><v>#VALUE!</v></c>`,
	)
	return path
}

// CreateValidateProducedFrozenFixture mirrors the baseline but the B3 formula
// was replaced by the literal 5 (the agent froze a live formula).
// Layout (Sheet1):
//
//	B1: 2
//	B2: 3
//	B3: 5       (literal — was =B1+B2 in baseline)
func CreateValidateProducedFrozenFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "B1", 2)
	f.SetCellValue("Sheet1", "B2", 3)
	f.SetCellValue("Sheet1", "B3", 5)
	f.SetSheetDimension("Sheet1", "A1:B3")

	path := filepath.Join(t.TempDir(), "validate_produced_frozen.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidateBaselineInputFixture creates a baseline template with one true
// fill-target plus decoys that must NOT be flagged.
// Layout (Sheet1):
//
//	A1:B1  empty, unlocked, MERGED  -> excluded (merged-range member)
//	C5     empty, unlocked          -> the only fill-target
//	D5     empty, LOCKED            -> excluded (locked)
//	E5     =C5                      -> excluded (formula cell)
//	C7     "anchor"                 -> keeps the data bounds past C5/D5/E5 so
//	                                   the empty target cells are iterated
func CreateValidateBaselineInputFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	unlocked, err := f.NewStyle(&excelize.Style{
		Protection: &excelize.Protection{Locked: false},
	})
	if err != nil {
		t.Fatalf("create unlocked style: %v", err)
	}

	// Merged, unlocked, empty range — members must be excluded.
	if err := f.MergeCell("Sheet1", "A1", "B1"); err != nil {
		t.Fatalf("merge cells: %v", err)
	}
	f.SetCellValue("Sheet1", "A1", "")
	f.SetCellStyle("Sheet1", "A1", "B1", unlocked)

	// The real fill-target: unlocked, empty, non-formula.
	f.SetCellValue("Sheet1", "C5", "")
	f.SetCellStyle("Sheet1", "C5", "C5", unlocked)

	// Locked empty decoy (default style is locked) — excluded.
	f.SetCellValue("Sheet1", "D5", "")

	// Formula decoy — excluded.
	f.SetCellFormula("Sheet1", "E5", "C5")

	// Anchor so the data bounds extend past row 5 / column E and the empty
	// target cells fall inside the iterated region (GetRows trims trailing
	// blanks, so a value below keeps C5 reachable).
	f.SetCellValue("Sheet1", "C7", "anchor")

	f.SetSheetDimension("Sheet1", "A1:E7")

	path := filepath.Join(t.TempDir(), "validate_baseline_input.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidateProducedUnfilledFixture mirrors the input baseline but the C5
// fill-target was left empty (the agent never filled it).
func CreateValidateProducedUnfilledFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	unlocked, err := f.NewStyle(&excelize.Style{
		Protection: &excelize.Protection{Locked: false},
	})
	if err != nil {
		t.Fatalf("create unlocked style: %v", err)
	}

	if err := f.MergeCell("Sheet1", "A1", "B1"); err != nil {
		t.Fatalf("merge cells: %v", err)
	}
	f.SetCellValue("Sheet1", "A1", "")
	f.SetCellStyle("Sheet1", "A1", "B1", unlocked)

	// C5 STILL empty (unfilled).
	f.SetCellValue("Sheet1", "C5", "")
	f.SetCellStyle("Sheet1", "C5", "C5", unlocked)

	f.SetCellValue("Sheet1", "D5", "")
	f.SetCellFormula("Sheet1", "E5", "C5")
	f.SetCellValue("Sheet1", "C7", "anchor")
	f.SetSheetDimension("Sheet1", "A1:E7")

	path := filepath.Join(t.TempDir(), "validate_produced_unfilled.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidateProtectedUnfilledFixture creates a workbook whose Sheet1 is
// PROTECTED and carries targetCount unlocked, empty, non-formula input slots
// (C5, C6, ... down column C). An anchor value is written one row below the
// last target so GetRows does not trim the target rows away (mirrors the
// anchor technique in CreateValidateBaselineInputFixture).
//
// Calling this twice with the same targetCount produces two independent
// files with identical content — suitable as a (baseline, produced) pair
// where the targets are still empty in produced (unfilled).
func CreateValidateProtectedUnfilledFixture(t *testing.T, targetCount int) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	unlocked, err := f.NewStyle(&excelize.Style{
		Protection: &excelize.Protection{Locked: false},
	})
	if err != nil {
		t.Fatalf("create unlocked style: %v", err)
	}

	const startRow = 5
	for i := 0; i < targetCount; i++ {
		cell := fmt.Sprintf("C%d", startRow+i)
		f.SetCellValue("Sheet1", cell, "")
		f.SetCellStyle("Sheet1", cell, cell, unlocked)
	}
	anchorRow := startRow + targetCount
	anchorCell := fmt.Sprintf("C%d", anchorRow)
	f.SetCellValue("Sheet1", anchorCell, "anchor")
	f.SetSheetDimension("Sheet1", fmt.Sprintf("A1:C%d", anchorRow))

	if err := f.ProtectSheet("Sheet1", &excelize.SheetProtectionOptions{
		Password:          "test",
		SelectLockedCells: true,
	}); err != nil {
		t.Fatalf("protect sheet: %v", err)
	}

	path := filepath.Join(t.TempDir(), fmt.Sprintf("validate_protected_unfilled_%d.xlsx", targetCount))
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidateBaselineGateFixture creates the clean baseline of the
// gate-cascade shape: a gate IF that holds (A1==B1) so the total computes a
// number with no error.
// Layout (Sheet1):
//
//	A1: 33
//	B1: 33          (A1==B1 -> gate passes)
//	D1: 7
//	C1: =IF(A1=B1,D1,"")  (gate -> 7)
//	E1: =VALUE(C1)        (total -> 7, no error)
func CreateValidateBaselineGateFixture(t *testing.T) string {
	t.Helper()
	return writeGateFixture(t, "validate_baseline_gate.xlsx", 33)
}

// CreateValidateProducedGateFixture creates the broken produced variant: A1 was
// changed to 95 so the gate (A1=B1) fails, C1 becomes "" and the VALUE() total
// cascades to a real #VALUE!. validate reads only the CACHED value of a
// formula cell (LibreOffice recalcs before scanning), so E1's cached error is
// injected via raw XML patching to simulate what xlsx-recalc would persist.
// Layout (Sheet1):
//
//	A1: 95          (A1!=B1 -> gate fails)
//	B1: 33
//	D1: 7
//	C1: =IF(A1=B1,D1,"")  (gate -> "")
//	E1: =VALUE(C1) with cached value "#VALUE!" (cell type "e")
func CreateValidateProducedGateFixture(t *testing.T) string {
	t.Helper()
	basePath := writeGateFixture(t, "base_validate_produced_gate.xlsx", 95)

	path := filepath.Join(t.TempDir(), "validate_produced_gate.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="E1" t="str"><f>VALUE(C1)</f></c>`,
		`<c r="E1" t="e"><f>VALUE(C1)</f><v>#VALUE!</v></c>`,
	)
	return path
}

// writeGateFixture builds the shared gate layout with A1 set to a1.
func writeGateFixture(t *testing.T, name string, a1 int) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", a1)
	f.SetCellValue("Sheet1", "B1", 33)
	f.SetCellValue("Sheet1", "D1", 7)
	f.SetCellFormula("Sheet1", "C1", `IF(A1=B1,D1,"")`)
	f.SetCellFormula("Sheet1", "E1", "VALUE(C1)")
	f.SetSheetDimension("Sheet1", "A1:E1")

	path := filepath.Join(t.TempDir(), name)
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("create fixture: %v", err)
	}
	return path
}

// CreateValidatePlaceholderFixture builds the real bid-template shape: a
// submittable formula cell whose IF false-branch returns the literal string
// "#WERT!" when an unfilled gating input fails the gate. LibreOffice's recalc
// caches the literal "#WERT!" (not a recognised Excel error value), so
// scanFormulaErrors misses it — yet it is a real, fixable defect: filling the
// gating input flips the gate and the placeholder disappears.
//
// Layout (Sheet1):
//
//	A16: a16              gating input (unlocked, non-formula); empty by default
//	B20: 4                a numeric operand for the true branch
//	D20: 5                a numeric operand for the true branch
//	C20: =IF(A16>15, B20*D20, "#WERT!")   placeholder when A16 fails the gate
//	G40: =C20             grand total propagating the placeholder
//
// When a16 == "" (empty), A16>15 is false and C20 evaluates to "#WERT!",
// which propagates to G40. validate reads only the CACHED value of a formula
// cell, so when the gate fails (a16 == nil) C20's and G40's cached "#WERT!"
// results are injected via raw XML patching to simulate what xlsx-recalc
// would persist. Filling A16 with a value > 15 makes the gate pass and C20
// (and G40) compute a number — no cache injection needed there since an
// uncached, non-error result is correctly left un-flagged either way.
func CreateValidatePlaceholderFixture(t *testing.T, a16 any) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	unlocked, err := f.NewStyle(&excelize.Style{
		Protection: &excelize.Protection{Locked: false},
	})
	if err != nil {
		t.Fatalf("create unlocked style: %v", err)
	}

	// The gating input: unlocked so the agent can fill it. Empty by default
	// (a16 == nil leaves it blank).
	if a16 != nil {
		f.SetCellValue("Sheet1", "A16", a16)
	} else {
		f.SetCellValue("Sheet1", "A16", "")
	}
	f.SetCellStyle("Sheet1", "A16", "A16", unlocked)

	f.SetCellValue("Sheet1", "B20", 4)
	f.SetCellValue("Sheet1", "D20", 5)
	f.SetCellFormula("Sheet1", "C20", `IF(A16>15,B20*D20,"#WERT!")`)
	f.SetCellFormula("Sheet1", "G40", "C20")
	f.SetSheetDimension("Sheet1", "A1:G40")

	basePath := filepath.Join(t.TempDir(), "base_validate_placeholder.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatalf("create fixture: %v", err)
	}

	if a16 != nil {
		// Gate passes: no placeholder produced, nothing to cache.
		return basePath
	}

	// Gate fails: inject the cached "#WERT!" result LibreOffice would persist
	// for both the placeholder cell and the total that propagates it.
	path := filepath.Join(t.TempDir(), "validate_placeholder.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="C20" t="str"><f>IF(A16&gt;15,B20*D20,&#34;#WERT!&#34;)</f></c>`,
		`<c r="C20" t="str"><f>IF(A16&gt;15,B20*D20,&#34;#WERT!&#34;)</f><v>#WERT!</v></c>`,
	)
	patchXlsxCell(t, path, path,
		`<c r="G40" t="str"><f>C20</f></c>`,
		`<c r="G40" t="str"><f>C20</f><v>#WERT!</v></c>`,
	)
	return path
}

// CreateValidateExternalErrorFixture builds an error cell whose operand chain
// reaches ONLY a locked, template-owned input — there is no agent-fillable leaf,
// so the error must be classified `external` (informational, does not block ok).
//
// Layout (Sheet1, protected):
//
//	A1: "x"          template-owned text input, LOCKED (Excel default)
//	B1: =VALUE(A1) with cached value "#VALUE!" (cell type "e") — A1 is
//	                 non-numeric text the agent cannot change
//	C1: "anchor"     keeps the data bounds past A1/B1
//
// The sheet is protected so the locked A1 is genuinely template-owned; the chain
// from B1 reaches only A1 (locked) -> external. validate reads only the
// CACHED value of a formula cell, so B1's cached error is injected via raw
// XML patching to simulate what xlsx-recalc would persist.
func CreateValidateExternalErrorFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()

	f.SetCellValue("Sheet1", "A1", "x")
	f.SetCellFormula("Sheet1", "B1", "VALUE(A1)")
	f.SetCellValue("Sheet1", "C1", "anchor")
	f.SetSheetDimension("Sheet1", "A1:C1")

	if err := f.ProtectSheet("Sheet1", &excelize.SheetProtectionOptions{
		Password:          "test",
		SelectLockedCells: true,
	}); err != nil {
		t.Fatalf("protect sheet: %v", err)
	}

	basePath := filepath.Join(t.TempDir(), "base_validate_external_error.xlsx")
	if err := f.SaveAs(basePath); err != nil {
		t.Fatalf("create fixture: %v", err)
	}

	path := filepath.Join(t.TempDir(), "validate_external_error.xlsx")
	patchXlsxCell(t, basePath, path,
		`<c r="B1" t="str"><f>VALUE(A1)</f></c>`,
		`<c r="B1" t="e"><f>VALUE(A1)</f><v>#VALUE!</v></c>`,
	)
	return path
}

// patchXlsxCell patches the sheet1.xml inside an xlsx file by replacing
// oldXML with newXML. Used to create test fixtures with cell types that
// excelize cannot produce natively (e.g. error cells).
func patchXlsxCell(t *testing.T, srcPath, dstPath, oldXML, newXML string) {
	t.Helper()
	data, err := os.ReadFile(srcPath)
	if err != nil {
		t.Fatalf("read base xlsx: %v", err)
	}
	r, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("open zip: %v", err)
	}

	var buf bytes.Buffer
	w := zip.NewWriter(&buf)

	for _, zf := range r.File {
		rc, err := zf.Open()
		if err != nil {
			t.Fatalf("open zip entry %s: %v", zf.Name, err)
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatalf("read zip entry %s: %v", zf.Name, err)
		}

		if zf.Name == "xl/worksheets/sheet1.xml" {
			s := strings.Replace(string(content), oldXML, newXML, 1)
			content = []byte(s)
		}

		fw, err := w.CreateHeader(&zf.FileHeader)
		if err != nil {
			t.Fatalf("create zip entry %s: %v", zf.Name, err)
		}
		if _, err := fw.Write(content); err != nil {
			t.Fatalf("write zip entry %s: %v", zf.Name, err)
		}
	}

	if err := w.Close(); err != nil {
		t.Fatalf("close zip: %v", err)
	}
	if err := os.WriteFile(dstPath, buf.Bytes(), 0644); err != nil {
		t.Fatalf("write patched xlsx: %v", err)
	}
}

// ReadCellValue reads a cell value from an xlsx file (for test assertions).
func ReadCellValue(t *testing.T, path, sheet, cell string) string {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open file for assertion: %v", err)
	}
	defer f.Close()

	val, err := f.GetCellValue(sheet, cell)
	if err != nil {
		t.Fatalf("read cell %s!%s: %v", sheet, cell, err)
	}
	return val
}

// ReadCellFormula reads a cell formula from an xlsx file (for test assertions).
func ReadCellFormula(t *testing.T, path, sheet, cell string) string {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open file for assertion: %v", err)
	}
	defer f.Close()

	formula, err := f.GetCellFormula(sheet, cell)
	if err != nil {
		t.Fatalf("read formula %s!%s: %v", sheet, cell, err)
	}
	return formula
}

// GetCellStyleFont reads font properties of a cell (for formatting preservation tests).
func GetCellStyleFont(t *testing.T, path, sheet, cell string) *excelize.Font {
	t.Helper()
	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open file: %v", err)
	}
	defer f.Close()

	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		t.Fatalf("get style: %v", err)
	}

	style, err := f.GetStyle(styleID)
	if err != nil || style == nil {
		return nil
	}
	return style.Font
}

// FileExists checks if a file exists at the given path.
func FileExists(path string) bool {
	_, err := os.Stat(path)
	return err == nil
}
