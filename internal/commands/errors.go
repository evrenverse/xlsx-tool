package commands

import (
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

// FormulaError represents a cell whose formula evaluates to an Excel error value.
type FormulaError struct {
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Error   string `json:"error"`
	Formula string `json:"formula"`
}

// errorValues contains all standard Excel formula error strings.
var errorValues = []string{
	"#VALUE!", "#DIV/0!", "#REF!", "#NAME?", "#NULL!", "#NUM!", "#N/A",
}

// isErrorValue returns true if the string matches a known Excel error value.
func isErrorValue(s string) bool {
	for _, e := range errorValues {
		if s == e {
			return true
		}
	}
	return false
}

// scanFormulaErrors iterates all cells in all sheets and returns a list of
// formula cells whose CACHED value is a known Excel error value.
//
// Prerequisite: the file must already have been recalced by LibreOffice
// (xlsx-recalc) — see runValidate, which always recalcs the produced
// workbook in place before calling this. This function never evaluates a
// formula itself; it only reads whatever value is already cached.
func scanFormulaErrors(f *excelize.File) []FormulaError {
	var errors []FormulaError

	for _, sheet := range f.GetSheetList() {
		// Iterate actual data bounds, never the stored dimension — it can lie
		// in both directions (oversized: per-cell hang; undersized: skipped
		// cells); formula cells always exist in sheetData regardless.
		endCol, endRow, err := dataBounds(f, sheet)
		if err != nil || endCol == 0 || endRow == 0 {
			continue
		}

		for r := 1; r <= endRow; r++ {
			for c := 1; c <= endCol; c++ {
				colName, _ := excelize.ColumnNumberToName(c)
				cellRef := fmt.Sprintf("%s%d", colName, r)

				formula, _ := f.GetCellFormula(sheet, cellRef)
				if formula == "" {
					continue
				}

				// Read the CACHED value (LibreOffice recalced the file before
				// this scan) — never evaluate. LibreOffice stores error cells
				// as t="e" with the error string as the cached value.
				result, err := f.GetCellValue(sheet, cellRef, excelize.Options{RawCellValue: true})
				if err != nil {
					continue
				}
				if isErrorValue(result) {
					errors = append(errors, FormulaError{
						Sheet:   sheet,
						Cell:    cellRef,
						Error:   result,
						Formula: formula,
					})
				}
			}
		}
	}

	return errors
}

// writeNumericWarnings prints near-numeric-stored-as-text warnings to w.
// The message is taken verbatim from each NumericWarning so plain-text and
// JSON output share the single canonical wording (numericTextWarningMessage).
func writeNumericWarnings(w io.Writer, warnings []NumericWarning) {
	for _, nw := range warnings {
		fmt.Fprintf(w, "Warning: %s!%s: value %q %s\n", nw.Sheet, nw.Cell, nw.Value, nw.Message)
	}
}
