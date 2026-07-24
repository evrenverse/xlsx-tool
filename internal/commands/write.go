package commands

import (
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

type writeOptions struct {
	jsonOut  bool
	asString bool
}

func NewWriteCmd() *cobra.Command {
	var opts writeOptions

	cmd := &cobra.Command{
		Use:   "write <file> <sheet> <cell> <value>",
		Short: "Write a value or formula to a single cell",
		Long: `Set the value of a single cell and save the file.

The value is auto-detected as integer, float, boolean, formula, or string.
Beware the coercion edge cases — the echo always shows the STORED value:
  "42"          -> integer 42
  "00123"       -> integer 123 (leading zeros lost!)
  "3.14"        -> float 3.14
  "1e5"         -> float 100000
  "0x1p-2"      -> float 0.25
  "true", "t"   -> boolean true
  "=SUM(B1:B2)" -> formula (a leading "=" is always written as a formula;
                   no cached value is computed until 'validate' recalcs)
  "anything"    -> string

Use --string to store the value as a literal string with NO coercion and
NO formula interpretation (IDs with leading zeros, texts starting with
"="). batch's explicit "value" field behaves the same way.

Existing formatting, formulas, and other cells are preserved.
For writing multiple cells at once, use the 'batch' command instead.

Use --json for a structured result ({"written", "skipped_locked",
"warnings"}) instead of prose — easier for agents to verify. Formulas
in JSON results are reported without the leading "=" (as stored by
Excel), matching the read and find commands. Formula caches are not
recalculated — run 'validate' (LibreOffice recalc) before reading
computed sums.`,
		Example: `  # Write a string
  xlsx-tool write report.xlsx Sheet1 A1 "Project Name"

  # Write a number (auto-detected)
  xlsx-tool write report.xlsx Sheet1 B2 "1299.99"

  # Write a formula (leading "=" means formula)
  xlsx-tool write report.xlsx Sheet1 B4 "=SUM(B2:B3)"

  # Write to a specific sheet, structured result
  xlsx-tool write report.xlsx "Sales Data" C5 "42" --json`,
		Args: cobra.ExactArgs(4),
		RunE: func(cmd *cobra.Command, args []string) error {
			return runWriteOpts(args[0], args[1], args[2], args[3], opts, os.Stdout)
		},
	}

	cmd.Flags().BoolVar(&opts.jsonOut, "json", false,
		"Output a structured JSON result instead of plain text")
	cmd.Flags().BoolVar(&opts.asString, "string", false,
		"Store the value as a literal string (no type coercion, no formula interpretation)")

	return cmd
}

// runWrite is the plain-text entry point (kept for backwards-compatible callers).
func runWrite(filePath, sheet, cell, value string, w io.Writer) error {
	return runWriteOpts(filePath, sheet, cell, value, writeOptions{}, w)
}

func runWriteOpts(filePath, sheet, cell, value string, opts writeOptions, w io.Writer) error {
	// Validate the cell reference BEFORE any protection check: the lock
	// detection fails safe (treats lookup errors as locked), so an invalid
	// ref on a protected sheet would be misreported as skipped_locked.
	if _, _, err := excelize.CellNameToCoordinates(cell); err != nil {
		return fmt.Errorf("invalid cell reference %q: %w", cell, err)
	}

	// RawCellValue is a no-op now that write no longer recalcs in-process
	// (validate recalcs via LibreOffice); kept for parity with batch.go.
	f, err := openWorkbook(filePath, true)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	// Verify sheet exists
	idx, err := f.GetSheetIndex(sheet)
	if err != nil || idx == -1 {
		available := f.GetSheetList()
		return fmt.Errorf("sheet not found: %q\nAvailable sheets: %v", sheet, available)
	}

	result := newWriteResult()

	// Refuse to overwrite a locked cell on a protected sheet — these are
	// formula/total cells that compute themselves; writing a literal corrupts
	// the model and cascades into formula errors.
	if isWriteBlocked(f, sheet, cell) {
		result.SkippedLocked = append(result.SkippedLocked,
			skippedLockedCell(f, sheet, cell))
		if opts.jsonOut {
			return result.emit(w)
		}
		fmt.Fprintf(w, "Skipped (locked on protected sheet — not written): %s!%s\n", sheet, cell)
		fmt.Fprintln(w, "This is a protected formula/total cell. It computes itself from the "+
			"input cells — do not try to unlock it or edit the XML.")
		return nil
	}

	// A leading "=" always means formula (same behavior as batch's "formula"
	// field) — unless --string forces a literal. With --string the value is
	// stored verbatim: no type coercion, no formula interpretation.
	isFormula := !opts.asString && strings.HasPrefix(value, "=")
	var storedValue any
	if isFormula {
		formula, err := normalizeFormula(value)
		if err != nil {
			return err
		}
		// Clear any existing cached value/type FIRST: SetCellFormula alone
		// does not reset a cell's previous <v>/type — overwriting e.g. a
		// string cell would otherwise leak its old shared-string index
		// through GetCellValue as a plausible-looking wrong number. No
		// recalc happens here (see below), so this cell now correctly reads
		// as empty until `validate` recalcs.
		if err := f.SetCellValue(sheet, cell, nil); err != nil {
			return fmt.Errorf("clear cell %s!%s: %w", sheet, cell, err)
		}
		if err := f.SetCellFormula(sheet, cell, formula); err != nil {
			return fmt.Errorf("set formula %s!%s: %w", sheet, cell, err)
		}
		result.Written = append(result.Written,
			writeResultCell{Sheet: sheet, Cell: cell, Formula: formula})
	} else {
		if opts.asString {
			storedValue = value
		} else {
			// Coerce to number/boolean when possible; collect a warning when the
			// value looks numeric but stayed text (shared with the batch path).
			stored, warning := coerceCellString(sheet, cell, value)
			storedValue = stored
			if warning != nil {
				result.Warnings = append(result.Warnings, *warning)
			}
		}
		if err := f.SetCellValue(sheet, cell, storedValue); err != nil {
			return fmt.Errorf("set cell %s!%s: %w", sheet, cell, err)
		}
		result.Written = append(result.Written,
			writeResultCell{Sheet: sheet, Cell: cell, Value: storedValue})
	}

	// Keep the stored dimension covering the written cell — real Excel
	// files always keep it in sync with content, and displayBounds trusts
	// it for trailing cells that dataBounds alone would trim.
	if err := ensureDimensionContains(f, sheet, cell); err != nil {
		return fmt.Errorf("update dimension for %s!%s: %w", sheet, cell, err)
	}

	// No recalc here — deliberately. Dependent formula caches stay stale
	// until `validate` runs (it recalcs via LibreOffice: correct semantics,
	// ~75x faster than excelize on large workbooks). FullCalcOnLoad makes
	// plain Excel/LibreOffice recompute on open as the safety net.
	forceRecalc := true
	if err := f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &forceRecalc}); err != nil {
		return fmt.Errorf("set calc props: %w", err)
	}

	if err := saveAtomic(f, filePath); err != nil {
		return fmt.Errorf("save file: %w", err)
	}

	if opts.jsonOut {
		return result.emit(w)
	}

	if isFormula {
		fmt.Fprintf(w, "Written: %s!%s = %s\n", sheet, cell, value)
	} else {
		// Echo the STORED value and type so coercion is always visible.
		fmt.Fprintf(w, "Written: %s!%s = %s\n", sheet, cell, describeStored(storedValue))
	}
	writeNumericWarnings(w, result.Warnings)
	return nil
}

// describeStored renders a stored cell value with its detected type, e.g.
// `"text" (string)`, `123 (integer)`, `100000 (float)`, `true (boolean)`.
func describeStored(v any) string {
	switch tv := v.(type) {
	case string:
		return fmt.Sprintf("%q (string)", tv)
	case int64:
		return fmt.Sprintf("%d (integer)", tv)
	case float64:
		return fmt.Sprintf("%v (float)", tv)
	case bool:
		return fmt.Sprintf("%t (boolean)", tv)
	default:
		return fmt.Sprintf("%v", tv)
	}
}

// normalizeFormula strips one leading "=" and rejects degenerate formulas:
// a bare/whitespace-only "=" and a double leading "=" (e.g. "==B1") would
// otherwise be stored silently as uncalculable formula text.
func normalizeFormula(raw string) (string, error) {
	formula := strings.TrimPrefix(raw, "=")
	trimmed := strings.TrimSpace(formula)
	if trimmed == "" {
		return "", fmt.Errorf("invalid formula %q: empty after \"=\"", raw)
	}
	if strings.HasPrefix(trimmed, "=") {
		return "", fmt.Errorf("invalid formula %q: double leading \"=\"", raw)
	}
	return formula, nil
}

// parseValue tries to interpret the string as a number or boolean.
// European notation is supported: when the string contains a decimal
// comma, dots are treated as thousands separators ("1.234,56" → 1234.56).
// Strings with units or percent signs ("24,75 €", "5,94 %") stay strings —
// Excel unit semantics are the caller's job, silent magic would be worse.
func parseValue(s string) any {
	trimmed := strings.TrimSpace(s)
	if i, err := strconv.ParseInt(trimmed, 10, 64); err == nil {
		return i
	}
	if f, err := strconv.ParseFloat(trimmed, 64); err == nil {
		return f
	}
	if strings.Count(trimmed, ",") == 1 {
		normalized := strings.ReplaceAll(trimmed, ".", "")
		normalized = strings.Replace(normalized, ",", ".", 1)
		if f, err := strconv.ParseFloat(normalized, 64); err == nil {
			return f
		}
	}
	if b, err := strconv.ParseBool(trimmed); err == nil {
		return b
	}
	return s
}

// numericTextWarningMessage is the single canonical wording for the
// "near-numeric value stored as text" warning (used in both plain and JSON output).
const numericTextWarningMessage = "looks numeric but was written as TEXT — formulas " +
	"referencing it will break (#VALUE!). Write a plain number (e.g. 23.45 or 23,45) " +
	"or use --string to silence."

// coerceCellString coerces a raw string cell value to a number/bool when possible
// and returns an optional warning when it looks numeric but stayed text.
func coerceCellString(sheet, cell, raw string) (any, *NumericWarning) {
	stored := parseValue(raw)
	if looksNumericButUnparseable(raw) {
		return stored, &NumericWarning{
			Sheet: sheet, Cell: cell, Value: raw, Message: numericTextWarningMessage,
		}
	}
	return stored, nil
}

// looksNumericButUnparseable reports whether the value reads like a number a
// human wrote (digits with separators/units) but did not coerce to one — the
// classic "#VALUE! later" trap. Used to warn at write time.
func looksNumericButUnparseable(s string) bool {
	trimmed := strings.TrimSpace(s)
	if trimmed == "" {
		return false
	}
	if _, ok := parseValue(trimmed).(string); !ok {
		return false // it coerced fine
	}
	hasDigit := false
	for _, r := range trimmed {
		switch {
		case r >= '0' && r <= '9':
			hasDigit = true
		case r == '.' || r == ',' || r == ' ' || r == '%' || r == '€' || r == '-' || r == '+':
			// allowed near-numeric characters
		default:
			return false
		}
	}
	return hasDigit
}
