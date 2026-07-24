package commands

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

const (
	maxBatchJSONBytes = 10 << 20
	maxBatchChanges   = 100_000
)

// BatchChange represents a single cell change in a batch operation.
type BatchChange struct {
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Value   any    `json:"value,omitempty"`
	Formula string `json:"formula,omitempty"`
}

func NewBatchCmd() *cobra.Command {
	var opts batchOptions

	cmd := &cobra.Command{
		Use:   "batch <xlsx-file> <changes.json | ->",
		Short: "Write multiple cell values from a JSON file (use - for stdin)",
		Long: `Apply multiple cell changes from a JSON file in a single operation.

The JSON file must contain an array of change objects:
  [
    {"sheet": "Sheet1", "cell": "A1", "value": "Hello"},
    {"sheet": "Sheet1", "cell": "B2", "value": 42},
    {"sheet": "Sheet1", "cell": "C3", "formula": "=SUM(B1:B2)"}
  ]

Each object requires "sheet" and "cell". Then either:
  "value"   - a string, number, or boolean to write
  "formula" - an Excel formula (with or without leading "=")

A "value" is NEVER interpreted as a formula — use it to write literal
strings that start with "=" (the 'write' command treats those as formulas).

String values are coerced like the 'write' command: a European "1.234,56"
becomes the number 1234.56, "true" becomes a boolean, and so on. A "="-
prefixed string stays literal text. To store a numeric-looking string
verbatim (IDs with leading zeros, "5,94 %"), use 'write --string'.

All referenced sheets are validated before any changes are applied.
If a sheet doesn't exist, the command fails without modifying the file.

Use "-" as the JSON path to read from stdin — useful for piping
from other tools or AI agents.

Use --json for a structured result ({"written", "skipped_locked",
"warnings"}) instead of prose — easier for agents to verify. Formulas
in JSON results are reported without the leading "=" (as stored by
Excel), matching the read and find commands. Formula caches are not
recalculated — run 'validate' (LibreOffice recalc) before reading
computed sums.`,
		Example: `  # Write from a JSON file
  xlsx-tool batch report.xlsx changes.json

  # Pipe from stdin
  echo '[{"sheet":"Sheet1","cell":"A1","value":"Hello"}]' | \
    xlsx-tool batch report.xlsx -

  # Write values and formulas together
  cat <<'EOF' > changes.json
  [
    {"sheet": "Sheet1", "cell": "A1", "value": "Revenue"},
    {"sheet": "Sheet1", "cell": "A2", "value": 50000},
    {"sheet": "Sheet1", "cell": "A3", "value": 30000},
    {"sheet": "Sheet1", "cell": "A4", "formula": "=SUM(A2:A3)"}
  ]
  EOF
  xlsx-tool batch report.xlsx changes.json`,
		Args: cobra.ExactArgs(2),
		RunE: func(cmd *cobra.Command, args []string) error {
			jsonPath := args[1]
			if jsonPath == "-" {
				return runBatchFromReaderOpts(args[0], os.Stdin, opts, os.Stdout)
			}
			return runBatchOpts(args[0], jsonPath, opts, os.Stdout)
		},
	}

	cmd.Flags().BoolVar(&opts.jsonOut, "json", false,
		"Output a structured JSON result instead of plain text")

	return cmd
}

type batchOptions struct {
	jsonOut bool
}

// runBatch is the plain-text entry point (kept for backwards-compatible callers).
func runBatch(xlsxPath, jsonPath string, w io.Writer) error {
	return runBatchOpts(xlsxPath, jsonPath, batchOptions{}, w)
}

func runBatchOpts(xlsxPath, jsonPath string, opts batchOptions, w io.Writer) error {
	jsonFile, err := os.Open(jsonPath)
	if err != nil {
		return fmt.Errorf("open JSON file: %w", err)
	}
	defer jsonFile.Close()

	return runBatchFromReaderOpts(xlsxPath, jsonFile, opts, w)
}

// runBatchFromReader is the plain-text entry point (kept for backwards-compatible callers).
func runBatchFromReader(xlsxPath string, r io.Reader, w io.Writer) error {
	return runBatchFromReaderOpts(xlsxPath, r, batchOptions{}, w)
}

func runBatchFromReaderOpts(xlsxPath string, r io.Reader, opts batchOptions, w io.Writer) error {
	payload, err := io.ReadAll(io.LimitReader(r, maxBatchJSONBytes+1))
	if err != nil {
		return fmt.Errorf("read JSON: %w", err)
	}
	if len(payload) > maxBatchJSONBytes {
		return fmt.Errorf("JSON input exceeds %d bytes", maxBatchJSONBytes)
	}

	// Parse JSON; reject trailing data after the array (a concatenated or
	// malformed payload must never be half-applied silently).
	dec := json.NewDecoder(bytes.NewReader(payload))
	var changes []BatchChange
	if err := dec.Decode(&changes); err != nil {
		return fmt.Errorf("parse JSON: %w", err)
	}
	if _, err := dec.Token(); err != io.EOF {
		return fmt.Errorf("unexpected trailing data after the JSON array")
	}
	if len(changes) > maxBatchChanges {
		return fmt.Errorf("batch contains %d changes; maximum is %d", len(changes), maxBatchChanges)
	}

	if len(changes) == 0 {
		if opts.jsonOut {
			return newWriteResult().emit(w)
		}
		fmt.Fprintln(w, "No changes to apply.")
		return nil
	}

	// RawCellValue is a no-op now that batch no longer recalcs in-process
	// (validate recalcs via LibreOffice); kept for parity with write.go.
	f, err := openWorkbook(xlsxPath, true)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	// Validate ALL changes before applying anything — sheet existence, the
	// value/formula contract, and formula sanity. A failing entry must not
	// leave any change applied.
	for i, ch := range changes {
		idx, err := f.GetSheetIndex(ch.Sheet)
		if err != nil || idx == -1 {
			available := f.GetSheetList()
			return fmt.Errorf(
				"sheet not found: %q (cell %s)\nAvailable sheets: %v",
				ch.Sheet, ch.Cell, available,
			)
		}
		if _, _, err := excelize.CellNameToCoordinates(ch.Cell); err != nil {
			return fmt.Errorf("change %d: invalid cell reference %q: %w", i, ch.Cell, err)
		}
		if ch.Formula == "" && ch.Value == nil {
			return fmt.Errorf(
				"change %d (%s!%s): requires either \"value\" or \"formula\"",
				i, ch.Sheet, ch.Cell,
			)
		}
		if ch.Formula != "" && ch.Value != nil {
			return fmt.Errorf(
				"change %d (%s!%s): requires either \"value\" or \"formula\", not both",
				i, ch.Sheet, ch.Cell,
			)
		}
		switch ch.Value.(type) {
		case nil, string, float64, bool:
			// scalar JSON types (numbers decode as float64)
		default:
			return fmt.Errorf(
				"change %d (%s!%s): \"value\" must be a scalar "+
					"(string, number, boolean), got %T",
				i, ch.Sheet, ch.Cell, ch.Value,
			)
		}
		if ch.Formula != "" {
			if _, err := normalizeFormula(ch.Formula); err != nil {
				return fmt.Errorf("change %d (%s!%s): %w", i, ch.Sheet, ch.Cell, err)
			}
		}
	}

	// Apply changes. Cells that are locked on a protected sheet are refused
	// (never overwritten) — writing a locked formula/total cell corrupts the
	// model. Such cells are skipped and reported; the remaining writes apply.
	applied := make([]BatchChange, 0, len(changes))
	blocked := make([]BatchChange, 0)
	warnings := make([]NumericWarning, 0)
	for _, ch := range changes {
		if isWriteBlocked(f, ch.Sheet, ch.Cell) {
			blocked = append(blocked, ch)
			continue
		}
		if ch.Formula != "" {
			formula, _ := normalizeFormula(ch.Formula) // validated above
			// Clear any existing cached value/type FIRST: SetCellFormula alone
			// does not reset a cell's previous <v>/type — overwriting e.g. a
			// string cell would otherwise leak its old shared-string index
			// through GetCellValue as a plausible-looking wrong number. No
			// recalc happens here (see below), so this cell now correctly
			// reads as empty until `validate` recalcs.
			if err := f.SetCellValue(ch.Sheet, ch.Cell, nil); err != nil {
				return fmt.Errorf("clear cell %s!%s: %w", ch.Sheet, ch.Cell, err)
			}
			if err := f.SetCellFormula(ch.Sheet, ch.Cell, formula); err != nil {
				return fmt.Errorf("set formula %s!%s: %w", ch.Sheet, ch.Cell, err)
			}
		} else {
			// Coerce string values the same way `write` does (European decimal
			// comma, ints, floats, bools). JSON numbers (float64) and bools
			// pass through unchanged. A leading-"=" string stays literal text
			// (parseValue returns it unchanged) — values are never formulas.
			if strVal, isStr := ch.Value.(string); isStr {
				stored, warning := coerceCellString(ch.Sheet, ch.Cell, strVal)
				ch.Value = stored
				if warning != nil {
					warnings = append(warnings, *warning)
				}
			}
			if err := f.SetCellValue(ch.Sheet, ch.Cell, ch.Value); err != nil {
				return fmt.Errorf("set value %s!%s: %w", ch.Sheet, ch.Cell, err)
			}
		}
		// Keep the stored dimension covering the written cell — real Excel
		// files always keep it in sync with content, and displayBounds
		// trusts it for trailing cells that dataBounds alone would trim.
		if err := ensureDimensionContains(f, ch.Sheet, ch.Cell); err != nil {
			return fmt.Errorf("update dimension for %s!%s: %w", ch.Sheet, ch.Cell, err)
		}
		applied = append(applied, ch)
	}

	// No recalc here — deliberately. Dependent formula caches stay stale
	// until `validate` runs (it recalcs via LibreOffice: correct semantics,
	// ~75x faster than excelize on large workbooks). FullCalcOnLoad makes
	// plain Excel/LibreOffice recompute on open as the safety net.
	forceRecalc := true
	if err := f.SetCalcProps(&excelize.CalcPropsOptions{FullCalcOnLoad: &forceRecalc}); err != nil {
		return fmt.Errorf("set calc props: %w", err)
	}

	// Save
	if err := saveAtomic(f, xlsxPath); err != nil {
		return fmt.Errorf("save file: %w", err)
	}

	// Output summary
	if opts.jsonOut {
		result := newWriteResult()
		for _, ch := range applied {
			cell := writeResultCell{Sheet: ch.Sheet, Cell: ch.Cell}
			if ch.Formula != "" {
				cell.Formula, _ = normalizeFormula(ch.Formula) // validated above
			} else {
				cell.Value = ch.Value
			}
			result.Written = append(result.Written, cell)
		}
		for _, ch := range blocked {
			result.SkippedLocked = append(result.SkippedLocked,
				skippedLockedCell(f, ch.Sheet, ch.Cell))
		}
		result.Warnings = warnings
		return result.emit(w)
	}

	fmt.Fprintf(w, "Written: %d cells\n", len(applied))
	for _, ch := range applied {
		if ch.Formula != "" {
			fmt.Fprintf(w, "  - %s!%s = =%s\n", ch.Sheet, ch.Cell, ch.Formula)
		} else {
			fmt.Fprintf(w, "  - %s!%s = %v\n", ch.Sheet, ch.Cell, ch.Value)
		}
	}
	if len(blocked) > 0 {
		fmt.Fprintf(w, "Skipped (locked on protected sheet — not written): %d cells\n", len(blocked))
		for _, ch := range blocked {
			fmt.Fprintf(w, "  - %s!%s\n", ch.Sheet, ch.Cell)
		}
		fmt.Fprintln(w, "These cells are protected formula/total cells. They compute "+
			"themselves from the input cells — do not try to unlock them or edit the XML.")
	}
	writeNumericWarnings(w, warnings)

	return nil
}
