package commands

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"strconv"
	"strings"

	"github.com/evrenverse/xlsx-tool/internal/format"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

type readOptions struct {
	cellRange    string
	cell         string
	cells        string
	sheets       string
	allSheets    bool
	jsonOut      bool
	confirmLarge bool
}

func NewReadCmd() *cobra.Command {
	var opts readOptions

	cmd := &cobra.Command{
		Use:   "read <file> [sheet]",
		Short: "Read Excel sheet content",
		Long: `Read cell content from one or more sheets and display it.

Output is compact plain text by default. Use --json for structured JSON output.

Modes:
  Full sheet      Read the entire sheet (default).
  Range           Read a rectangular area with --range A1:D10.
  Single cell     Read one cell's value, formula, and fill color with --cell B5.
  Scattered cells Read multiple specific cells in one call with --cells C21,L21,AS21.
  Multiple sheets Read several sheets in one call with --sheets "January,February",
                  or every sheet with --all-sheets (combine with --range to
                  apply the same rectangle to each sheet).

Each invocation opens the whole workbook — for N scattered cells use ONE
--cells call (or --range for a rectangle), and for N sheets use ONE
--sheets/--all-sheets call, never a per-cell or per-sheet loop.

With --cells, JSON output is keyed by cell reference (alphabetical);
plain output preserves input order.

Sheet names containing a comma cannot be addressed via --sheets (the
list is split on commas) — use the positional sheet argument with one
call per sheet, or --all-sheets.

Use "xlsx-tool info" first to discover available sheet names.
Formula cells show the formula AND its cached computed value
(e.g. =SUM(A1:A2) → "30").
Cell background colors (fill) are included when present.`,
		Example: `  # Read a sheet (plain text)
  xlsx-tool read report.xlsx "Sales Data"

  # Read as JSON
  xlsx-tool read report.xlsx "Sales Data" --json

  # Read a range
  xlsx-tool read report.xlsx Sheet1 --range A1:D10

  # Read a single cell
  xlsx-tool read report.xlsx Sheet1 --cell B5

  # Read many scattered cells in ONE invocation (instead of a --cell loop)
  xlsx-tool read report.xlsx Sheet1 --cells C21,L21,AS21

  # Read several sheets in ONE invocation (single JSON document)
  xlsx-tool read report.xlsx --sheets "January,February" --json

  # Read every sheet in the workbook
  xlsx-tool read report.xlsx --all-sheets --json`,
		Args: cobra.RangeArgs(1, 2),
		RunE: func(cmd *cobra.Command, args []string) error {
			sheet := ""
			if len(args) == 2 {
				sheet = args[1]
			}
			return runRead(args[0], sheet, opts, os.Stdout)
		},
	}

	cmd.Flags().StringVar(&opts.cellRange, "range", "", "Rectangular cell range to read (e.g. A1:D10)")
	cmd.Flags().StringVar(&opts.cell, "cell", "", "Single cell reference to read (e.g. B5)")
	cmd.Flags().StringVar(&opts.cells, "cells", "",
		"Comma-separated scattered cell references to read in one call (e.g. C21,L21,AS21)")
	cmd.Flags().StringVar(&opts.sheets, "sheets", "",
		"Comma-separated sheet names to read in one call (e.g. \"January,February\")")
	cmd.Flags().BoolVar(&opts.allSheets, "all-sheets", false,
		"Read every sheet in the workbook (workbook order)")
	cmd.Flags().BoolVar(&opts.jsonOut, "json", false, "Output as JSON instead of plain text")
	cmd.Flags().BoolVar(&opts.confirmLarge, "confirm-large", false,
		"Allow a --sheets/--all-sheets read that spans a very large number of cells")

	return cmd
}

func runRead(filePath, sheet string, opts readOptions, w io.Writer) error {
	if err := validateReadModes(sheet, opts); err != nil {
		return err
	}

	var cellRefs []string
	if opts.cells != "" {
		refs, err := parseCellRefs(opts.cells)
		if err != nil {
			return err
		}
		cellRefs = refs
	}

	var sheetNames []string
	if opts.sheets != "" {
		names, err := parseSheetNames(opts.sheets)
		if err != nil {
			return err
		}
		sheetNames = names
	}

	f, err := openWorkbook(filePath, false)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	if opts.allSheets {
		sheetNames = f.GetSheetList()
	}

	if len(sheetNames) > 0 {
		for _, s := range sheetNames {
			if err := verifySheet(f, s); err != nil {
				return err
			}
		}
		if !opts.confirmLarge {
			if err := guardMultiSheetSize(f, sheetNames, opts); err != nil {
				return err
			}
		}
		return readSheets(f, sheetNames, opts.cellRange, opts.jsonOut, w)
	}

	if err := verifySheet(f, sheet); err != nil {
		return err
	}

	if len(cellRefs) > 0 {
		return readCells(f, sheet, cellRefs, opts.jsonOut, w)
	}

	if opts.cell != "" {
		return readSingleCell(f, sheet, opts.cell, opts.jsonOut, w)
	}

	return readRange(f, sheet, opts.cellRange, opts.jsonOut, w)
}

// validateReadModes rejects invalid flag combinations before any file I/O.
func validateReadModes(sheet string, opts readOptions) error {
	if opts.cells != "" && (opts.cell != "" || opts.cellRange != "") {
		return fmt.Errorf("--cells cannot be combined with --cell or --range")
	}
	if opts.cell != "" && opts.cellRange != "" {
		return fmt.Errorf("--cell cannot be combined with --range")
	}
	if opts.cell != "" {
		if _, _, err := excelize.CellNameToCoordinates(opts.cell); err != nil {
			return fmt.Errorf("invalid cell reference %q: %w", opts.cell, err)
		}
	}
	if opts.sheets != "" && opts.allSheets {
		return fmt.Errorf("--sheets cannot be combined with --all-sheets")
	}
	if opts.sheets != "" || opts.allSheets {
		flag := "--sheets"
		if opts.allSheets {
			flag = "--all-sheets"
		}
		if sheet != "" {
			return fmt.Errorf("%s cannot be combined with a positional sheet argument", flag)
		}
		if opts.cell != "" || opts.cells != "" {
			return fmt.Errorf("%s cannot be combined with --cell or --cells", flag)
		}
	}
	return nil
}

// maxAutoMultiSheetCells bounds how many cells a --sheets/--all-sheets read may
// span before it requires an explicit --confirm-large. A multi-sheet read of a
// large workbook (e.g. a full LV calculation across every sheet) can emit
// millions of tokens in one call and overflow an agent's context window; this
// gate forces a deliberate confirmation or a scoped read instead.
const maxAutoMultiSheetCells = 10_000

// guardMultiSheetSize blocks a --sheets/--all-sheets read whose estimated cell
// span exceeds maxAutoMultiSheetCells, telling the caller how to confirm or
// scope it. It fails open: if the span cannot be estimated, the read proceeds
// (the read itself will surface any real error).
func guardMultiSheetSize(f *excelize.File, sheets []string, opts readOptions) error {
	est, err := estimateMultiSheetCells(f, sheets, opts.cellRange)
	if err != nil || est <= maxAutoMultiSheetCells {
		return nil
	}
	mode := "--sheets"
	if opts.allSheets {
		mode = "--all-sheets"
	}
	return fmt.Errorf(
		"%s spans ~%d cells across %d sheet(s) — this can emit a very large output "+
			"(potentially millions of tokens) and overflow the agent context window.\n"+
			"Either re-run with --confirm-large to force it, or scope the read: a single "+
			"sheet (positional arg), --range A1:D50, or --cells C21,L21.\n"+
			"Run \"xlsx-tool info <file>\" first to see per-sheet dimensions.",
		mode, est, len(sheets))
}

// estimateMultiSheetCells approximates the number of cells a multi-sheet read
// will span. With an explicit --range the same rectangle applies to every
// sheet; otherwise each sheet contributes its full data extent (displayBounds).
// The bounding-box estimate counts blank cells too, which is the safe side for
// a guard.
func estimateMultiSheetCells(f *excelize.File, sheets []string, cellRange string) (int, error) {
	if cellRange != "" {
		norm, err := normalizeRange(cellRange)
		if err != nil {
			return 0, err
		}
		cols, rows, err := rangeCellSpan(norm)
		if err != nil {
			return 0, err
		}
		return cols * rows * len(sheets), nil
	}
	total := 0
	for _, s := range sheets {
		endCol, endRow, err := displayBounds(f, s)
		if err != nil {
			return 0, err
		}
		total += endCol * endRow
	}
	return total, nil
}

// rangeCellSpan returns the column and row span of a normalized "A1:B2" range
// (a bare "A1" spans 1x1).
func rangeCellSpan(cellRange string) (cols, rows int, err error) {
	colonIdx := findColon(cellRange)
	if colonIdx == -1 {
		return 1, 1, nil
	}
	startCol, startRow, err := excelize.CellNameToCoordinates(cellRange[:colonIdx])
	if err != nil {
		return 0, 0, fmt.Errorf("invalid range start: %w", err)
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(cellRange[colonIdx+1:])
	if err != nil {
		return 0, 0, fmt.Errorf("invalid range end: %w", err)
	}
	return endCol - startCol + 1, endRow - startRow + 1, nil
}

// verifySheet returns an error listing available sheets if the sheet does not exist.
func verifySheet(f *excelize.File, sheet string) error {
	idx, err := f.GetSheetIndex(sheet)
	if err != nil || idx == -1 {
		available := f.GetSheetList()
		return fmt.Errorf("sheet not found: %q\nAvailable sheets: %v", sheet, available)
	}
	return nil
}

// parseCellRefs parses a comma-separated list of A1-style cell references.
// Whitespace around each reference is tolerated; empty segments are skipped.
// Duplicate references are collapsed to their first occurrence so plain and
// JSON output (keyed by reference) agree.
func parseCellRefs(s string) ([]string, error) {
	var refs []string
	seen := make(map[string]bool)
	for _, part := range strings.Split(s, ",") {
		ref := strings.TrimSpace(part)
		if ref == "" {
			continue
		}
		if _, _, err := excelize.CellNameToCoordinates(ref); err != nil {
			return nil, fmt.Errorf("invalid cell reference %q: %w", ref, err)
		}
		if seen[ref] {
			continue
		}
		seen[ref] = true
		refs = append(refs, ref)
	}
	if len(refs) == 0 {
		return nil, fmt.Errorf("--cells: no cell references given")
	}
	return refs, nil
}

// parseSheetNames parses a comma-separated list of sheet names.
// Names may contain spaces, so the input is split on commas only;
// whitespace around each name is trimmed and empty segments are skipped.
func parseSheetNames(s string) ([]string, error) {
	var names []string
	for _, part := range strings.Split(s, ",") {
		name := strings.TrimSpace(part)
		if name == "" {
			continue
		}
		names = append(names, name)
	}
	if len(names) == 0 {
		return nil, fmt.Errorf("--sheets: no sheet names given")
	}
	return names, nil
}

// cellData holds the rich data for a single cell.
type cellData struct {
	value    string            // raw stored value (unformatted)
	display  string            // formatted display value (e.g. " 0.23 € ", "22.50%")
	formula  string            // formula source without leading "="
	fill     string            // hex color like "#FFFF00", empty if no fill
	cellType excelize.CellType // resolved cell type, used to type the JSON value
}

func readSingleCell(f *excelize.File, sheet, cell string, jsonOut bool, w io.Writer) error {
	cd := getCellData(f, sheet, cell)

	if jsonOut {
		result := buildCellJSON(cd)
		enc := json.NewEncoder(w)
		enc.SetIndent("", "  ")
		return enc.Encode(result)
	}

	plain := format.FormatCellPlain(format.PlainCell{
		Value:   cd.value,
		Display: cd.display,
		Formula: cd.formula,
		Fill:    cd.fill,
	})
	if plain == "" {
		plain = "(empty)"
	}
	_, err := fmt.Fprintln(w, plain)
	return err
}

// readCells reads multiple scattered cells in input order from an already-open workbook.
func readCells(f *excelize.File, sheet string, refs []string, jsonOut bool, w io.Writer) error {
	if jsonOut {
		result := &format.CellsJSON{
			Sheet: sheet,
			Cells: make(map[string]any, len(refs)),
		}
		for _, ref := range refs {
			result.Cells[ref] = buildCellJSON(getCellData(f, sheet, ref))
		}
		return format.WriteCellsJSON(w, result)
	}

	for _, ref := range refs {
		cd := getCellData(f, sheet, ref)
		plain := format.FormatCellPlain(format.PlainCell{
			Value:   cd.value,
			Display: cd.display,
			Formula: cd.formula,
			Fill:    cd.fill,
		})
		if plain == "" {
			plain = "(empty)"
		}
		if _, err := fmt.Fprintf(w, "%s: %s\n", ref, plain); err != nil {
			return err
		}
	}
	return nil
}

// readSheets reads multiple sheets in order from an already-open workbook.
// The same cellRange (if any) is applied to every sheet.
func readSheets(f *excelize.File, sheets []string, cellRange string, jsonOut bool, w io.Writer) error {
	if jsonOut {
		workbook := &format.WorkbookJSON{}
		for _, sheet := range sheets {
			cells, startCol, startRow, rangeStr, err := collectSheet(f, sheet, cellRange)
			if err != nil {
				return err
			}
			workbook.Sheets = append(workbook.Sheets,
				buildSheetJSON(sheet, rangeStr, cells, startCol, startRow))
		}
		return format.WriteWorkbookJSON(w, workbook)
	}

	for i, sheet := range sheets {
		if i > 0 {
			if _, err := fmt.Fprintln(w); err != nil {
				return err
			}
		}
		if err := readRange(f, sheet, cellRange, false, w); err != nil {
			return err
		}
	}
	return nil
}

func readRange(f *excelize.File, sheet, cellRange string, jsonOut bool, w io.Writer) error {
	cells, startCol, startRow, rangeStr, err := collectSheet(f, sheet, cellRange)
	if err != nil {
		return err
	}

	if jsonOut {
		return format.WriteJSON(w, buildSheetJSON(sheet, rangeStr, cells, startCol, startRow))
	}

	return readRangePlain(w, sheet, rangeStr, cells, startCol, startRow)
}

// collectSheet reads a sheet's cells and resolves the effective range string.
func collectSheet(
	f *excelize.File, sheet, cellRange string,
) ([][]cellData, int, int, string, error) {
	if cellRange == "" {
		resolved, err := fullSheetRange(f, sheet)
		if err != nil {
			return nil, 0, 0, "", err
		}
		cellRange = resolved
	} else {
		normalized, err := normalizeRange(cellRange)
		if err != nil {
			return nil, 0, 0, "", err
		}
		cellRange = normalized
	}

	cells, startCol, startRow, err := getCellsInRange(f, sheet, cellRange)
	if err != nil {
		return nil, 0, 0, "", err
	}

	return cells, startCol, startRow, cellRange, nil
}

// fullSheetRange returns the effective full-sheet range ("A1:<end>") derived
// from actual data bounds (plus a plausible stored dimension — see
// displayBounds). Returns "" for a sheet without any cells.
func fullSheetRange(f *excelize.File, sheet string) (string, error) {
	endCol, endRow, err := displayBounds(f, sheet)
	if err != nil {
		return "", err
	}
	if endCol == 0 || endRow == 0 {
		return "", nil
	}
	endRef, err := excelize.CoordinatesToCellName(endCol, endRow)
	if err != nil {
		return "", err
	}
	if endRef == "A1" {
		return "A1", nil
	}
	return "A1:" + endRef, nil
}

// buildSheetJSON converts a sheet's cell grid to its SheetJSON representation.
func buildSheetJSON(
	sheet, rangeStr string, cells [][]cellData, startCol, startRow int,
) *format.SheetJSON {
	result := &format.SheetJSON{
		Sheet: sheet,
		Range: rangeStr,
	}

	for ri, row := range cells {
		rowJSON := format.RowJSON{
			Row:   startRow + ri,
			Cells: make(map[string]any),
		}
		for ci, cd := range row {
			colName, _ := excelize.ColumnNumberToName(startCol + ci)
			rowJSON.Cells[colName] = buildCellJSON(cd)
		}
		result.Rows = append(result.Rows, rowJSON)
	}

	return result
}

func readRangePlain(w io.Writer, sheet, rangeStr string, cells [][]cellData, startCol, startRow int) error {
	numCols := 0
	if len(cells) > 0 {
		numCols = len(cells[0])
	}

	colNames := make([]string, numCols)
	for i := range numCols {
		colNames[i], _ = excelize.ColumnNumberToName(startCol + i)
	}

	plainRows := make([][]format.PlainCell, len(cells))
	for ri, row := range cells {
		plainRow := make([]format.PlainCell, len(row))
		for ci, cd := range row {
			plainRow[ci] = format.PlainCell{
				Value:   cd.value,
				Display: cd.display,
				Formula: cd.formula,
				Fill:    cd.fill,
			}
		}
		plainRows[ri] = plainRow
	}

	return format.WritePlainSheet(w, sheet, rangeStr, plainRows, startCol, startRow, colNames)
}

// buildCellJSON converts cellData to the JSON representation.
// Returns nil for truly empty cells (no value, no formula, no fill).
// "value" carries the raw typed value: a JSON number for numeric/date cells,
// else the raw string. "display" is added only when the formatted display value
// differs from the raw value. Formula cells whose cached value is a known Excel
// error get an "error" field (and no typed/display value).
func buildCellJSON(cd cellData) any {
	if cd.formula == "" && cd.value == "" && cd.fill == "" {
		return nil
	}

	cell := map[string]any{}
	if cd.formula != "" {
		cell["formula"] = cd.formula
		if cd.value != "" {
			if format.IsErrorValue(cd.value) {
				cell["value"] = nil
				cell["error"] = cd.value
			} else {
				cell["value"] = typedCellValue(cd)
				addDisplay(cell, cd)
			}
		}
	} else if cd.value != "" {
		cell["value"] = typedCellValue(cd)
		addDisplay(cell, cd)
	} else {
		cell["value"] = nil
	}

	if cd.fill != "" {
		cell["fill"] = cd.fill
	}

	return cell
}

// maxExactFloatInt is 2^53, the largest integer a float64 represents exactly.
// Integer literals with a larger magnitude lose precision once parsed as float64.
const maxExactFloatInt = 1 << 53 // 9007199254740992

// typedCellValue returns the raw value typed for JSON: a number when the cell
// is numeric and the raw value parses as one, else the raw string.
//
// Numbers in the xlsx format carry no explicit type attribute, so excelize
// reports them (and numeric formula results) as CellTypeUnset rather than
// CellTypeNumber. We therefore type as a number for every type EXCEPT the
// explicit non-numeric ones (shared/inline strings, bool, error). This keeps
// text cells like "00123" (a shared string) as strings while correctly typing
// real numbers and cached numeric formula results.
//
// Decimals and small integers are emitted as float64 so the common cases stay
// clean (0.015, 0.225, 42, 30). Large integer literals beyond float64's exact
// range (2^53) — e.g. 14-digit article numbers — would silently lose precision
// as a float, so those are emitted verbatim via json.Number (marshals unquoted,
// lossless).
func typedCellValue(cd cellData) any {
	switch cd.cellType {
	case excelize.CellTypeSharedString, excelize.CellTypeInlineString,
		excelize.CellTypeBool, excelize.CellTypeError:
		return cd.value
	}
	if isExactIntBeyondFloat53(cd.value) {
		return json.Number(cd.value)
	}
	if n, err := strconv.ParseFloat(cd.value, 64); err == nil {
		return n
	}
	return cd.value
}

// isExactIntBeyondFloat53 reports whether s is an integer literal (optional
// leading "-", digits only — no ".", "e", or "E") whose absolute value exceeds
// 2^53, i.e. would lose precision when parsed as a float64.
func isExactIntBeyondFloat53(s string) bool {
	digits := strings.TrimPrefix(s, "-")
	if digits == "" {
		return false
	}
	for _, r := range digits {
		if r < '0' || r > '9' {
			return false
		}
	}
	// Pure integer literal: ParseUint on the digit string is exact (no float
	// rounding). Overflow of uint64 means it is far beyond 2^53 and definitely
	// needs verbatim preservation.
	n, err := strconv.ParseUint(digits, 10, 64)
	if err != nil {
		return true
	}
	return n > maxExactFloatInt
}

// addDisplay attaches the formatted display value to the JSON cell when it
// differs from the raw value (no redundant display for plain numbers/text).
func addDisplay(cell map[string]any, cd cellData) {
	if cd.display != "" && cd.display != cd.value {
		cell["display"] = cd.display
	}
}

// getCellData extracts the raw value, formatted display value, formula, cell
// type, and fill color for a single cell. The raw value (unformatted, e.g.
// "0.225") and the formatted display value (e.g. "22.50%") are read separately
// so callers can surface both. For formula cells the cached computed value is
// read the same way (raw + display).
func getCellData(f *excelize.File, sheet, cell string) cellData {
	var cd cellData

	if formula, _ := f.GetCellFormula(sheet, cell); formula != "" {
		cd.formula = formula
	}

	// Raw stored value (unformatted) and formatted display value.
	raw, _ := f.GetCellValue(sheet, cell, excelize.Options{RawCellValue: true})
	disp, _ := f.GetCellValue(sheet, cell)
	cd.value = raw
	cd.display = disp

	cd.cellType, _ = f.GetCellType(sheet, cell)
	cd.fill = getCellFill(f, sheet, cell)
	return cd
}

// getCellFill extracts the background fill color of a cell.
// Returns hex color like "#FFFF00" or empty string if no fill.
func getCellFill(f *excelize.File, sheet, cell string) string {
	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil || styleID == 0 {
		return ""
	}
	style, err := f.GetStyle(styleID)
	if err != nil || style == nil {
		return ""
	}
	if style.Fill.Type != "pattern" || style.Fill.Pattern == 0 || len(style.Fill.Color) == 0 {
		return ""
	}
	color := style.Fill.Color[0]
	if color == "" {
		return ""
	}
	// Normalize to #RRGGBB
	color = strings.TrimPrefix(color, "#")
	switch len(color) {
	case 8: // AARRGGBB → RRGGBB
		return "#" + color[2:]
	case 6:
		return "#" + color
	default:
		return "#" + color
	}
}

func getCellsInRange(
	f *excelize.File, sheet, cellRange string,
) ([][]cellData, int, int, error) {
	if cellRange == "" {
		// Empty sheet (resolved upstream via fullSheetRange).
		return nil, 0, 0, nil
	}

	colonIdx := findColon(cellRange)
	if colonIdx == -1 {
		// Single cell dimension (e.g. "A1") → treat as "A1:A1"
		cellRange = cellRange + ":" + cellRange
		colonIdx = findColon(cellRange)
	}

	startCol, startRow, err := excelize.CellNameToCoordinates(cellRange[:colonIdx])
	if err != nil {
		return nil, 0, 0, fmt.Errorf("invalid range start: %w", err)
	}
	endCol, endRow, err := excelize.CellNameToCoordinates(cellRange[colonIdx+1:])
	if err != nil {
		return nil, 0, 0, fmt.Errorf("invalid range end: %w", err)
	}

	var result [][]cellData
	for r := startRow; r <= endRow; r++ {
		var dataRow []cellData
		for c := startCol; c <= endCol; c++ {
			colName, _ := excelize.ColumnNumberToName(c)
			cell := fmt.Sprintf("%s%d", colName, r)
			cd := getCellData(f, sheet, cell)
			dataRow = append(dataRow, cd)
		}
		result = append(result, dataRow)
	}

	return result, startCol, startRow, nil
}

// normalizeRange orders a rectangular range so the start is the top-left and
// the end the bottom-right corner (e.g. "B2:A1" -> "A1:B2", "A4:B1" ->
// "A1:B4") — a reversed range would otherwise silently yield zero cells.
// Single cell references pass through unchanged.
func normalizeRange(cellRange string) (string, error) {
	colonIdx := findColon(cellRange)
	if colonIdx == -1 {
		return cellRange, nil
	}
	c1, r1, err := excelize.CellNameToCoordinates(cellRange[:colonIdx])
	if err != nil {
		return "", fmt.Errorf("invalid range start: %w", err)
	}
	c2, r2, err := excelize.CellNameToCoordinates(cellRange[colonIdx+1:])
	if err != nil {
		return "", fmt.Errorf("invalid range end: %w", err)
	}
	startRef, err := excelize.CoordinatesToCellName(min(c1, c2), min(r1, r2))
	if err != nil {
		return "", err
	}
	endRef, err := excelize.CoordinatesToCellName(max(c1, c2), max(r1, r2))
	if err != nil {
		return "", err
	}
	return startRef + ":" + endRef, nil
}

func findColon(s string) int {
	for i, c := range s {
		if c == ':' {
			return i
		}
	}
	return -1
}
