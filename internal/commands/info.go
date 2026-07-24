package commands

import (
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"

	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// InfoJSON represents the JSON output for the info command.
type InfoJSON struct {
	File        string           `json:"file"`
	Sheets      []SheetInfo      `json:"sheets"`
	NamedRanges []NamedRangeInfo `json:"named_ranges"`
}

// SheetInfo represents a single sheet in the info output.
type SheetInfo struct {
	Name      string `json:"name"`
	Dimension string `json:"dimension"`
	Formulas  int    `json:"formulas"`
	Protected bool   `json:"protected"`
}

// NamedRangeInfo represents a defined name in the workbook.
type NamedRangeInfo struct {
	Name     string `json:"name"`
	RefersTo string `json:"refers_to"`
}

func NewInfoCmd() *cobra.Command {
	var jsonOut bool

	cmd := &cobra.Command{
		Use:   "info <file>",
		Short: "Show Excel file structure",
		Long: `Display an overview of an Excel file:
filename, sheets with their names and used cell ranges (dimensions).

Output is plain text by default. Use --json for structured JSON output.

Use this as a first step to understand an unfamiliar spreadsheet
before reading or writing specific cells.`,
		Example: `  # Show file overview
  xlsx-tool info report.xlsx

  # Output:
  #   File: report.xlsx
  #   Sheets: 2
  #     - Summary (A1:F20) | Formulas: 12
  #     - Data (A1:Z500) | Formulas: 0 | Protected

  # JSON output
  xlsx-tool info report.xlsx --json`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			return runInfo(args[0], jsonOut, os.Stdout)
		},
	}

	cmd.Flags().BoolVar(&jsonOut, "json", false, "Output as JSON instead of plain text")

	return cmd
}

// countFormulas counts the number of formula cells in a sheet. It iterates
// actual data bounds, never the stored dimension — a stored dimension can
// lie in either direction; formula cells always exist in sheetData regardless.
func countFormulas(f *excelize.File, sheet string) int {
	endCol, endRow, err := dataBounds(f, sheet)
	if err != nil || endCol == 0 || endRow == 0 {
		return 0
	}

	count := 0
	for r := 1; r <= endRow; r++ {
		for c := 1; c <= endCol; c++ {
			colName, _ := excelize.ColumnNumberToName(c)
			cellRef := fmt.Sprintf("%s%d", colName, r)
			formula, _ := f.GetCellFormula(sheet, cellRef)
			if formula != "" {
				count++
			}
		}
	}
	return count
}

// isSheetProtected returns true if the sheet has protection enabled.
func isSheetProtected(f *excelize.File, sheet string) bool {
	opts, err := f.GetSheetProtection(sheet)
	if err != nil {
		return false
	}
	// An unprotected sheet returns all zero values. A protected sheet usually
	// has at least some boolean fields set to true (negated from XML defaults).
	if opts.AlgorithmName != "" ||
		opts.AutoFilter || opts.DeleteColumns || opts.DeleteRows ||
		opts.EditObjects || opts.EditScenarios ||
		opts.FormatCells || opts.FormatColumns || opts.FormatRows ||
		opts.InsertColumns || opts.InsertHyperlinks || opts.InsertRows ||
		opts.PivotTables || opts.SelectLockedCells || opts.SelectUnlockedCells ||
		opts.Sort {
		return true
	}
	// Legacy-password protection with every option attribute explicitly
	// blocked is invisible above (GetSheetProtection exposes neither the
	// password hash nor the sheet flag) — fall back to the raw XML.
	return hasRawSheetProtection(f, sheet)
}

func runInfo(filePath string, jsonOut bool, w io.Writer) error {
	f, err := openWorkbook(filePath, false)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	result := InfoJSON{
		File: filepath.Base(filePath),
	}

	for _, name := range f.GetSheetList() {
		dim, err := f.GetSheetDimension(name)
		if err != nil {
			dim = "unknown"
		}
		formulas := countFormulas(f, name)
		protected := isSheetProtected(f, name)
		result.Sheets = append(result.Sheets, SheetInfo{
			Name:      name,
			Dimension: dim,
			Formulas:  formulas,
			Protected: protected,
		})
	}

	for _, dn := range f.GetDefinedName() {
		result.NamedRanges = append(result.NamedRanges, NamedRangeInfo{
			Name:     dn.Name,
			RefersTo: dn.RefersTo,
		})
	}

	if jsonOut {
		enc := json.NewEncoder(w)
		enc.SetIndent("", "  ")
		return enc.Encode(result)
	}

	fmt.Fprintf(w, "File: %s\n", result.File)
	fmt.Fprintf(w, "Sheets: %d\n", len(result.Sheets))
	for _, s := range result.Sheets {
		line := fmt.Sprintf("  - %s (%s) | Formulas: %d", s.Name, s.Dimension, s.Formulas)
		if s.Protected {
			line += " | Protected"
		}
		fmt.Fprintln(w, line)
	}

	if len(result.NamedRanges) > 0 {
		fmt.Fprintf(w, "\nNamed Ranges: %d\n", len(result.NamedRanges))
		for _, nr := range result.NamedRanges {
			fmt.Fprintf(w, "  - %s -> %s\n", nr.Name, nr.RefersTo)
		}
	}

	return nil
}
