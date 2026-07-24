package commands

import (
	"errors"
	"fmt"
	"io"
	"os"
	"strings"

	"github.com/evrenverse/xlsx-tool/internal/format"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// ErrNoMatches signals a zero-match find. The result (message or JSON) has
// already been written to stdout; the process exits non-zero without any
// additional error output (cross-tool convention with docx/pdf tools).
var ErrNoMatches = errors.New("no matches found")

type findOptions struct {
	exact    bool
	formulas bool
	max      int
	jsonOut  bool
}

// findMatch is the location of one matching cell.
type findMatch struct {
	sheet string
	cell  string
}

// findMatchJSON is one match in --json output. Kind is always "cell"
// (cross-tool parity with the docx/pdf find commands). Formula carries the
// formula without the leading "=" (as stored by Excel), matching read and
// the write/batch JSON results. Value is the raw stored value and Display the
// formatted value (e.g. "22.50%"), emitted only when it differs from Value —
// the same "raw value + display" shape as the read command, so a user who
// searched by the formatted text still sees it.
type findMatchJSON struct {
	Kind    string `json:"kind"`
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Value   string `json:"value"`
	Display string `json:"display,omitempty"`
	Formula string `json:"formula,omitempty"`
}

// findResultJSON is the --json output of the find command.
type findResultJSON struct {
	Matches   []findMatchJSON `json:"matches"`
	Total     int             `json:"total"`
	Truncated bool            `json:"truncated"`
}

func NewFindCmd() *cobra.Command {
	var opts findOptions

	cmd := &cobra.Command{
		Use:   "find <file> <text> [sheet]",
		Short: "Find cells by text content",
		Long: `Search cell values for a text and print each match's location.

By default the search is a case-insensitive substring match over ALL
sheets. Pass a sheet name as third argument to search one sheet only.

This is the agent-friendly way to locate label cells ("Total",
"Owner:") in one call instead of dumping whole sheets. Typical
workflow: find the label, then read its neighbor cells with
'read --cells' or fill them with 'write'/'batch'.

Matching:
  default      case-insensitive substring match on cell values
  --exact      case-sensitive full-cell match
  --formulas   additionally search formula strings (e.g. find "SUM(")

At most --max matches are listed (default 50). When more exist, the
output states "(showing first N of M matches — raise --max)".

Zero matches exit with code 1 and the message
'No matches found for "<text>".' (JSON mode emits the empty result).
JSON matches carry "kind": "cell"; formulas are reported without the
leading "=" (as stored by Excel), matching the read command.`,
		Example: `  # Find a label anywhere in the workbook
  xlsx-tool find report.xlsx "Total"

  # Find in one sheet only, structured output
  xlsx-tool find report.xlsx "Owner:" "Summary" --json

  # Find cells whose formula references SUM
  xlsx-tool find report.xlsx "SUM(" --formulas

  # Locate a label, then read its neighbors in one call
  xlsx-tool find report.xlsx "Grand total"
  xlsx-tool read report.xlsx "Summary" --cells D21,E21`,
		Args: cobra.RangeArgs(2, 3),
		RunE: func(cmd *cobra.Command, args []string) error {
			sheet := ""
			if len(args) == 3 {
				sheet = args[2]
			}
			return runFind(args[0], args[1], sheet, opts, os.Stdout)
		},
	}

	cmd.Flags().BoolVar(&opts.exact, "exact", false,
		"Case-sensitive full-cell match instead of case-insensitive substring")
	cmd.Flags().BoolVar(&opts.formulas, "formulas", false,
		"Also search formula strings, not just cell values")
	cmd.Flags().IntVar(&opts.max, "max", 50, "Maximum number of matches to show")
	cmd.Flags().BoolVar(&opts.jsonOut, "json", false, "Output as JSON instead of plain text")

	return cmd
}

func runFind(filePath, text, sheet string, opts findOptions, w io.Writer) error {
	if text == "" {
		return fmt.Errorf("search text must not be empty")
	}
	if opts.max < 1 {
		return fmt.Errorf("--max must be at least 1, got %d", opts.max)
	}

	f, err := openWorkbook(filePath, false)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	sheets := f.GetSheetList()
	if sheet != "" {
		if err := verifySheet(f, sheet); err != nil {
			return err
		}
		sheets = []string{sheet}
	}

	matches, err := findMatches(f, sheets, text, opts)
	if err != nil {
		return err
	}

	shown := matches
	truncated := len(matches) > opts.max
	if truncated {
		shown = matches[:opts.max]
	}

	if opts.jsonOut {
		if err := writeFindJSON(w, f, shown, len(matches), truncated); err != nil {
			return err
		}
	} else {
		if err := writeFindPlain(w, f, text, shown, len(matches), truncated); err != nil {
			return err
		}
	}

	if len(matches) == 0 {
		return ErrNoMatches
	}
	return nil
}

// findMatches scans the given sheets row-wise (GetRows, fast bulk access)
// and returns all cells whose value — or, with --formulas, formula — matches.
func findMatches(
	f *excelize.File, sheets []string, text string, opts findOptions,
) ([]findMatch, error) {
	matches := []findMatch{}

	for _, sheet := range sheets {
		rows, err := f.GetRows(sheet)
		if err != nil {
			return nil, fmt.Errorf("read sheet %q: %w", sheet, err)
		}

		for ri, row := range rows {
			for ci, val := range row {
				cellRef, err := excelize.CoordinatesToCellName(ci+1, ri+1)
				if err != nil {
					return nil, fmt.Errorf("cell coordinates (%d,%d): %w", ci+1, ri+1, err)
				}

				matched := matchText(val, text, opts.exact)
				if !matched && opts.formulas {
					formula, _ := f.GetCellFormula(sheet, cellRef)
					matched = formula != "" && matchText(formula, text, opts.exact)
				}
				if matched {
					matches = append(matches, findMatch{sheet: sheet, cell: cellRef})
				}
			}
		}
	}

	return matches, nil
}

// matchText reports whether s matches text under the chosen mode.
func matchText(s, text string, exact bool) bool {
	if s == "" {
		return false
	}
	if exact {
		return s == text
	}
	return strings.Contains(strings.ToLower(s), strings.ToLower(text))
}

func writeFindPlain(
	w io.Writer, f *excelize.File, text string, shown []findMatch, total int, truncated bool,
) error {
	if total == 0 {
		_, err := fmt.Fprintf(w, "No matches found for %q.\n", text)
		return err
	}

	for _, m := range shown {
		cd := getCellData(f, m.sheet, m.cell)
		plain := format.FormatCellPlain(format.PlainCell{
			Value:   cd.value,
			Display: cd.display,
			Formula: cd.formula,
			Fill:    cd.fill,
		})
		if _, err := fmt.Fprintf(w, "%s!%s: %s\n", m.sheet, m.cell, plain); err != nil {
			return err
		}
	}

	if truncated {
		_, err := fmt.Fprintf(w, "(showing first %d of %d matches — raise --max)\n",
			len(shown), total)
		return err
	}
	return nil
}

func writeFindJSON(
	w io.Writer, f *excelize.File, shown []findMatch, total int, truncated bool,
) error {
	result := &findResultJSON{
		Matches:   []findMatchJSON{},
		Total:     total,
		Truncated: truncated,
	}

	for _, m := range shown {
		cd := getCellData(f, m.sheet, m.cell)
		match := findMatchJSON{
			Kind:    "cell",
			Sheet:   m.sheet,
			Cell:    m.cell,
			Value:   cd.value,
			Formula: cd.formula,
		}
		// Surface the formatted value (e.g. "22.50%") only when it differs from
		// the raw one, mirroring the read command's display field.
		if cd.display != "" && cd.display != cd.value {
			match.Display = cd.display
		}
		result.Matches = append(result.Matches, match)
	}

	return format.WriteIndentedJSON(w, result)
}
