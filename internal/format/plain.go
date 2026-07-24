package format

import (
	"fmt"
	"io"
	"strconv"
	"strings"
)

// PlainCell holds the data for a single cell in plain-text output.
type PlainCell struct {
	Value   string
	Display string // formatted display value, appended in brackets when it differs from Value
	Formula string
	Fill    string // hex color like "#FFFF00", empty if no fill
}

// FormatCellPlain formats a single cell for plain-text output.
// Returns empty string for truly empty cells (no value, no formula, no fill).
// Formula cells whose cached value is a known Excel error get an [ERROR] marker.
// When a cell carries a Display value that differs from its raw Value, the
// verbatim formatted string is appended in brackets, e.g. 0.225 [display "22.50%"].
func FormatCellPlain(c PlainCell) string {
	switch {
	case c.Formula != "":
		s := "=" + c.Formula
		if c.Value != "" {
			s += " → " + formatValue(c.Value)
			if IsErrorValue(c.Value) {
				s += " [ERROR]"
			} else {
				s += displaySuffix(c)
			}
		}
		if c.Fill != "" {
			s += " [fill=" + c.Fill + "]"
		}
		return s
	case c.Value != "":
		s := formatValue(c.Value) + displaySuffix(c)
		if c.Fill != "" {
			s += " [fill=" + c.Fill + "]"
		}
		return s
	case c.Fill != "":
		return "[fill=" + c.Fill + "]"
	default:
		return ""
	}
}

// displaySuffix returns ` [display "<display>"]` when the formatted display
// value differs from the raw value, else an empty string. The display string is
// emitted verbatim, including any format-defined leading/trailing spaces.
func displaySuffix(c PlainCell) string {
	if c.Display == "" || c.Display == c.Value {
		return ""
	}
	return ` [display "` + c.Display + `"]`
}

// WritePlainSheet writes a sheet range in compact plain-text format.
// Empty rows (all cells empty) are skipped entirely.
func WritePlainSheet(w io.Writer, sheet, rangeStr string, rows [][]PlainCell, startCol, startRow int, colNames []string) error {
	fmt.Fprintf(w, "Sheet: %s | Range: %s\n", sheet, rangeStr)

	for ri, row := range rows {
		rowNum := startRow + ri
		var parts []string
		for ci, c := range row {
			s := FormatCellPlain(c)
			if s == "" {
				continue
			}
			parts = append(parts, colNames[ci]+": "+s)
		}
		if len(parts) == 0 {
			continue
		}
		fmt.Fprintf(w, "[%d] %s\n", rowNum, strings.Join(parts, " | "))
	}

	return nil
}

// formatValue formats a cell value: numbers without quotes, strings with quotes.
func formatValue(v string) string {
	if _, err := strconv.ParseFloat(v, 64); err == nil {
		return v
	}
	return `"` + v + `"`
}

// errorValues contains all standard Excel formula error strings.
var errorValues = map[string]bool{
	"#VALUE!": true,
	"#DIV/0!": true,
	"#REF!":   true,
	"#NAME?":  true,
	"#NULL!":  true,
	"#NUM!":   true,
	"#N/A":    true,
}

// IsErrorValue returns true if the string matches a known Excel error value.
func IsErrorValue(s string) bool {
	return errorValues[s]
}
