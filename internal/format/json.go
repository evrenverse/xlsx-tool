package format

import (
	"encoding/json"
	"io"
)

// SheetJSON represents JSON output for a sheet range.
type SheetJSON struct {
	Sheet string    `json:"sheet"`
	Range string    `json:"range,omitempty"`
	Rows  []RowJSON `json:"rows"`
}

// RowJSON represents a single row in JSON output.
type RowJSON struct {
	Row   int            `json:"row"`
	Cells map[string]any `json:"cells"`
}

// CellsJSON represents JSON output for scattered cell reads (--cells).
// Cells maps each requested cell reference to its cell object (or null if empty).
type CellsJSON struct {
	Sheet string         `json:"sheet"`
	Cells map[string]any `json:"cells"`
}

// WorkbookJSON represents JSON output for multi-sheet reads (--sheets / --all-sheets).
// Sheets are listed in requested (or workbook) order.
type WorkbookJSON struct {
	Sheets []*SheetJSON `json:"sheets"`
}

// WriteJSON writes the SheetJSON structure as formatted JSON.
func WriteJSON(w io.Writer, data *SheetJSON) error {
	return WriteIndentedJSON(w, data)
}

// WriteCellsJSON writes the CellsJSON structure as formatted JSON.
func WriteCellsJSON(w io.Writer, data *CellsJSON) error {
	return WriteIndentedJSON(w, data)
}

// WriteWorkbookJSON writes the WorkbookJSON structure as formatted JSON.
func WriteWorkbookJSON(w io.Writer, data *WorkbookJSON) error {
	return WriteIndentedJSON(w, data)
}

// WriteIndentedJSON encodes v with the shared two-space indentation.
func WriteIndentedJSON(w io.Writer, v any) error {
	enc := json.NewEncoder(w)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}
