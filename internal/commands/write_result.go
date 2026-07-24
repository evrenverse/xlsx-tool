package commands

import (
	"io"

	"github.com/evrenverse/xlsx-tool/internal/format"
	"github.com/xuri/excelize/v2"
)

// writeResultCell describes one written or skipped cell in --json write output.
type writeResultCell struct {
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Value   any    `json:"value,omitempty"`
	Formula string `json:"formula,omitempty"`
	Hint    string `json:"hint,omitempty"`
}

// skippedLockedCell builds a writeResultCell for a locked/skipped cell,
// enriching it with the cell's current formula (if any) and an instructive hint
// so AI agents understand whether the cell auto-computes and what to do instead.
func skippedLockedCell(f *excelize.File, sheet, cell string) writeResultCell {
	entry := writeResultCell{Sheet: sheet, Cell: cell}
	if formula, err := f.GetCellFormula(sheet, cell); err == nil && formula != "" {
		entry.Formula = formula
		entry.Hint = "locked formula cell — computes automatically; write its input cells instead"
	} else {
		entry.Hint = "locked cell — protected by the sheet; do not treat as a blocker"
	}
	return entry
}

// NumericWarning flags a cell whose value looks numeric but was stored as text.
type NumericWarning struct {
	Sheet   string `json:"sheet"`
	Cell    string `json:"cell"`
	Value   string `json:"value"`
	Message string `json:"message"`
}

// writeResultJSON is the structured (--json) result of the write/batch commands.
// All slices are initialized so the JSON always contains arrays, never null —
// agents can index the keys without nil checks.
type writeResultJSON struct {
	Written       []writeResultCell `json:"written"`
	SkippedLocked []writeResultCell `json:"skipped_locked"`
	Warnings      []NumericWarning  `json:"warnings"`
}

// newWriteResult returns an empty result with all arrays present.
func newWriteResult() *writeResultJSON {
	return &writeResultJSON{
		Written:       []writeResultCell{},
		SkippedLocked: []writeResultCell{},
		Warnings:      []NumericWarning{},
	}
}

// emit writes the result as formatted JSON.
func (r *writeResultJSON) emit(w io.Writer) error {
	if r.Warnings == nil {
		r.Warnings = []NumericWarning{}
	}
	return format.WriteIndentedJSON(w, r)
}
