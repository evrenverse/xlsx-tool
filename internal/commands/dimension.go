package commands

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// maxTrustedDimensionCells caps how large a stored dimension may be before it
// is considered a lie and ignored for display purposes. Stored dimensions are
// just an attribute — files in the wild carry values like "A1:XFD1048576"
// (17 billion cells), which must never drive per-cell iteration.
const maxTrustedDimensionCells = 1_000_000

// dataBounds returns the extent (1-based, inclusive) of cells actually
// present in sheetData, derived via GetRows — which parses real rows instead
// of trusting the stored dimension. Formula cells are included even without
// a cached value; only trailing blank cells/rows are trimmed.
func dataBounds(f *excelize.File, sheet string) (endCol, endRow int, err error) {
	rows, err := f.GetRows(sheet)
	if err != nil {
		return 0, 0, fmt.Errorf("read rows of %q: %w", sheet, err)
	}
	endRow = len(rows)
	for _, row := range rows {
		endCol = max(endCol, len(row))
	}
	return endCol, endRow, nil
}

// dimensionEnd parses the stored dimension's end coordinates.
// ok is false when the dimension is absent or unparsable.
func dimensionEnd(f *excelize.File, sheet string) (endCol, endRow int, ok bool) {
	dim, err := f.GetSheetDimension(sheet)
	if err != nil || dim == "" {
		return 0, 0, false
	}
	end := dim
	if colonIdx := findColon(dim); colonIdx != -1 {
		end = dim[colonIdx+1:]
	}
	c, r, err := excelize.CellNameToCoordinates(end)
	if err != nil {
		return 0, 0, false
	}
	return c, r, true
}

// displayBounds returns the cell area full-sheet reads should show: the
// actual data bounds, expanded by the stored dimension when it is plausible.
// Expansion keeps trailing style-only cells (e.g. fill targets, which
// GetRows trims) visible on honest files; the plausibility cap prevents a
// lying oversized dimension from driving per-cell iteration, and an
// undersized dimension can never hide data because the data bounds win.
func displayBounds(f *excelize.File, sheet string) (endCol, endRow int, err error) {
	endCol, endRow, err = dataBounds(f, sheet)
	if err != nil {
		return 0, 0, err
	}
	if dimCol, dimRow, ok := dimensionEnd(f, sheet); ok {
		if int64(dimCol)*int64(dimRow) <= maxTrustedDimensionCells {
			endCol, endRow = max(endCol, dimCol), max(endRow, dimRow)
		}
	}
	return endCol, endRow, nil
}

// ensureDimensionContains extends a sheet's stored dimension so it covers cell.
//
// Excelize does not update the dimension attribute on writes, and real Excel
// files always keep it in sync with content. displayBounds trusts a
// plausible stored dimension to include trailing style-only cells that
// dataBounds (which only sees real sheetData) would otherwise trim — so a
// write landing outside the old dimension must still push it out, mirroring
// what Excel itself does on save.
func ensureDimensionContains(f *excelize.File, sheet, cell string) error {
	col, row, err := excelize.CellNameToCoordinates(cell)
	if err != nil {
		return fmt.Errorf("invalid cell reference %q: %w", cell, err)
	}

	minCol, minRow, maxCol, maxRow := col, row, col, row

	dim, err := f.GetSheetDimension(sheet)
	if err != nil {
		return fmt.Errorf("get dimension of %q: %w", sheet, err)
	}
	if dim != "" {
		start, end := dim, dim
		if colonIdx := findColon(dim); colonIdx != -1 {
			start, end = dim[:colonIdx], dim[colonIdx+1:]
		}
		startCol, startRow, err := excelize.CellNameToCoordinates(start)
		if err != nil {
			return fmt.Errorf("invalid dimension start %q: %w", dim, err)
		}
		endCol, endRow, err := excelize.CellNameToCoordinates(end)
		if err != nil {
			return fmt.Errorf("invalid dimension end %q: %w", dim, err)
		}
		minCol, minRow = min(minCol, startCol), min(minRow, startRow)
		maxCol, maxRow = max(maxCol, endCol), max(maxRow, endRow)
		if minCol == startCol && minRow == startRow && maxCol == endCol && maxRow == endRow {
			return nil // already covered
		}
	}

	startRef, err := excelize.CoordinatesToCellName(minCol, minRow)
	if err != nil {
		return err
	}
	endRef, err := excelize.CoordinatesToCellName(maxCol, maxRow)
	if err != nil {
		return err
	}

	newDim := startRef
	if endRef != startRef {
		newDim = startRef + ":" + endRef
	}
	return f.SetSheetDimension(sheet, newDim)
}
