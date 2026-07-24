package commands

import "github.com/xuri/excelize/v2"

const (
	maxWorkbookExpandedBytes = 512 << 20
	maxWorkbookXMLBytes      = 64 << 20
)

// openWorkbook applies explicit ZIP expansion limits to attacker-controlled
// XLSX packages. XML parts larger than maxWorkbookXMLBytes spill to temporary
// files; the complete expanded package may never exceed maxWorkbookExpandedBytes.
func openWorkbook(path string, rawCellValue bool) (*excelize.File, error) {
	return excelize.OpenFile(path, excelize.Options{
		RawCellValue:      rawCellValue,
		UnzipSizeLimit:    maxWorkbookExpandedBytes,
		UnzipXMLSizeLimit: maxWorkbookXMLBytes,
	})
}
