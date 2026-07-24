package commands

import (
	"encoding/xml"
	"strings"

	"github.com/xuri/excelize/v2"
)

// isCellLocked reports whether a cell carries the "locked" protection flag.
//
// Excel's default is locked=true: a cell is only editable under sheet
// protection when its style explicitly sets Locked=false. We mirror that —
// any cell whose style does not explicitly unlock it counts as locked, and a
// lookup error is treated as locked (fail safe, never overwrite).
func isCellLocked(f *excelize.File, sheet, cell string) bool {
	styleID, err := f.GetCellStyle(sheet, cell)
	if err != nil {
		return true
	}
	style, err := f.GetStyle(styleID)
	if err != nil || style == nil || style.Protection == nil {
		return true
	}
	return style.Protection.Locked
}

// isWriteBlocked reports whether writing a cell must be refused because the
// sheet is protected and the target cell is locked. Writing such a cell
// corrupts formula/total cells (e.g. overwriting "=SUM(...)" with a literal),
// which cascades into hundreds of #DIV/0!/#REF! errors — so we never do it.
// On an unprotected sheet, locked cells are freely editable, so we don't block.
func isWriteBlocked(f *excelize.File, sheet, cell string) bool {
	return isSheetProtected(f, sheet) && isCellLocked(f, sheet, cell)
}

// hasRawSheetProtection reports whether the sheet's raw XML carries an
// enforced <sheetProtection> element.
//
// Needed because GetSheetProtection exposes only AlgorithmName plus negated
// option booleans: a sheet protected via the legacy password attribute with
// every option attribute explicitly blocked ("1") yields all-false options
// and an empty AlgorithmName — indistinguishable from an unprotected sheet
// through that API. The raw worksheet XML (available via the exported
// File.Pkg part map) is the only reliable signal for this variant.
func hasRawSheetProtection(f *excelize.File, sheet string) bool {
	data := rawSheetXML(f, sheet)
	if data == nil {
		return false
	}

	var ws struct {
		SheetProtection *struct {
			Sheet         bool   `xml:"sheet,attr"`
			Password      string `xml:"password,attr"`
			AlgorithmName string `xml:"algorithmName,attr"`
		} `xml:"sheetProtection"`
	}
	if err := xml.Unmarshal(data, &ws); err != nil || ws.SheetProtection == nil {
		return false
	}
	sp := ws.SheetProtection
	return sp.Sheet || sp.Password != "" || sp.AlgorithmName != ""
}

// rawSheetXML returns the raw worksheet XML for a sheet name by resolving
// workbook.xml -> workbook.xml.rels -> worksheet part, all from the in-memory
// package map. Returns nil when any part is unavailable (e.g. spilled to a
// temp file for very large workbooks) — callers must treat that as "unknown".
func rawSheetXML(f *excelize.File, sheet string) []byte {
	wbData := pkgPart(f, "xl/workbook.xml")
	if wbData == nil {
		return nil
	}
	var workbook struct {
		Sheets struct {
			Sheet []struct {
				Name string `xml:"name,attr"`
				RID  string `xml:"http://schemas.openxmlformats.org/officeDocument/2006/relationships id,attr"`
			} `xml:"sheet"`
		} `xml:"sheets"`
	}
	if err := xml.Unmarshal(wbData, &workbook); err != nil {
		return nil
	}

	rid := ""
	for _, s := range workbook.Sheets.Sheet {
		if s.Name == sheet {
			rid = s.RID
			break
		}
	}
	if rid == "" {
		return nil
	}

	relsData := pkgPart(f, "xl/_rels/workbook.xml.rels")
	if relsData == nil {
		return nil
	}
	var rels struct {
		Relationship []struct {
			ID     string `xml:"Id,attr"`
			Target string `xml:"Target,attr"`
		} `xml:"Relationship"`
	}
	if err := xml.Unmarshal(relsData, &rels); err != nil {
		return nil
	}

	for _, r := range rels.Relationship {
		if r.ID != rid {
			continue
		}
		target := r.Target
		if strings.HasPrefix(target, "/") {
			target = strings.TrimPrefix(target, "/")
		} else {
			target = "xl/" + target
		}
		return pkgPart(f, target)
	}
	return nil
}

// pkgPart returns the raw bytes of a package part, or nil if absent.
func pkgPart(f *excelize.File, name string) []byte {
	v, ok := f.Pkg.Load(name)
	if !ok {
		return nil
	}
	data, _ := v.([]byte)
	return data
}
