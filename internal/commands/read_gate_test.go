package commands

import (
	"bytes"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

// createLargeFixture writes a workbook whose single sheet spans more than
// maxAutoMultiSheetCells cells (data reaching AX300 = 50 cols x 300 rows).
func createLargeFixture(t *testing.T) string {
	t.Helper()
	f := excelize.NewFile()
	defer f.Close()
	if err := f.SetCellValue("Sheet1", "AX300", 1); err != nil {
		t.Fatalf("set far cell: %v", err)
	}
	path := filepath.Join(t.TempDir(), "large.xlsx")
	if err := f.SaveAs(path); err != nil {
		t.Fatalf("save large fixture: %v", err)
	}
	return path
}

func TestReadCommand_AllSheets_GatedWhenLarge(t *testing.T) {
	path := createLargeFixture(t)

	var out bytes.Buffer
	err := runRead(path, "", readOptions{allSheets: true, jsonOut: true}, &out)
	if err == nil {
		t.Fatal("expected --all-sheets on a large workbook to be gated")
	}
	if !strings.Contains(err.Error(), "--confirm-large") {
		t.Errorf("gate error should mention --confirm-large, got: %v", err)
	}
}

func TestReadCommand_AllSheets_ConfirmLargeBypassesGate(t *testing.T) {
	path := createLargeFixture(t)

	var out bytes.Buffer
	err := runRead(path, "", readOptions{allSheets: true, jsonOut: true, confirmLarge: true}, &out)
	if err != nil {
		t.Fatalf("--confirm-large should allow the read: %v", err)
	}
	if out.Len() == 0 {
		t.Error("expected output from confirmed large read")
	}
}

func TestReadCommand_AllSheets_SmallWorkbookNotGated(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var out bytes.Buffer
	if err := runRead(path, "", readOptions{allSheets: true, jsonOut: true}, &out); err != nil {
		t.Fatalf("small workbook --all-sheets should not be gated: %v", err)
	}
}
