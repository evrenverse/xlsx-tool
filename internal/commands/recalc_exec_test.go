package commands

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/xuri/excelize/v2"
)

// stubRecalc points XLSX_RECALC_BIN at a no-op script that records its args.
// Returns the path of the args log file.
func stubRecalc(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	logPath := filepath.Join(dir, "args.log")
	stub := filepath.Join(dir, "custom-recalc")
	script := "#!/bin/sh\necho \"$@\" > " + logPath + "\nexit 0\n"
	if err := os.WriteFile(stub, []byte(script), 0o755); err != nil {
		t.Fatal(err)
	}
	t.Setenv("XLSX_RECALC_BIN", stub)
	return logPath
}

func TestRunExternalRecalcInvokesBinaryWithPath(t *testing.T) {
	logPath := stubRecalc(t)
	if err := runExternalRecalc("/tmp/some.xlsx"); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}
	got, err := os.ReadFile(logPath)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(got), "/tmp/some.xlsx") {
		t.Fatalf("stub did not receive the xlsx path, got: %q", got)
	}
}

func TestRunExternalRecalcFailsOnNonZeroExit(t *testing.T) {
	dir := t.TempDir()
	stub := filepath.Join(dir, "custom-recalc")
	os.WriteFile(stub, []byte("#!/bin/sh\necho '{\"status\":\"error\",\"error\":\"boom\"}'\nexit 1\n"), 0o755)
	t.Setenv("XLSX_RECALC_BIN", stub)
	err := runExternalRecalc("/tmp/some.xlsx")
	if err == nil || !strings.Contains(err.Error(), "boom") {
		t.Fatalf("expected error containing stub output, got: %v", err)
	}
}

func TestRunExternalRecalcMissingLibreOffice(t *testing.T) {
	t.Setenv("XLSX_RECALC_BIN", "")
	t.Setenv("PATH", t.TempDir())
	err := runExternalRecalc("/tmp/some.xlsx")
	if err == nil || !strings.Contains(err.Error(), "LibreOffice not found") {
		t.Fatalf("expected missing LibreOffice error, got: %v", err)
	}
}

func TestRecalcWithLibreOffice(t *testing.T) {
	if _, err := findLibreOffice(); err != nil {
		t.Skip("LibreOffice is not installed")
	}

	path := filepath.Join(t.TempDir(), "formula.xlsx")
	workbook := excelize.NewFile()
	if err := workbook.SetCellValue("Sheet1", "A1", 1); err != nil {
		t.Fatal(err)
	}
	if err := workbook.SetCellValue("Sheet1", "A2", 2); err != nil {
		t.Fatal(err)
	}
	if err := workbook.SetCellFormula("Sheet1", "A3", "=SUM(A1:A2)"); err != nil {
		t.Fatal(err)
	}
	if err := workbook.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatal(err)
	}

	if err := recalcWithLibreOffice(path); err != nil {
		t.Fatal(err)
	}
	recalculated, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer recalculated.Close()
	value, err := recalculated.GetCellValue("Sheet1", "A3")
	if err != nil {
		t.Fatal(err)
	}
	if value != "3" {
		t.Fatalf("expected cached formula value 3, got %q", value)
	}
}
