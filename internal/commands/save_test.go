package commands

import (
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

func TestSaveAtomic_WritesValidFile(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()

	if err := f.SetCellValue("Sheet1", "A5", "atomic"); err != nil {
		t.Fatalf("set: %v", err)
	}
	if err := saveAtomic(f, path); err != nil {
		t.Fatalf("saveAtomic: %v", err)
	}

	val := testutil.ReadCellValue(t, path, "Sheet1", "A5")
	if val != "atomic" {
		t.Errorf("expected 'atomic' after atomic save, got %q", val)
	}
}

func TestSaveAtomic_LeavesNoTempFiles(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()

	if err := saveAtomic(f, path); err != nil {
		t.Fatalf("saveAtomic: %v", err)
	}

	entries, err := os.ReadDir(filepath.Dir(path))
	if err != nil {
		t.Fatalf("read dir: %v", err)
	}
	for _, e := range entries {
		if strings.Contains(e.Name(), ".tmp") {
			t.Errorf("temp file left behind: %s", e.Name())
		}
	}
}

func TestSaveAtomic_PreservesFileMode(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)
	if err := os.Chmod(path, 0o600); err != nil {
		t.Fatalf("chmod: %v", err)
	}

	f, err := excelize.OpenFile(path)
	if err != nil {
		t.Fatalf("open: %v", err)
	}
	defer f.Close()

	if err := saveAtomic(f, path); err != nil {
		t.Fatalf("saveAtomic: %v", err)
	}

	fi, err := os.Stat(path)
	if err != nil {
		t.Fatalf("stat: %v", err)
	}
	if fi.Mode().Perm() != 0o600 {
		t.Errorf("expected mode 0600 preserved, got %v", fi.Mode().Perm())
	}
}

func TestSyncDirectory(t *testing.T) {
	if err := syncDirectory(t.TempDir()); err != nil {
		t.Fatalf("syncDirectory: %v", err)
	}
}

func TestWriteCommand_UsesAtomicSave_NoTempLeftovers(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout strings.Builder
	if err := runWrite(path, "Sheet1", "A5", "x", &stdout); err != nil {
		t.Fatalf("write: %v", err)
	}

	entries, err := os.ReadDir(filepath.Dir(path))
	if err != nil {
		t.Fatalf("read dir: %v", err)
	}
	for _, e := range entries {
		if e.Name() != filepath.Base(path) {
			t.Errorf("unexpected file in directory after write: %s", e.Name())
		}
	}
}
