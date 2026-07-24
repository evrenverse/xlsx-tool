package commands

import (
	"fmt"
	"os"
	"path/filepath"
	"runtime"

	"github.com/xuri/excelize/v2"
)

// saveAtomic saves the workbook by writing to a temp file in the SAME
// directory, fsyncing it, and renaming it over the original.
//
// excelize's File.Save truncates the target in place — a crash mid-save
// (kill -9, power loss, full disk) leaves a 0-byte or partial file and the
// original content is gone. The temp+fsync+rename sequence guarantees the
// destination always holds either the old or the new complete workbook.
func saveAtomic(f *excelize.File, path string) error {
	dir := filepath.Dir(path)
	tmp, err := os.CreateTemp(dir, ".xlsx-tool-*.tmp")
	if err != nil {
		return fmt.Errorf("create temp file: %w", err)
	}
	tmpName := tmp.Name()
	// Best-effort cleanup; a no-op once the rename has succeeded.
	defer os.Remove(tmpName)

	if _, err := f.WriteTo(tmp); err != nil {
		tmp.Close()
		return fmt.Errorf("write workbook: %w", err)
	}
	if err := tmp.Sync(); err != nil {
		tmp.Close()
		return fmt.Errorf("sync temp file: %w", err)
	}
	if err := tmp.Close(); err != nil {
		return fmt.Errorf("close temp file: %w", err)
	}

	// Preserve the original file's permissions (CreateTemp uses 0600).
	if fi, err := os.Stat(path); err == nil {
		if err := os.Chmod(tmpName, fi.Mode().Perm()); err != nil {
			return fmt.Errorf("preserve file mode: %w", err)
		}
	}

	if err := os.Rename(tmpName, path); err != nil {
		return fmt.Errorf("replace original file: %w", err)
	}
	if err := syncDirectory(dir); err != nil {
		return fmt.Errorf("sync output directory: %w", err)
	}
	return nil
}

// syncDirectory persists the rename itself on platforms that support syncing
// directory handles. Windows does not expose a portable equivalent.
func syncDirectory(path string) error {
	if runtime.GOOS == "windows" {
		return nil
	}
	directory, err := os.Open(path)
	if err != nil {
		return err
	}
	defer directory.Close()
	return directory.Sync()
}
