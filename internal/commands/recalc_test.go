package commands

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

// requireRealRecalc skips when LibreOffice is absent. The contract under test
// is what a genuine recalculation does, so a stub would prove nothing here.
func requireRealRecalc(t *testing.T) {
	t.Helper()
	t.Setenv("XLSX_RECALC_BIN", "")
	for _, candidate := range []string{"soffice", "libreoffice"} {
		if _, err := exec.LookPath(candidate); err == nil {
			return
		}
	}
	t.Skip("LibreOffice not installed; skipping real recalculation test")
}

// TestRecalc_ContentErrorsDoNotFail is the contract: recalc reports whether
// recalculation worked, never what the formulas evaluated to. The same
// workbook that recalculates fine here is rejected by validate below, which
// is exactly the separation callers depend on — a workbook with unfilled
// required cells mid-pipeline must not read as an infrastructure failure.
func TestRecalc_ContentErrorsDoNotFail(t *testing.T) {
	requireRealRecalc(t)

	produced := testutil.CreateValidateProducedNewErrFixture(t)
	if err := NewRecalcCmd().RunE(nil, []string{produced}); err != nil {
		t.Fatalf("recalc must not fail on a workbook with formula errors: %v", err)
	}

	baseline := testutil.CreateValidateBaselineFixture(t)
	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatalf("validate: %v", err)
	}
	if res.OK {
		t.Fatal("fixture no longer carries a content error; the test proves nothing")
	}
}

// TestRecalc_PersistsFreshCaches covers the in-place half of the contract: the
// file itself is rewritten, so a later reader sees the recalculated values.
func TestRecalc_PersistsFreshCaches(t *testing.T) {
	requireRealRecalc(t)

	path := testutil.CreateChainedFormulaFixture(t)
	before, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	info, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}

	if err := NewRecalcCmd().RunE(nil, []string{path}); err != nil {
		t.Fatalf("recalc: %v", err)
	}

	after, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	if string(before) == string(after) {
		t.Fatal("recalc left the file byte-identical; nothing was written back")
	}
	got, err := os.Stat(path)
	if err != nil {
		t.Fatal(err)
	}
	if got.Mode().Perm() != info.Mode().Perm() {
		t.Errorf("permissions changed: %v -> %v", info.Mode().Perm(), got.Mode().Perm())
	}
}

// TestRecalc_MissingFileFails keeps a typo from being reported as a success.
func TestRecalc_MissingFileFails(t *testing.T) {
	err := NewRecalcCmd().RunE(nil, []string{filepath.Join(t.TempDir(), "absent.xlsx")})
	if err == nil {
		t.Fatal("expected an error for a missing file")
	}
	if !strings.Contains(err.Error(), "open file") {
		t.Errorf("unhelpful error for a missing file: %v", err)
	}
}

// TestRecalc_InfrastructureFailurePropagates is the other side of the
// contract: a broken recalculation must exit non-zero.
func TestRecalc_InfrastructureFailurePropagates(t *testing.T) {
	dir := t.TempDir()
	stub := filepath.Join(dir, "failing-recalc")
	script := "#!/bin/sh\necho 'recalculation exploded' >&2\nexit 1\n"
	if err := os.WriteFile(stub, []byte(script), 0o755); err != nil {
		t.Fatal(err)
	}
	t.Setenv("XLSX_RECALC_BIN", stub)

	err := NewRecalcCmd().RunE(nil, []string{testutil.CreateSimpleFixture(t)})
	if err == nil {
		t.Fatal("expected an error when recalculation fails")
	}
	if !strings.Contains(err.Error(), "recalculation exploded") {
		t.Errorf("error should carry the stub's diagnostic, got: %v", err)
	}
}
