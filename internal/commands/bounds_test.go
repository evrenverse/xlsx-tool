package commands

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
	"time"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

// runBounded fails the test if fn does not return within the deadline —
// guards against per-cell iteration over a lying, huge stored dimension.
func runBounded(t *testing.T, what string, fn func() error) {
	t.Helper()
	done := make(chan error, 1)
	go func() { done <- fn() }()
	select {
	case err := <-done:
		if err != nil {
			t.Fatalf("%s failed: %v", what, err)
		}
	case <-time.After(15 * time.Second):
		t.Fatalf("%s hung on lying oversized dimension", what)
	}
}

func TestRead_OversizedDimension_CompletesQuickly(t *testing.T) {
	path := testutil.CreateDimensionLieFixture(t, "A1:XFD1048576")

	var stdout bytes.Buffer
	runBounded(t, "full-sheet read", func() error {
		return runRead(path, "Sheet1", readOptions{}, &stdout)
	})

	if !strings.Contains(stdout.String(), `"Alpha"`) {
		t.Errorf("expected actual data in output, got:\n%s", stdout.String())
	}
}

func TestWrite_OversizedDimension_CompletesQuickly(t *testing.T) {
	path := testutil.CreateDimensionLieFixture(t, "A1:XFD1048576")

	var stdout bytes.Buffer
	runBounded(t, "write", func() error {
		return runWrite(path, "Sheet1", "B2", "50", &stdout)
	})

	// write no longer recalcs — B4's cache stays whatever it was (here:
	// absent, since excelize never computed it) regardless of the lying
	// dimension. validate owns recalculation via LibreOffice.
	val := testutil.ReadCellValue(t, path, "Sheet1", "B4")
	if val != "" {
		t.Errorf("expected no cached value before validate recalcs, got %q", val)
	}
}

func TestInfo_OversizedDimension_CompletesQuickly(t *testing.T) {
	path := testutil.CreateDimensionLieFixture(t, "A1:XFD1048576")

	var stdout bytes.Buffer
	runBounded(t, "info (formula counting)", func() error {
		return runInfo(path, false, &stdout)
	})
}

func TestRead_UndersizedDimension_ShowsAllData(t *testing.T) {
	path := testutil.CreateDimensionLieFixture(t, "A1")

	var stdout bytes.Buffer
	err := runRead(path, "Sheet1", readOptions{}, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, `"Beta"`) {
		t.Errorf("undersized dimension must not hide row 3, got:\n%s", output)
	}
	if !strings.Contains(output, "20") {
		t.Errorf("undersized dimension must not hide column B, got:\n%s", output)
	}
	if !strings.Contains(output, "=SUM(B2:B3)") {
		t.Errorf("undersized dimension must not hide the formula cell, got:\n%s", output)
	}
}

func TestInfo_UndersizedDimension_CountsFormulas(t *testing.T) {
	path := testutil.CreateDimensionLieFixture(t, "A1")

	var stdout bytes.Buffer
	if err := runInfo(path, true, &stdout); err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	var result map[string]any
	if err := json.Unmarshal(stdout.Bytes(), &result); err != nil {
		t.Fatalf("info output should be valid JSON: %v\n%s", err, stdout.String())
	}

	sheets := result["sheets"].([]any)
	first := sheets[0].(map[string]any)
	if first["formulas"] != float64(1) {
		t.Errorf("expected 1 formula counted despite undersized dimension, got %v",
			first["formulas"])
	}
}
