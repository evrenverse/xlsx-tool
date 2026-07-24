package main

import (
	"bytes"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

// TestE2E_AgentWorkflow simulates how an AI agent would use xlsx-tool.
func TestE2E_AgentWorkflow(t *testing.T) {
	// Build the binary
	binPath := filepath.Join(t.TempDir(), "xlsx-tool")
	build := exec.Command("go", "build", "-o", binPath, "./cmd/xlsx-tool/")
	build.Dir = filepath.Join("..", "..")
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build failed: %v\n%s", err, out)
	}

	// Create test fixture
	xlsxPath := testutil.CreateSimpleFixture(t)

	// Step 1: Agent runs `xlsx-tool info` (plain text default)
	out := runCLI(t, binPath, "info", xlsxPath)
	if !strings.Contains(out, "File:") {
		t.Errorf("info plain should contain 'File:', got:\n%s", out)
	}
	if !strings.Contains(out, "Sheet1") {
		t.Errorf("info plain should list Sheet1, got:\n%s", out)
	}

	// Step 1b: Agent runs `xlsx-tool info --json` for structured data
	out = runCLI(t, binPath, "info", xlsxPath, "--json")
	var infoResult map[string]any
	if err := json.Unmarshal([]byte(out), &infoResult); err != nil {
		t.Fatalf("info --json should return valid JSON: %v\n%s", err, out)
	}
	sheets := infoResult["sheets"].([]any)
	if len(sheets) != 2 {
		t.Errorf("info should list 2 sheets, got %d", len(sheets))
	}

	// Step 2: Agent runs `xlsx-tool read` (plain text default)
	out = runCLI(t, binPath, "read", xlsxPath, "Sheet1")
	if !strings.Contains(out, "Sheet: Sheet1") {
		t.Errorf("read plain should contain header, got:\n%s", out)
	}
	if !strings.Contains(out, `"Name"`) {
		t.Errorf("read plain should contain 'Name', got:\n%s", out)
	}

	// Step 2b: Agent runs `xlsx-tool read --json` for structured data
	out = runCLI(t, binPath, "read", xlsxPath, "Sheet1", "--json")
	var readResult map[string]any
	if err := json.Unmarshal([]byte(out), &readResult); err != nil {
		t.Fatalf("read --json should return valid JSON: %v\n%s", err, out)
	}
	if readResult["sheet"] != "Sheet1" {
		t.Errorf("read should show Sheet1, got %v", readResult["sheet"])
	}

	// Step 3: Agent runs `xlsx-tool batch` to fill cells
	changes := `[
		{"sheet": "Sheet1", "cell": "A5", "value": "Gamma"},
		{"sheet": "Sheet1", "cell": "B5", "value": 30}
	]`
	jsonPath := filepath.Join(t.TempDir(), "changes.json")
	os.WriteFile(jsonPath, []byte(changes), 0644)

	out = runCLI(t, binPath, "batch", xlsxPath, jsonPath)
	if out == "" {
		t.Error("batch should produce output")
	}

	// Step 4: Agent verifies with single cell read (plain text default)
	out = runCLI(t, binPath, "read", xlsxPath, "Sheet1", "--cell", "A5")
	if !strings.Contains(strings.TrimSpace(out), `"Gamma"`) {
		t.Errorf("single cell plain should contain 'Gamma', got: %s", out)
	}

	// Step 4b: Agent verifies with single cell read (JSON)
	out = runCLI(t, binPath, "read", xlsxPath, "Sheet1", "--cell", "A5", "--json")
	var cellResult map[string]any
	if err := json.Unmarshal([]byte(out), &cellResult); err != nil {
		t.Fatalf("single cell --json should return valid JSON: %v\n%s", err, out)
	}
	if cellResult["value"] != "Gamma" {
		t.Errorf("expected value 'Gamma' after batch write, got %v", cellResult["value"])
	}

	// Step 5: Verify existing data was not corrupted
	val := testutil.ReadCellValue(t, xlsxPath, "Sheet1", "A1")
	if val != "Name" {
		t.Errorf("existing data corrupted: expected 'Name', got %q", val)
	}

	formula := testutil.ReadCellFormula(t, xlsxPath, "Sheet1", "B4")
	if formula != "SUM(B2:B3)" {
		t.Errorf("formula corrupted: expected 'SUM(B2:B3)', got %q", formula)
	}
}

// TestE2E_Validate exercises the built binary's `validate` command end-to-end:
// exit 0 + ok:true on a clean file, exit 1 + ok:false + a JSON body carrying
// the new error and its operand diagnostic on a file with a broken gate.
func TestE2E_Validate(t *testing.T) {
	binPath := filepath.Join(t.TempDir(), "xlsx-tool")
	build := exec.Command("go", "build", "-o", binPath, "./cmd/xlsx-tool/")
	build.Dir = filepath.Join("..", "..")
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build failed: %v\n%s", err, out)
	}

	baseline := testutil.CreateValidateBaselineGateFixture(t)
	produced := testutil.CreateValidateProducedGateFixture(t)

	// Clean run: produced == baseline (gate holds) -> exit 0, ok:true.
	clean := testutil.CreateValidateBaselineGateFixture(t)
	cmd := exec.Command(binPath, "validate", clean, "--baseline", baseline, "--json")
	var cOut, cErr bytes.Buffer
	cmd.Stdout = &cOut
	cmd.Stderr = &cErr
	if err := cmd.Run(); err != nil {
		t.Fatalf("clean validate should exit 0, got err=%v\nstderr: %s", err, cErr.String())
	}
	var cleanRes map[string]any
	if err := json.Unmarshal(cOut.Bytes(), &cleanRes); err != nil {
		t.Fatalf("clean validate --json invalid: %v\n%s", err, cOut.String())
	}
	if cleanRes["ok"] != true {
		t.Fatalf("clean validate must report ok:true, got %v", cleanRes["ok"])
	}

	// Failing run: produced with a broken gate -> exit 1, ok:false, a new error
	// + a gate diagnostic with the operand live values.
	cmd = exec.Command(binPath, "validate", produced, "--baseline", baseline, "--json")
	var fOut, fErr bytes.Buffer
	cmd.Stdout = &fOut
	cmd.Stderr = &fErr
	err := cmd.Run()
	exitErr, ok := err.(*exec.ExitError)
	if !ok || exitErr.ExitCode() != 1 {
		t.Fatalf("failing validate must exit 1, got err=%v\nstderr: %s", err, fErr.String())
	}
	var failRes struct {
		OK     bool `json:"ok"`
		Errors []struct {
			Sheet    string `json:"sheet"`
			Cell     string `json:"cell"`
			Error    string `json:"error"`
			Class    string `json:"class"`
			FixInput string `json:"fix_input"`
		} `json:"errors"`
		Diagnostics []struct {
			Cell     string            `json:"cell"`
			Formula  string            `json:"formula"`
			Operands map[string]string `json:"operands"`
		} `json:"diagnostics"`
	}
	if err := json.Unmarshal(fOut.Bytes(), &failRes); err != nil {
		t.Fatalf("failing validate --json invalid: %v\n%s", err, fOut.String())
	}
	if failRes.OK {
		t.Fatalf("failing validate must report ok:false, got body:\n%s", fOut.String())
	}
	var mustFix int
	for _, e := range failRes.Errors {
		if e.Class == "must_fix" && e.Error == "#VALUE!" {
			mustFix++
		}
	}
	if mustFix != 1 {
		t.Fatalf("want 1 must_fix #VALUE! error, got %+v", failRes.Errors)
	}
	var foundGate bool
	for _, d := range failRes.Diagnostics {
		if d.Cell == "Sheet1!C1" && d.Operands["A1"] == "95" && d.Operands["B1"] == "33" {
			foundGate = true
		}
	}
	if !foundGate {
		t.Fatalf("want a Sheet1!C1 gate diagnostic with A1=95,B1=33, got %+v", failRes.Diagnostics)
	}
}

// TestE2E_ValidateMissingBaseline verifies --baseline is required.
func TestE2E_ValidateMissingBaseline(t *testing.T) {
	binPath := filepath.Join(t.TempDir(), "xlsx-tool")
	build := exec.Command("go", "build", "-o", binPath, "./cmd/xlsx-tool/")
	build.Dir = filepath.Join("..", "..")
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build failed: %v\n%s", err, out)
	}

	produced := testutil.CreateValidateBaselineFixture(t)
	cmd := exec.Command(binPath, "validate", produced)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err == nil {
		t.Fatal("expected non-zero exit when --baseline is missing")
	}
	if !strings.Contains(stderr.String(), "baseline") {
		t.Errorf("expected a 'baseline' required-flag error, got:\n%s", stderr.String())
	}
}

func runCLI(t *testing.T, binPath string, args ...string) string {
	t.Helper()
	cmd := exec.Command(binPath, args...)
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		t.Fatalf("command failed: %v\nstderr: %s\nstdout: %s",
			err, stderr.String(), stdout.String())
	}
	return stdout.String()
}

// TestE2E_ErrorOutput verifies single-line error reporting: exit code 1,
// exactly one "Error:" line, no usage dump, no duplicate message.
func TestE2E_ErrorOutput(t *testing.T) {
	binPath := filepath.Join(t.TempDir(), "xlsx-tool")
	build := exec.Command("go", "build", "-o", binPath, "./cmd/xlsx-tool/")
	build.Dir = filepath.Join("..", "..")
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build failed: %v\n%s", err, out)
	}

	xlsxPath := testutil.CreateSimpleFixture(t)

	cmd := exec.Command(binPath, "read", xlsxPath, "NoSuchSheet")
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	err := cmd.Run()
	if err == nil {
		t.Fatal("expected non-zero exit for unknown sheet")
	}

	combined := stdout.String() + stderr.String()
	if strings.Count(combined, "not found") != 1 {
		t.Errorf("error message must appear exactly once, got:\n%s", combined)
	}
	if strings.Contains(combined, "Usage:") {
		t.Errorf("errors must not dump usage, got:\n%s", combined)
	}
	if !strings.HasPrefix(stderr.String(), "Error: ") {
		t.Errorf("expected single 'Error: ...' line on stderr, got:\n%s", stderr.String())
	}
}

// TestE2E_FindZeroMatches_ExitCode verifies find exits 1 on zero matches
// while still printing the zero-match message / valid JSON.
func TestE2E_FindZeroMatches_ExitCode(t *testing.T) {
	binPath := filepath.Join(t.TempDir(), "xlsx-tool")
	build := exec.Command("go", "build", "-o", binPath, "./cmd/xlsx-tool/")
	build.Dir = filepath.Join("..", "..")
	if out, err := build.CombinedOutput(); err != nil {
		t.Fatalf("build failed: %v\n%s", err, out)
	}

	xlsxPath := testutil.CreateSimpleFixture(t)

	cmd := exec.Command(binPath, "find", xlsxPath, "zzz-no-such-text")
	var stdout, stderr bytes.Buffer
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	err := cmd.Run()

	exitErr, ok := err.(*exec.ExitError)
	if !ok || exitErr.ExitCode() != 1 {
		t.Fatalf("expected exit code 1 on zero matches, got err=%v", err)
	}
	if !strings.Contains(stdout.String(), `No matches found for "zzz-no-such-text".`) {
		t.Errorf("expected zero-match message on stdout, got:\n%s", stdout.String())
	}
	if stderr.String() != "" {
		t.Errorf("zero matches must not print an extra error line, got:\n%s", stderr.String())
	}
}

// TestResolveVersion covers both branches: an injected GoReleaser version wins,
// and an un-injected build falls back without ever returning an empty string.
func TestResolveVersion(t *testing.T) {
	original := version
	t.Cleanup(func() { version = original })

	version = "1.2.3"
	if got := resolveVersion(); got != "1.2.3" {
		t.Errorf("injected version should win, got %q", got)
	}

	version = "dev"
	if got := resolveVersion(); got == "" {
		t.Error("fallback must not return an empty version")
	}
}
