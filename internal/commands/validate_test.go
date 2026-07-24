package commands

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
	"github.com/xuri/excelize/v2"
)

// --- Check 1: new formula errors -------------------------------------------

func TestValidate_NewFormulaError(t *testing.T) {
	stubRecalc(t)
	// baseline: a clean workbook. produced: same but an edit introduced #VALUE!.
	// B1 was turned into text and B3=VALUE(B1) recalcs to #VALUE! (the fixture
	// carries the cached error, as xlsx-recalc/LibreOffice would persist it).
	// B1 is an unlocked, non-numeric, agent-fillable leaf -> the error is must_fix.
	produced := testutil.CreateValidateProducedNewErrFixture(t)
	baseline := testutil.CreateValidateBaselineFixture(t)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if res.OK {
		t.Fatal("expected ok=false")
	}
	var mustFix []ErrorCell
	for _, e := range res.Errors {
		if e.Class == ClassMustFix {
			mustFix = append(mustFix, e)
		}
	}
	if len(mustFix) != 1 {
		t.Fatalf("want 1 must_fix error, got %d (%+v)", len(mustFix), res.Errors)
	}
	if mustFix[0].Error != "#VALUE!" {
		t.Fatalf("got %q", mustFix[0].Error)
	}
	if !strings.Contains(mustFix[0].FixInput, "B1") {
		t.Fatalf("want fix_input naming B1, got %q", mustFix[0].FixInput)
	}
	// B3 is a GENUINE cached error (t="e"), not a template placeholder literal
	// (t="str"). Even though its string value "#VALUE!" also happens to be a
	// placeholderMarkers entry, the stored cell type must gate it out of
	// scanPlaceholderErrors so it is never mislabelled Placeholder: true.
	if mustFix[0].Placeholder {
		t.Fatalf("a genuine t=\"e\" cached error must not be flagged Placeholder, got %+v", mustFix[0])
	}
}

// --- False-positive guards / latent-#VALUE! detection ----------------------
//
// The latent-#VALUE! text-arithmetic scan (scanLatentValueErrors) that used
// to compensate for excelize's calc engine silently coercing non-numeric text
// to 0 in arithmetic is gone: LibreOffice's recalc materialises every real
// #VALUE! into the cache, so scanFormulaErrors (reading cached values) now
// catches what that scan used to catch, and the false-positive guards for it
// no longer apply (there is no live-evaluation path left to false-positive on).

func TestValidate_CleanFileNoErrors(t *testing.T) {
	stubRecalc(t)
	// A clean file with no error cells at all: zero errors, ok=true. (The blank
	// baseline template carries no #WERT! placeholder and no real errors.)
	baseline := testutil.CreateValidateBaselineFixture(t)

	res, err := runValidate(baseline, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Errors) != 0 {
		t.Fatalf("a clean file must have 0 errors, got %d", len(res.Errors))
	}
	if !res.OK {
		t.Fatalf("a clean file must be ok=true, got %+v", res)
	}
}

// --- F10a: pre-existing template placeholder (#WERT!) -----------------------

func TestValidate_PlaceholderPreexistingFixable(t *testing.T) {
	stubRecalc(t)
	// A representative template shape: C20 = IF(A16>15, B20*D20, "#WERT!") with an
	// EMPTY gating input A16, so C20 evaluates to the literal "#WERT!" and the
	// total G40 propagates it. The error is PRE-EXISTING in the blank template,
	// so the produced and baseline can be the SAME file — the OLD diff-vs-baseline
	// logic excluded it (ok:true). The reworked validate must surface it as a
	// must_fix naming the gating input A16.
	wb := testutil.CreateValidatePlaceholderFixture(t, nil) // A16 empty
	baseline := testutil.CreateValidatePlaceholderFixture(t, nil)

	res, err := runValidate(wb, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if res.OK {
		t.Fatalf("a pre-existing fixable placeholder must force ok=false, got %+v", res)
	}

	var mustFix []ErrorCell
	for _, e := range res.Errors {
		if e.Class == ClassMustFix {
			mustFix = append(mustFix, e)
		}
	}
	if len(mustFix) == 0 {
		t.Fatalf("want at least one must_fix error, got errors=%+v", res.Errors)
	}

	// The placeholder cell C20 must be reported as a must_fix whose fix_input
	// names the gating input A16, flagged as a template placeholder.
	var c20 *ErrorCell
	for i := range mustFix {
		if mustFix[i].Cell == "C20" {
			c20 = &mustFix[i]
			break
		}
	}
	if c20 == nil {
		t.Fatalf("want a must_fix for Sheet1!C20, got %+v", mustFix)
	}
	if !c20.Placeholder {
		t.Fatalf("C20 must be flagged as an IF(...,\"#WERT!\") placeholder, got %+v", *c20)
	}
	if !strings.Contains(c20.FixInput, "A16") {
		t.Fatalf("want fix_input naming A16, got %q", c20.FixInput)
	}
}

func TestValidate_PlaceholderResolvedAfterFillingGate(t *testing.T) {
	stubRecalc(t)
	// Filling the gating input A16 with a value that passes the gate (>15) makes
	// C20 (and G40) compute a number -> no placeholder, no error -> ok=true.
	wb := testutil.CreateValidatePlaceholderFixture(t, 20) // A16=20 passes A16>15
	baseline := testutil.CreateValidatePlaceholderFixture(t, nil)

	res, err := runValidate(wb, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if !res.OK {
		t.Fatalf("filling the gating input must yield ok=true, got %+v", res)
	}
	for _, e := range res.Errors {
		if e.Class == ClassMustFix {
			t.Fatalf("no must_fix expected after filling the gate, got %+v", e)
		}
	}
}

func TestValidate_ExternalErrorDoesNotBlockOK(t *testing.T) {
	stubRecalc(t)
	// An error tracing ONLY to a locked/template-owned cell is informational.
	// and must NOT block ok=true.
	wb := testutil.CreateValidateExternalErrorFixture(t)
	baseline := testutil.CreateValidateExternalErrorFixture(t)

	res, err := runValidate(wb, baseline)
	if err != nil {
		t.Fatal(err)
	}

	var external, mustFix []ErrorCell
	for _, e := range res.Errors {
		switch e.Class {
		case ClassExternal:
			external = append(external, e)
		case ClassMustFix:
			mustFix = append(mustFix, e)
		}
	}
	if len(external) == 0 {
		t.Fatalf("want the locked-chain error classified external, got %+v", res.Errors)
	}
	if len(mustFix) != 0 {
		t.Fatalf("a locked-only chain must NOT be must_fix, got %+v", mustFix)
	}
	if !res.OK {
		t.Fatalf("an external-only error must keep ok=true, got %+v", res)
	}
}

func TestValidate_MissingProducedFileReturnsError(t *testing.T) {
	stubRecalc(t)
	// A missing produced file (an OpenFile failure, after recalc succeeds on
	// nothing to recalc — the stub is a no-op) must return a Go error — never
	// a silent ok:true.
	baseline := testutil.CreateValidateBaselineFixture(t)
	_, err := runValidate(filepath.Join(t.TempDir(), "does-not-exist.xlsx"), baseline)
	if err == nil {
		t.Fatal("a missing produced file must return an error, not a silent ok:true")
	}
}

func TestValidate_RecalcFailureReturnsError(t *testing.T) {
	// Drive the actual "recalc produced file: %w" branch in runValidate: a
	// failing xlsx-recalc stub (non-zero exit) must surface as a wrapped Go
	// error naming that branch — not be swallowed, and not be misattributed to
	// the later OpenFile step. Mirrors the failing-stub construction in
	// TestRunExternalRecalcFailsOnNonZeroExit (recalc_exec_test.go).
	dir := t.TempDir()
	stub := filepath.Join(dir, "xlsx-recalc")
	script := "#!/bin/sh\necho '{\"status\":\"error\",\"error\":\"boom\"}'\nexit 1\n"
	if err := os.WriteFile(stub, []byte(script), 0o755); err != nil {
		t.Fatal(err)
	}
	t.Setenv("XLSX_RECALC_BIN", stub)

	produced := testutil.CreateValidateBaselineFixture(t)
	baseline := testutil.CreateValidateBaselineFixture(t)

	_, err := runValidate(produced, baseline)
	if err == nil {
		t.Fatal("a failing recalc must return an error, not a silent ok:true")
	}
	if !strings.Contains(err.Error(), "recalc produced file") {
		t.Fatalf("want error naming the recalc-produced-file branch, got: %v", err)
	}
}

// --- Check 2: frozen formulas ----------------------------------------------

func TestValidate_FrozenFormula(t *testing.T) {
	stubRecalc(t)
	// baseline B3 is =B1+B2; produced B3 is the literal 5 (formula replaced).
	produced := testutil.CreateValidateProducedFrozenFixture(t)
	baseline := testutil.CreateValidateBaselineFixture(t)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.FrozenFormulas) != 1 || res.FrozenFormulas[0].Cell != "B3" {
		t.Fatalf("want B3 frozen, got %+v", res.FrozenFormulas)
	}
	if res.FrozenFormulas[0].BaselineFormula != "B1+B2" {
		t.Fatalf("want baseline formula B1+B2, got %q", res.FrozenFormulas[0].BaselineFormula)
	}
	if res.OK {
		t.Fatal("a frozen formula must force ok=false")
	}
}

// --- Check 3: unfilled fill-targets (protected-only, aggregated, informational) --

func TestValidate_UnprotectedSheetNeverFlagged(t *testing.T) {
	stubRecalc(t)
	// baseline has an unlocked, empty, non-formula input cell C5 on an
	// UNPROTECTED sheet; produced left it empty too. Unlocked-on-an-
	// unprotected-sheet carries no author intent (every cell there is
	// unlocked), so the sheet is skipped entirely: Total stays 0 and, since
	// unfilled never gates ok anyway, the result is ok=true.
	produced := testutil.CreateValidateProducedUnfilledFixture(t)
	baseline := testutil.CreateValidateBaselineInputFixture(t)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if res.Unfilled.Total != 0 {
		t.Fatalf("want Total=0 for an unprotected sheet, got %+v", res.Unfilled)
	}
	if len(res.Unfilled.BySheet) != 0 {
		t.Fatalf("want no sheet groups for an unprotected sheet, got %+v", res.Unfilled.BySheet)
	}
	if !res.OK {
		t.Fatalf("unfilled targets must never block ok, got %+v", res)
	}
}

func TestValidate_FilledTargetNotFlagged(t *testing.T) {
	stubRecalc(t)
	// When the produced file fills C5, it is no longer an unfilled target.
	baseline := testutil.CreateValidateBaselineInputFixture(t)
	produced := testutil.CreateValidateProducedUnfilledFixture(t)

	// Fill C5 in the produced copy.
	f, err := excelize.OpenFile(produced, excelize.Options{RawCellValue: true})
	if err != nil {
		t.Fatal(err)
	}
	if err := f.SetCellValue("Sheet1", "C5", 42); err != nil {
		t.Fatal(err)
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	f.Close()

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if res.Unfilled.Total != 0 {
		t.Fatalf("a filled target must not be flagged, got %+v", res.Unfilled)
	}
}

func TestValidate_ProtectedSheetUnfilledAggregated(t *testing.T) {
	stubRecalc(t)
	// A PROTECTED sheet with 2 empty unlocked slots: aggregated into a single
	// sheet group, both addresses sampled, not truncated. Informational only
	// -> must not block ok.
	baseline := testutil.CreateValidateProtectedUnfilledFixture(t, 2)
	produced := testutil.CreateValidateProtectedUnfilledFixture(t, 2)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if res.Unfilled.Total != 2 {
		t.Fatalf("want Total=2, got %+v", res.Unfilled)
	}
	if len(res.Unfilled.BySheet) != 1 {
		t.Fatalf("want exactly one sheet group, got %+v", res.Unfilled.BySheet)
	}
	group := res.Unfilled.BySheet[0]
	if group.Sheet != "Sheet1" || group.Count != 2 {
		t.Fatalf("want Sheet1 count=2, got %+v", group)
	}
	if len(group.Sample) != 2 || group.Sample[0] != "C5" || group.Sample[1] != "C6" {
		t.Fatalf("want sample [C5 C6], got %v", group.Sample)
	}
	if res.Unfilled.Truncated {
		t.Fatalf("2 targets with a sample of 2 must not be truncated, got %+v", res.Unfilled)
	}
	if !res.OK {
		t.Fatalf("unfilled targets must never block ok, got %+v", res)
	}
}

func TestValidate_ProtectedSheetUnfilledTruncated(t *testing.T) {
	stubRecalc(t)
	// A PROTECTED sheet with 12 empty unlocked slots: the sample caps at 10
	// and truncated becomes true, but ok stays unaffected.
	baseline := testutil.CreateValidateProtectedUnfilledFixture(t, 12)
	produced := testutil.CreateValidateProtectedUnfilledFixture(t, 12)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Unfilled.BySheet) != 1 {
		t.Fatalf("want exactly one sheet group, got %+v", res.Unfilled.BySheet)
	}
	group := res.Unfilled.BySheet[0]
	if group.Count != 12 {
		t.Fatalf("want count=12, got %+v", group)
	}
	if len(group.Sample) != 10 {
		t.Fatalf("want sample capped at 10, got %d (%v)", len(group.Sample), group.Sample)
	}
	if res.Unfilled.Total != 12 {
		t.Fatalf("want Total=12, got %+v", res.Unfilled)
	}
	if !res.Unfilled.Truncated {
		t.Fatalf("want Truncated=true once count exceeds the sample, got %+v", res.Unfilled)
	}
	if !res.OK {
		t.Fatalf("unfilled targets must never block ok, got %+v", res)
	}
}

func TestValidate_UnfilledOnlyIssueIsOKAndExitsZero(t *testing.T) {
	stubRecalc(t)
	// A file whose ONLY defect is unfilled targets on a protected sheet must
	// be ok=true end to end, and the CLI command must exit 0 (RunE returns a
	// nil error, which main.go maps to os.Exit(0)).
	baseline := testutil.CreateValidateProtectedUnfilledFixture(t, 2)
	produced := testutil.CreateValidateProtectedUnfilledFixture(t, 2)

	cmd := NewValidateCmd()
	var out bytes.Buffer
	cmd.SetOut(&out)
	cmd.SetErr(&out)
	cmd.SetArgs([]string{produced, "--baseline", baseline})

	if err := cmd.Execute(); err != nil {
		t.Fatalf("want the command to exit 0 when the only issue is unfilled targets, got err=%v output=%s",
			err, out.String())
	}
}

// --- Check 4: operand diagnostics ------------------------------------------

func TestValidate_OperandDiagnostics(t *testing.T) {
	stubRecalc(t)
	// produced has C1=IF(A1=B1,D1,"") feeding a #VALUE! total E1=VALUE(C1);
	// A1=95, B1=33 -> gate fails -> C1="" -> VALUE("") -> #VALUE! (the fixture
	// carries the cached error, as xlsx-recalc/LibreOffice would persist it).
	produced := testutil.CreateValidateProducedGateFixture(t)
	baseline := testutil.CreateValidateBaselineGateFixture(t)

	res, err := runValidate(produced, baseline)
	if err != nil {
		t.Fatal(err)
	}
	if len(res.Diagnostics) == 0 {
		t.Fatalf("expected operand diagnostics, got %+v", res)
	}

	// Find the gate diagnostic (the cell whose formula carries the IF gate).
	var gate *Diagnostic
	for i := range res.Diagnostics {
		if res.Diagnostics[i].Cell == "Sheet1!C1" {
			gate = &res.Diagnostics[i]
			break
		}
	}
	if gate == nil {
		t.Fatalf("want a diagnostic for Sheet1!C1, got %+v", res.Diagnostics)
	}
	if gate.Operands["A1"] != "95" || gate.Operands["B1"] != "33" {
		t.Fatalf("want A1=95,B1=33 operands, got %+v", gate.Operands)
	}
	if gate.Formula != `IF(A1=B1,D1,"")` {
		t.Fatalf("want gate formula reported, got %q", gate.Formula)
	}
}

// --- Recalc wiring (spec §8) -------------------------------------------------

func TestValidate_RecalcsProducedBeforeOpening(t *testing.T) {
	// validate no longer recalcs/saves in-process (excelize was too slow/
	// incomplete on real workbooks) — it execs xlsx-recalc on the PRODUCED
	// path IN PLACE before opening the file, and persistence of fresh caches
	// is entirely xlsx-recalc's responsibility from that point on. Verify the
	// wiring: runValidate must invoke the external recalc with the produced
	// path, and must NEVER invoke it with the baseline path.
	logPath := stubRecalc(t)
	produced := testutil.CreateValidateBaselineFixture(t)
	baseline := testutil.CreateValidateBaselineFixture(t)

	if _, err := runValidate(produced, baseline); err != nil {
		t.Fatal(err)
	}

	got, err := os.ReadFile(logPath)
	if err != nil {
		t.Fatalf("recalc stub was not invoked: %v", err)
	}
	args := strings.TrimSpace(string(got))
	if !strings.Contains(args, produced) {
		t.Fatalf("want recalc invoked with the produced path %q, got args %q", produced, args)
	}
	if strings.Contains(args, baseline) {
		t.Fatalf("recalc must never be invoked with the baseline path, got args %q", args)
	}
}

func TestValidate_BaselineNotModified(t *testing.T) {
	stubRecalc(t)
	// The baseline is read-only and must never be saved/altered.
	produced := testutil.CreateValidateProducedNewErrFixture(t)
	baseline := testutil.CreateValidateBaselineFixture(t)

	beforeFormula := testutil.ReadCellFormula(t, baseline, "Sheet1", "B3")
	beforeValue := testutil.ReadCellValue(t, baseline, "Sheet1", "B1")

	if _, err := runValidate(produced, baseline); err != nil {
		t.Fatal(err)
	}

	if got := testutil.ReadCellFormula(t, baseline, "Sheet1", "B3"); got != beforeFormula {
		t.Fatalf("baseline formula changed: %q -> %q", beforeFormula, got)
	}
	if got := testutil.ReadCellValue(t, baseline, "Sheet1", "B1"); got != beforeValue {
		t.Fatalf("baseline value changed: %q -> %q", beforeValue, got)
	}
}
