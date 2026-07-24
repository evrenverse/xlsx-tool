package commands

import (
	"fmt"
	"io"
	"regexp"
	"strings"

	"github.com/evrenverse/xlsx-tool/internal/format"
	"github.com/spf13/cobra"
	"github.com/xuri/excelize/v2"
)

// Error classes assigned to every error cell that validate enumerates.
const (
	// ClassMustFix marks an error whose operand chain reaches an unlocked,
	// agent-fillable input the agent can resolve (empty, or a value that fails
	// the formula's gate). A must_fix forces ok=false.
	ClassMustFix = "must_fix"
	// ClassExternal marks an error whose operand chain reaches ONLY locked,
	// template-owned cells — nothing the agent can change. It is informational and
	// never blocks ok.
	ClassExternal = "external"
)

// placeholderMarkers are the literal strings a template uses as a "fill-me"
// marker in an IF false-branch (e.g. IF(B16>15, …, "#WERT!")). When the gate
// fails the cell evaluates to one of these literals. LibreOffice's recalc
// caches the literal string verbatim — it is NOT a recognised Excel error
// value (errorValues), so scanFormulaErrors misses it. We detect it explicitly
// and treat the cell as a template required-input marker.
var placeholderMarkers = []string{"#WERT!", "#VALUE!"}

// UnfilledSummary is the aggregated, truncation-safe view of unfilled
// template slots. A raw address list explodes on big templates (observed:
// 101,838 entries / 3.8 MB JSON on a large 32-sheet workbook) and would
// blow the calling agent's context; per-sheet counts plus a small sample
// carry strictly more signal in a few hundred bytes.
type UnfilledSummary struct {
	Total     int                  `json:"total"`
	Truncated bool                 `json:"truncated"`
	BySheet   []UnfilledSheetGroup `json:"by_sheet"`
}

// UnfilledSheetGroup is one sheet's unfilled-slot count with example cells.
type UnfilledSheetGroup struct {
	Sheet  string   `json:"sheet"`
	Count  int      `json:"count"`
	Sample []string `json:"sample"`
}

// ValidateResult is the machine-readable outcome of `xlsx-tool validate`.
//
// ok is true iff there are no must_fix errors and no frozen_formulas.
// external errors, unfilled_targets, and diagnostics are explanatory only and
// never affect ok.
type ValidateResult struct {
	OK             bool         `json:"ok"`
	Errors         []ErrorCell  `json:"errors"`
	FrozenFormulas []FrozenCell `json:"frozen_formulas"`
	// Unfilled aggregates still-empty template input slots (unlocked, non-
	// formula, empty in baseline AND produced) on PROTECTED sheets only.
	// Informational: it never gates ok — mandatory gaps surface as must_fix
	// placeholder chains with a fix_input instead. On unprotected sheets
	// "unlocked" carries no author intent (every cell is unlocked), so they
	// are not scanned; and huge templates legitimately ship thousands of
	// reserve rows that stay empty by design.
	Unfilled    UnfilledSummary `json:"unfilled_targets"`
	Diagnostics []Diagnostic    `json:"diagnostics"`
}

// ErrorCell is one error-valued (or placeholder) formula cell, classified by
// where its operand chain leads.
//
//   - Class is must_fix when the chain reaches an unlocked, agent-fillable input
//     the agent can resolve; external when it reaches only locked/template cells.
//   - FixInput (must_fix only) names the specific input cell to set, sheet-
//     qualified (e.g. "Kalkulation Tageskraft!B16").
//   - Placeholder is true when the cell is a template required-input marker
//     (an IF(…, "#WERT!"/"#VALUE!") false-branch literal) rather than a calc
//     error.
type ErrorCell struct {
	Sheet       string `json:"sheet"`
	Cell        string `json:"cell"`
	Error       string `json:"error"`
	Formula     string `json:"formula"`
	Class       string `json:"class"`
	FixInput    string `json:"fix_input,omitempty"`
	Placeholder bool   `json:"placeholder,omitempty"`
}

// FrozenCell is a baseline formula cell that became a literal in produced.
type FrozenCell struct {
	Sheet           string `json:"sheet"`
	Cell            string `json:"cell"`
	BaselineFormula string `json:"baseline_formula"`
}

// Diagnostic explains a failing formula: a cell it references (one hop), that
// cell's formula text, and the live values of the operands that formula reads.
type Diagnostic struct {
	Cell     string            `json:"cell"`
	Formula  string            `json:"formula"`
	Operands map[string]string `json:"operands"`
}

// errKey is the address key (Sheet!Cell) used to dedupe and to look up
// baseline errors.
func errKey(e FormulaError) string { return e.Sheet + "!" + e.Cell }

// cellRefRe matches A1-style cell references, optionally sheet-qualified and
// with $ anchors, inside a formula string.
var cellRefRe = regexp.MustCompile(`([A-Za-z_][A-Za-z0-9_.]*!)?\$?[A-Z]{1,3}\$?[0-9]{1,7}`)

// runValidate validates a produced workbook against the file the task received.
//
// It recalculates the produced workbook via LibreOffice BEFORE
// opening it, then runs the four domain-free checks by reading only cached
// values — never re-evaluating. LibreOffice already persisted the fresh
// computed formula caches to disk (spec §8), so no separate save step is
// needed here. The baseline is opened read-only and never written.
func runValidate(producedPath, baselinePath string) (ValidateResult, error) {
	// Recalculate the produced workbook IN PLACE via LibreOffice
	// BEFORE opening it. excelize's own calc engine is both too slow
	// (minutes on 50k-formula workbooks — the agent's 2-min Bash timeout
	// kills it) and incomplete (cross-sheet IF control chains yield wrong
	// results). LibreOffice writes correct cached values in ~seconds; every
	// scan below only READS caches. This also persists fresh computed values
	// for downstream readers (spec §8) — no separate save step needed.
	if err := runExternalRecalc(producedPath); err != nil {
		return ValidateResult{}, fmt.Errorf("recalc produced file: %w", err)
	}

	prod, err := openWorkbook(producedPath, true)
	if err != nil {
		return ValidateResult{}, fmt.Errorf("open produced file: %w", err)
	}
	defer prod.Close()

	base, err := openWorkbook(baselinePath, true)
	if err != nil {
		return ValidateResult{}, fmt.Errorf("open baseline file: %w", err)
	}
	defer base.Close()

	// Enumerate ALL error cells in the recalced produced workbook (NOT a
	// diff-vs-baseline). Two sources are merged:
	//   - scanFormulaErrors: cells whose recalced cached value is a known
	//     Excel error value.
	//   - scanPlaceholderErrors: template required-input markers, i.e. cells
	//     whose IF false-branch returns the literal "#WERT!"/"#VALUE!" when an
	//     unfilled gating input fails the gate. Not a recognised Excel error
	//     value, so scanFormulaErrors misses it.
	//
	// The latent-#VALUE! text-arithmetic scan that used to run here is gone:
	// LibreOffice materialises every real #VALUE! into the cache, so
	// scanFormulaErrors now catches what excelize's calc engine used to miss.
	prodErrs := mergeErrors(
		scanFormulaErrors(prod),
		scanPlaceholderErrors(prod),
	)

	// Which enumerated cells are template placeholder markers (their IF
	// false-branch returned a literal "#WERT!"/"#VALUE!") so each ErrorCell is
	// labelled regardless of which scan first reported it.
	placeholderKeys := placeholderKeySet(prod)

	// Classify every enumerated error by tracing its operand chain to leaf
	// inputs: must_fix when it reaches an unlocked, agent-fillable input the
	// agent can resolve; external when it reaches only locked/template cells.
	// Classification reads the produced workbook's own lock state — NOT a
	// diff-vs-baseline — so a pre-existing fixable error is reported, never
	// suppressed. The baseline drives only the frozen / unfilled checks below.
	errorCells := make([]ErrorCell, 0, len(prodErrs))
	for _, e := range prodErrs {
		errorCells = append(errorCells, classifyError(prod, e, placeholderKeys[errKey(e)]))
	}

	res := ValidateResult{
		Errors:         errorCells,
		FrozenFormulas: detectFrozen(prod, base),
		Unfilled:       detectUnfilled(prod, base),
		Diagnostics:    buildDiagnostics(prod, prodErrs),
	}
	res.OK = countMustFix(errorCells) == 0 && len(res.FrozenFormulas) == 0

	return res, nil
}

// countMustFix returns the number of error cells classified must_fix.
func countMustFix(cells []ErrorCell) int {
	n := 0
	for _, c := range cells {
		if c.Class == ClassMustFix {
			n++
		}
	}
	return n
}

// mergeErrors concatenates formula-error lists and de-duplicates by address
// (Sheet!Cell), keeping the first occurrence. The calc-engine scan wins over a
// latent flag for the same cell so a cell excelize already errors on is not
// double-reported.
func mergeErrors(lists ...[]FormulaError) []FormulaError {
	var out []FormulaError
	seen := map[string]bool{}
	for _, list := range lists {
		for _, e := range list {
			k := errKey(e)
			if seen[k] {
				continue
			}
			seen[k] = true
			out = append(out, e)
		}
	}
	return out
}

// scanPlaceholderErrors finds template required-input markers: formula cells
// whose CACHED value (written by the LibreOffice recalc that runs before this
// scan — see runValidate) is one of the literal placeholder strings ("#WERT!"
// / "#VALUE!") produced by an IF false-branch (e.g. IF(B16>15, …, "#WERT!")).
//
// The template uses such a literal as a "fill-me" marker: when the gating
// input is unfilled the gate fails and the cell returns the marker verbatim.
// LibreOffice caches that literal as a plain string result — it is NOT a
// recognised Excel error value (errorValues), so scanFormulaErrors misses it.
// Yet it is a real, fixable defect: filling the gating input flips the gate and
// the marker disappears, so we surface it like any other error cell and let the
// classifier name the gating input as the fix.
//
// A placeholder literal and a genuine calc error can render to the SAME
// string ("#VALUE!" is both a recognised Excel error AND a marker literal a
// template author might choose), so the marker-string match alone is
// ambiguous. The stored cell type disambiguates them: OOXML persists a
// genuine error's cache with t="e" (excelize.CellTypeError); a formula whose
// cached result is a string — including a placeholder literal from an IF
// false-branch — is persisted with t="str" (excelize.CellTypeFormula). Only
// the latter is eligible to be a placeholder; a t="e" cell is always a real
// calc error and is left to scanFormulaErrors.
func scanPlaceholderErrors(f *excelize.File) []FormulaError {
	var errs []FormulaError
	for _, sheet := range f.GetSheetList() {
		endCol, endRow, err := dataBounds(f, sheet)
		if err != nil || endCol == 0 || endRow == 0 {
			continue
		}
		for r := 1; r <= endRow; r++ {
			for c := 1; c <= endCol; c++ {
				cellRef, err := excelize.CoordinatesToCellName(c, r)
				if err != nil {
					continue
				}
				formula, _ := f.GetCellFormula(sheet, cellRef)
				if formula == "" {
					continue
				}
				cellType, err := f.GetCellType(sheet, cellRef)
				if err != nil {
					continue
				}
				if cellType == excelize.CellTypeError {
					continue // genuine calc error (t="e"), never a placeholder
				}
				value, err := f.GetCellValue(sheet, cellRef, excelize.Options{RawCellValue: true})
				if err != nil {
					continue
				}
				if isPlaceholderMarker(value) {
					errs = append(errs, FormulaError{
						Sheet:   sheet,
						Cell:    cellRef,
						Error:   value, // the literal marker, e.g. "#WERT!"
						Formula: formula,
					})
				}
			}
		}
	}
	return errs
}

// isPlaceholderMarker reports whether s is one of the literal template
// fill-me markers a formula's IF false-branch returns.
func isPlaceholderMarker(s string) bool {
	for _, m := range placeholderMarkers {
		if s == m {
			return true
		}
	}
	return false
}

// placeholderKeySet returns the set of address keys (Sheet!Cell) whose formula
// evaluates to a literal placeholder marker — i.e. the cells scanPlaceholderErrors
// would flag. Used to mark an ErrorCell as a template placeholder regardless of
// which scan first reported it.
func placeholderKeySet(f *excelize.File) map[string]bool {
	set := map[string]bool{}
	for _, e := range scanPlaceholderErrors(f) {
		set[errKey(e)] = true
	}
	return set
}

// detectFrozen returns baseline formula cells that are now plain literals in
// produced (a live formula was replaced with a number/text).
func detectFrozen(prod, base *excelize.File) []FrozenCell {
	out := []FrozenCell{}
	for _, sheet := range base.GetSheetList() {
		endCol, endRow, err := dataBounds(base, sheet)
		if err != nil || endCol == 0 || endRow == 0 {
			continue
		}
		for r := 1; r <= endRow; r++ {
			for c := 1; c <= endCol; c++ {
				cellRef, err := excelize.CoordinatesToCellName(c, r)
				if err != nil {
					continue
				}
				bf, _ := base.GetCellFormula(sheet, cellRef)
				if bf == "" {
					continue // baseline cell is not a formula
				}
				pf, _ := prod.GetCellFormula(sheet, cellRef)
				if pf != "" {
					continue // still a formula in produced
				}
				pv, _ := prod.GetCellValue(sheet, cellRef, excelize.Options{RawCellValue: true})
				if pv == "" {
					continue // formula gone AND empty -> deletion, not a freeze
				}
				out = append(out, FrozenCell{Sheet: sheet, Cell: cellRef, BaselineFormula: bf})
			}
		}
	}
	return out
}

// unfilledSamplePerSheet caps the example addresses carried per sheet group
// so the summary stays small even when a sheet has thousands of unfilled
// slots.
const unfilledSamplePerSheet = 10

// detectUnfilled returns an aggregated, per-sheet summary of template input
// slots still empty in produced, scanning PROTECTED sheets only.
//
// A target is a baseline cell that is unlocked + non-formula + empty, and not
// a merged-range member. It is flagged when the produced cell at the same
// address is also empty. On an unprotected sheet every cell is unlocked, so
// emptiness carries no author intent — those sheets are skipped entirely.
func detectUnfilled(prod, base *excelize.File) UnfilledSummary {
	sum := UnfilledSummary{BySheet: []UnfilledSheetGroup{}}
	for _, sheet := range base.GetSheetList() {
		// Unlocked-on-a-protected-sheet is a deliberate template-author
		// signal ("fill this in"). On an unprotected sheet every cell is
		// unlocked, so emptiness carries no signal — skip entirely.
		if !hasRawSheetProtection(base, sheet) {
			continue
		}
		group := UnfilledSheetGroup{Sheet: sheet, Sample: []string{}}
		merged := mergedMemberSet(base, sheet)
		endCol, endRow, err := dataBounds(base, sheet)
		if err != nil || endCol == 0 || endRow == 0 {
			continue
		}
		for r := 1; r <= endRow; r++ {
			for c := 1; c <= endCol; c++ {
				cellRef, err := excelize.CoordinatesToCellName(c, r)
				if err != nil {
					continue
				}
				if merged[cellRef] {
					continue // merged-range member -> not an individual slot
				}
				bf, _ := base.GetCellFormula(sheet, cellRef)
				if bf != "" {
					continue // formula -> not an input slot
				}
				bv, _ := base.GetCellValue(sheet, cellRef, excelize.Options{RawCellValue: true})
				if bv != "" {
					continue // non-empty in template -> not a target
				}
				if isCellLocked(base, sheet, cellRef) {
					continue // locked -> not an input slot
				}
				pv, _ := prod.GetCellValue(sheet, cellRef, excelize.Options{RawCellValue: true})
				if pv != "" {
					continue // filled in produced -> not unfilled
				}
				group.Count++
				if len(group.Sample) < unfilledSamplePerSheet {
					group.Sample = append(group.Sample, cellRef)
				}
			}
		}
		if group.Count > 0 {
			if group.Count > len(group.Sample) {
				sum.Truncated = true
			}
			sum.Total += group.Count
			sum.BySheet = append(sum.BySheet, group)
		}
	}
	return sum
}

// mergedMemberSet returns the set of every cell address that lies inside a
// merged range on the sheet.
func mergedMemberSet(f *excelize.File, sheet string) map[string]bool {
	set := map[string]bool{}
	cells, err := f.GetMergeCells(sheet)
	if err != nil {
		return set
	}
	for _, m := range cells {
		for _, ref := range expandRange(m.GetStartAxis(), m.GetEndAxis()) {
			set[ref] = true
		}
	}
	return set
}

// expandRange enumerates every cell address in the rectangle bounded by start
// and end (inclusive). Returns nil if either bound is unparsable.
func expandRange(start, end string) []string {
	sc, sr, err := excelize.CellNameToCoordinates(start)
	if err != nil {
		return nil
	}
	ec, er, err := excelize.CellNameToCoordinates(end)
	if err != nil {
		return nil
	}
	if sc > ec {
		sc, ec = ec, sc
	}
	if sr > er {
		sr, er = er, sr
	}
	var out []string
	for r := sr; r <= er; r++ {
		for c := sc; c <= ec; c++ {
			ref, err := excelize.CoordinatesToCellName(c, r)
			if err != nil {
				continue
			}
			out = append(out, ref)
		}
	}
	return out
}

// buildDiagnostics explains each new formula error by walking the cells its
// formula references (one hop): for each referenced cell that itself carries a
// formula, report that formula and the live values of the cells IT references.
// This is the generic gate-cascade diagnosis — it reads the workbook's own
// formulas and never hardcodes a column.
func buildDiagnostics(f *excelize.File, newErrs []FormulaError) []Diagnostic {
	out := []Diagnostic{}
	seen := map[string]bool{}
	for _, e := range newErrs {
		for _, ref := range cellRefRe.FindAllString(e.Formula, -1) {
			sheet, cell := splitRef(ref, e.Sheet)
			gf, _ := f.GetCellFormula(sheet, cell)
			if gf == "" {
				continue // referenced cell is not itself a formula
			}
			diagKey := sheet + "!" + cell
			if seen[diagKey] {
				continue
			}
			seen[diagKey] = true

			ops := map[string]string{}
			for _, op := range cellRefRe.FindAllString(gf, -1) {
				os, oc := splitRef(op, sheet)
				v, _ := f.GetCellValue(os, oc, excelize.Options{RawCellValue: true})
				ops[op] = v
			}
			out = append(out, Diagnostic{Cell: diagKey, Formula: gf, Operands: ops})
		}
	}
	return out
}

// maxTraceDepth caps the multi-hop operand walk so a cyclic or pathological
// formula graph can never loop forever. Real cross-sheet total chains are a
// handful of hops deep; this is generously above that.
const maxTraceDepth = 64

// leafRef is one non-formula cell reached by the operand trace.
type leafRef struct {
	sheet string
	cell  string
	gate  bool // appeared inside an IF condition on the path (a gating input)
	// locked reports whether the agent cannot write the cell — i.e. the sheet
	// is protected AND the cell carries the locked flag (isWriteBlocked). On an
	// unprotected sheet every cell is freely editable, so locked is false even
	// for a cell with Excel's default locked style. A template-owned cell on a
	// protected sheet is locked.
	locked bool
	empty  bool
}

func (l leafRef) addr() string { return l.sheet + "!" + l.cell }

// classifyError classifies one enumerated error cell by tracing its operand
// chain to leaf (non-formula) inputs.
//
//   - The chain reaches an unlocked, agent-fillable leaf -> must_fix: the agent
//     can change an input to resolve the error. FixInput names the input cell
//     to set, preferring an empty leaf, then a gating input from an IF
//     condition, then any unlocked leaf.
//   - The chain reaches ONLY locked / template-owned leaves -> external: nothing
//     the agent can change, so it is informational and does not block ok.
//
// isPlaceholder marks the cell as a template required-input marker
// (IF(…, "#WERT!"/"#VALUE!")) rather than a calc error.
func classifyError(f *excelize.File, e FormulaError, isPlaceholder bool) ErrorCell {
	out := ErrorCell{
		Sheet:       e.Sheet,
		Cell:        e.Cell,
		Error:       e.Error,
		Formula:     e.Formula,
		Placeholder: isPlaceholder,
	}

	leaves := traceLeaves(f, e.Sheet, e.Formula)

	// Pick the fix input among the unlocked leaves: prefer an empty one (the
	// agent must fill it), then a gating input (the IF condition the agent must
	// satisfy), then any unlocked leaf.
	var fix *leafRef
	for i := range leaves {
		l := &leaves[i]
		if l.locked {
			continue
		}
		switch {
		case fix == nil:
			fix = l
		case l.empty && !fix.empty:
			fix = l
		case l.gate && !fix.gate && fix.empty == l.empty:
			fix = l
		}
	}

	if fix != nil {
		out.Class = ClassMustFix
		out.FixInput = fix.addr()
	} else {
		// No unlocked leaf -> only locked/template cells feed this error.
		out.Class = ClassExternal
	}
	return out
}

// traceLeaves walks the operand graph of a formula, following referenced cells
// that are themselves formulas (multi-hop, depth-capped), and returns the
// non-formula leaf cells it reaches. A leaf referenced inside an IF condition is
// marked as a gating input. Each leaf is recorded once; lock/empty state is
// resolved from the workbook.
func traceLeaves(f *excelize.File, sheet, formula string) []leafRef {
	var leaves []leafRef
	seenLeaf := map[string]bool{}
	visited := map[string]bool{}

	var walkFormula func(curSheet, curFormula string, fromGate bool, depth int)
	walkFormula = func(curSheet, curFormula string, fromGate bool, depth int) {
		if depth > maxTraceDepth {
			return
		}
		gateRefs := ifConditionRefs(curFormula)
		for _, ref := range cellRefRe.FindAllString(curFormula, -1) {
			rSheet, rCell := splitRef(ref, curSheet)
			key := rSheet + "!" + rCell
			isGate := fromGate || gateRefs[normRef(ref)]

			gf, _ := f.GetCellFormula(rSheet, rCell)
			if gf != "" {
				// A formula cell -> recurse (multi-hop). Guard against cycles.
				if visited[key] {
					continue
				}
				visited[key] = true
				walkFormula(rSheet, gf, isGate, depth+1)
				continue
			}

			// A non-formula leaf input.
			if seenLeaf[key] {
				// Upgrade a previously-seen leaf to gate if reached via a gate
				// path on a later edge.
				if isGate {
					for i := range leaves {
						if leaves[i].addr() == key {
							leaves[i].gate = true
						}
					}
				}
				continue
			}
			seenLeaf[key] = true

			raw, _ := f.GetCellValue(rSheet, rCell, excelize.Options{RawCellValue: true})
			leaves = append(leaves, leafRef{
				sheet: rSheet,
				cell:  rCell,
				gate:  isGate,
				// Template-owned == not writable by the agent: locked AND the sheet
				// is protected (isWriteBlocked). isCellLocked alone would mark
				// every default-styled cell on an unprotected sheet as locked.
				locked: isWriteBlocked(f, rSheet, rCell),
				empty:  raw == "",
			})
		}
	}
	walkFormula(sheet, formula, false, 0)
	return leaves
}

// ifConditionRefs returns the set of normalized cell references that appear in
// the CONDITION (first argument) of an IF call within the formula — the gating
// inputs. References in the true/false branches are excluded. The set is keyed
// by the normalized (sheet-less or sheet-qualified, $-stripped) reference.
func ifConditionRefs(formula string) map[string]bool {
	refs := map[string]bool{}
	lower := strings.ToLower(formula)
	for _, cond := range extractIfConditions(formula, lower) {
		for _, ref := range cellRefRe.FindAllString(cond, -1) {
			refs[normRef(ref)] = true
		}
	}
	return refs
}

// extractIfConditions returns the text of every IF(...) condition (the segment
// from after "IF(" up to the top-level comma that ends the first argument) in
// the formula. It scans for "if(" tokens and balances parentheses to find the
// condition's terminating comma, so nested calls inside the condition are kept
// intact.
func extractIfConditions(formula, lower string) []string {
	var conds []string
	for i := 0; i+3 <= len(lower); i++ {
		if lower[i:i+3] != "if(" {
			continue
		}
		// Ensure "if" is a standalone token, not the tail of another name.
		if i > 0 && isFormulaNameByte(lower[i-1]) {
			continue
		}
		start := i + 3
		depth := 0
		end := -1
		for j := start; j < len(formula); j++ {
			switch formula[j] {
			case '(':
				depth++
			case ')':
				depth--
			case ',':
				if depth == 0 {
					end = j
				}
			}
			if end != -1 {
				break
			}
		}
		if end == -1 {
			end = len(formula) // malformed; take the rest
		}
		conds = append(conds, formula[start:end])
	}
	return conds
}

// isFormulaNameByte reports whether b can be part of a function/defined-name
// identifier (so "if" inside e.g. "MIDIF(" is not mistaken for the IF function).
func isFormulaNameByte(b byte) bool {
	return b == '_' || b == '.' ||
		(b >= 'a' && b <= 'z') || (b >= 'A' && b <= 'Z') || (b >= '0' && b <= '9')
}

// normRef normalizes a cell reference for set membership: strips $ anchors and
// surrounding single quotes around a sheet name.
func normRef(ref string) string {
	if i := strings.LastIndex(ref, "!"); i != -1 {
		return strings.Trim(ref[:i], "'") + "!" + strings.ReplaceAll(ref[i+1:], "$", "")
	}
	return strings.ReplaceAll(ref, "$", "")
}

// splitRef splits an A1-style reference into (sheet, cell), defaulting to
// defaultSheet when the reference carries no sheet qualifier. $ anchors are
// stripped so the address is usable with the excelize cell APIs.
func splitRef(ref, defaultSheet string) (sheet, cell string) {
	sheet = defaultSheet
	cell = ref
	if i := strings.LastIndex(ref, "!"); i != -1 {
		sheet = strings.Trim(ref[:i], "'")
		cell = ref[i+1:]
	}
	cell = strings.ReplaceAll(cell, "$", "")
	return sheet, cell
}

// NewValidateCmd builds the `validate` cobra command.
func NewValidateCmd() *cobra.Command {
	var baseline string
	var jsonOut bool

	cmd := &cobra.Command{
		Use:   "validate <file.xlsx>",
		Short: "Validate a produced workbook against the file the task received",
		Long: `Validate a produced workbook against the file the task received
(--baseline), using only the file's own structure and formulas (no
domain knowledge). Four checks run:

  1. errors           -- EVERY error cell in the recalced produced workbook
                         (standard Excel errors AND template "#WERT!"/
                         "#VALUE!" placeholder markers), each
                         classified by tracing its operand chain to leaf inputs:
                           must_fix -- the chain reaches an unlocked,
                                       agent-fillable input the agent can resolve;
                                       fix_input names the cell to set (and the
                                       cell is flagged a placeholder when it is
                                       an IF(...,"#WERT!") template marker).
                           external -- the chain reaches only locked / template-owned
                                       cells; informational, never blocks ok.
  2. frozen formulas  -- baseline formula cells now replaced by literals
  3. unfilled targets -- aggregated per-sheet summary of still-empty input
                         slots on PROTECTED sheets (informational; never
                         blocks ok -- mandatory inputs surface as must_fix
                         placeholder chains instead)
  4. operand diagnostics -- for each failing formula, the referenced cells'
                         formulas and their operands' live values

ok is true iff there are NO must_fix errors and no frozen formulas. Unlike a
diff-vs-baseline check, validate reports a pre-existing fixable error (e.g. a
template placeholder the agent must resolve) rather than hiding it. external
errors, unfilled targets, and diagnostics never affect ok.

validate first recalculates the produced workbook in place via headless
LibreOffice (` + "`soffice`" + ` or ` + "`libreoffice`" + ` on PATH), then checks the computed values.
The baseline is read-only and never modified.

Exits non-zero when validation fails (any must_fix error or frozen formula) so
a calling shell sees the failure.`,
		Example: `  # Validate the edited file against the received copy
  xlsx-tool validate produced.xlsx --baseline received.xlsx

  # Machine-readable output (errors[].class, errors[].fix_input)
  xlsx-tool validate produced.xlsx --baseline received.xlsx --json`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			res, err := runValidate(args[0], baseline)
			if err != nil {
				return err
			}
			if jsonOut {
				if err := format.WriteIndentedJSON(cmd.OutOrStdout(), res); err != nil {
					return err
				}
			} else {
				writeValidatePlain(cmd.OutOrStdout(), res)
			}
			if !res.OK {
				return fmt.Errorf("validation failed: %d must_fix error(s), %d frozen formula(s)",
					countMustFix(res.Errors), len(res.FrozenFormulas))
			}
			return nil
		},
	}

	cmd.Flags().StringVar(&baseline, "baseline", "", "the file as the task received it (required)")
	_ = cmd.MarkFlagRequired("baseline")
	cmd.Flags().BoolVar(&jsonOut, "json", false, "machine-readable JSON output")
	return cmd
}

// writeValidatePlain prints a short human-readable summary of the result.
func writeValidatePlain(w io.Writer, res ValidateResult) {
	if res.OK {
		fmt.Fprintln(w, "OK: no must_fix errors or frozen formulas.")
	} else {
		fmt.Fprintln(w, "FAILED")
	}
	for _, e := range res.Errors {
		marker := ""
		if e.Placeholder {
			marker = " [placeholder]"
		}
		switch e.Class {
		case ClassMustFix:
			fmt.Fprintf(w, "  must_fix: %s!%s %s (=%s)%s -> fill %s\n",
				e.Sheet, e.Cell, e.Error, e.Formula, marker, e.FixInput)
		default:
			fmt.Fprintf(w, "  external: %s!%s %s (=%s)%s\n",
				e.Sheet, e.Cell, e.Error, e.Formula, marker)
		}
	}
	for _, fz := range res.FrozenFormulas {
		fmt.Fprintf(w, "  frozen: %s!%s (was =%s)\n", fz.Sheet, fz.Cell, fz.BaselineFormula)
	}
	for _, g := range res.Unfilled.BySheet {
		fmt.Fprintf(w, "  unfilled (info): %s: %d slots (e.g. %s)\n",
			g.Sheet, g.Count, strings.Join(g.Sample, ", "))
	}
	for _, d := range res.Diagnostics {
		fmt.Fprintf(w, "  diagnostic: %s = %s ; operands %v\n", d.Cell, d.Formula, d.Operands)
	}
}
