package commands

import (
	"strings"
	"testing"
)

func FuzzNormalizeRangeNeverPanics(f *testing.F) {
	for _, seed := range []string{"A1", "B2:A1", "$C$3:$A$1", "", "A0", "A1:B2:C3", "Sheet1!A1"} {
		f.Add(seed)
	}
	f.Fuzz(func(t *testing.T, input string) {
		result, err := normalizeRange(input)
		if err == nil && strings.Contains(input, ":") && !strings.Contains(result, ":") {
			t.Fatalf("normalized range lost separator: input=%q result=%q", input, result)
		}
	})
}

func FuzzFormulaReferenceExtractionNeverPanics(f *testing.F) {
	for _, seed := range []string{
		"IF(A1>0,B1,C1)",
		"SUM(IF(A1>0, B1, C1))",
		`IF('Input Data'!$A$1="",0,'Input Data'!B2)`,
		"",
		"IF(",
	} {
		f.Add(seed)
	}
	f.Fuzz(func(t *testing.T, formula string) {
		lower := strings.ToLower(formula)
		conditions := extractIfConditions(formula, lower)
		for _, condition := range conditions {
			if len(condition) > len(formula) {
				t.Fatalf("condition longer than source formula")
			}
		}
		_ = ifConditionRefs(formula)
	})
}
