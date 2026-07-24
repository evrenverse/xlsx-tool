package format

import (
	"bytes"
	"strings"
	"testing"
)

func TestFormatCellPlain(t *testing.T) {
	tests := []struct {
		name string
		cell PlainCell
		want string
	}{
		{name: "empty", cell: PlainCell{}, want: ""},
		{name: "number", cell: PlainCell{Value: "42"}, want: "42"},
		{name: "text", cell: PlainCell{Value: "alpha"}, want: `"alpha"`},
		{
			name: "display",
			cell: PlainCell{Value: "0.25", Display: "25.00%"},
			want: `0.25 [display "25.00%"]`,
		},
		{
			name: "formula",
			cell: PlainCell{Formula: "SUM(A1:A2)", Value: "3", Fill: "#FFFF00"},
			want: "=SUM(A1:A2) → 3 [fill=#FFFF00]",
		},
		{
			name: "formula error",
			cell: PlainCell{Formula: "1/0", Value: "#DIV/0!"},
			want: `=1/0 → "#DIV/0!" [ERROR]`,
		},
		{name: "fill only", cell: PlainCell{Fill: "#00FF00"}, want: "[fill=#00FF00]"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			if got := FormatCellPlain(test.cell); got != test.want {
				t.Fatalf("got %q, want %q", got, test.want)
			}
		})
	}
}

func TestWritePlainSheetSkipsEmptyRows(t *testing.T) {
	var output bytes.Buffer
	rows := [][]PlainCell{
		{{}, {}},
		{{Value: "alpha"}, {Formula: "1+1", Value: "2"}},
	}
	if err := WritePlainSheet(&output, "Data", "A1:B2", rows, 1, 1, []string{"A", "B"}); err != nil {
		t.Fatal(err)
	}
	got := output.String()
	if strings.Contains(got, "[1]") {
		t.Fatalf("empty row was written: %s", got)
	}
	if !strings.Contains(got, `[2] A: "alpha" | B: =1+1 → 2`) {
		t.Fatalf("missing populated row: %s", got)
	}
}

func TestErrorValues(t *testing.T) {
	if !IsErrorValue("#N/A") {
		t.Fatal("expected #N/A to be an error")
	}
	if IsErrorValue("N/A") {
		t.Fatal("did not expect N/A to be an error")
	}
}
