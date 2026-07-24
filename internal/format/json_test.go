package format

import (
	"bytes"
	"encoding/json"
	"testing"
)

func TestJSONWriters(t *testing.T) {
	tests := []struct {
		name  string
		write func(*bytes.Buffer) error
	}{
		{
			name: "sheet",
			write: func(buffer *bytes.Buffer) error {
				return WriteJSON(buffer, &SheetJSON{
					Sheet: "Summary",
					Range: "A1:B1",
					Rows: []RowJSON{{
						Row:   1,
						Cells: map[string]any{"A": "Label", "B": 42},
					}},
				})
			},
		},
		{
			name: "cells",
			write: func(buffer *bytes.Buffer) error {
				return WriteCellsJSON(buffer, &CellsJSON{
					Sheet: "Summary",
					Cells: map[string]any{"A1": "Label"},
				})
			},
		},
		{
			name: "workbook",
			write: func(buffer *bytes.Buffer) error {
				return WriteWorkbookJSON(buffer, &WorkbookJSON{
					Sheets: []*SheetJSON{{Sheet: "Summary", Rows: []RowJSON{}}},
				})
			},
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			var output bytes.Buffer
			if err := test.write(&output); err != nil {
				t.Fatal(err)
			}
			var payload any
			if err := json.Unmarshal(output.Bytes(), &payload); err != nil {
				t.Fatalf("invalid JSON: %v\n%s", err, output.String())
			}
		})
	}
}

func TestWriteIndentedJSONPropagatesWriterFailure(t *testing.T) {
	if err := WriteIndentedJSON(failingWriter{}, map[string]string{"key": "value"}); err == nil {
		t.Fatal("expected writer error")
	}
}

type failingWriter struct{}

func (failingWriter) Write(_ []byte) (int, error) {
	return 0, errTestWriter
}

type testWriterError struct{}

func (testWriterError) Error() string { return "writer failed" }

var errTestWriter testWriterError
