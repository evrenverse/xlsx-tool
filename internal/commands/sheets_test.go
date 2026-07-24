package commands

import (
	"bytes"
	"strings"
	"testing"

	"github.com/evrenverse/xlsx-tool/internal/testutil"
)

func TestSheetsCommand(t *testing.T) {
	path := testutil.CreateSimpleFixture(t)

	var stdout bytes.Buffer
	err := runSheets(path, &stdout)
	if err != nil {
		t.Fatalf("unexpected error: %v", err)
	}

	output := stdout.String()
	if !strings.Contains(output, "Sheet1") {
		t.Errorf("expected 'Sheet1' in output, got:\n%s", output)
	}
	if !strings.Contains(output, "Data") {
		t.Errorf("expected 'Data' in output, got:\n%s", output)
	}
}

func TestSheetsCommand_FileNotFound(t *testing.T) {
	var stdout bytes.Buffer
	err := runSheets("/nonexistent/file.xlsx", &stdout)
	if err == nil {
		t.Fatal("expected error for nonexistent file")
	}
}
