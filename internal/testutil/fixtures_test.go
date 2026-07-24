package testutil

import "testing"

func TestCreateSimpleFixture(t *testing.T) {
	path := CreateSimpleFixture(t)
	if !FileExists(path) {
		t.Fatal("fixture file was not created")
	}

	val := ReadCellValue(t, path, "Sheet1", "A1")
	if val != "Name" {
		t.Errorf("expected 'Name', got %q", val)
	}

	val = ReadCellValue(t, path, "Sheet1", "B2")
	if val != "10" {
		t.Errorf("expected '10', got %q", val)
	}

	formula := ReadCellFormula(t, path, "Sheet1", "B4")
	if formula != "SUM(B2:B3)" {
		t.Errorf("expected 'SUM(B2:B3)', got %q", formula)
	}
}

func TestCreateFormattedFixture(t *testing.T) {
	path := CreateFormattedFixture(t)
	if !FileExists(path) {
		t.Fatal("fixture file was not created")
	}

	font := GetCellStyleFont(t, path, "Sheet1", "A1")
	if font == nil {
		t.Fatal("expected font style on A1")
	}
	if !font.Bold {
		t.Error("expected bold font on A1")
	}
}
