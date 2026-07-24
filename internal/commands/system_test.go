package commands

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestCapabilitiesJSON(t *testing.T) {
	cmd := NewCapabilitiesCmd("1.2.3")
	var output bytes.Buffer
	cmd.SetOut(&output)
	cmd.SetArgs([]string{"--json"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var payload map[string]any
	if err := json.Unmarshal(output.Bytes(), &payload); err != nil {
		t.Fatal(err)
	}
	if payload["tool"] != "xlsx-tool" || payload["contract_version"] != "1" {
		t.Fatalf("unexpected payload: %v", payload)
	}
}

func TestDoctorJSON(t *testing.T) {
	cmd := NewDoctorCmd("dev")
	var output bytes.Buffer
	cmd.SetOut(&output)
	cmd.SetArgs([]string{"--json"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	var payload doctorPayload
	if err := json.Unmarshal(output.Bytes(), &payload); err != nil {
		t.Fatal(err)
	}
	if payload.Tool != "xlsx-tool" || len(payload.Checks) < 3 {
		t.Fatalf("unexpected payload: %+v", payload)
	}
}

func TestSchema(t *testing.T) {
	cmd := NewSchemaCmd()
	var output bytes.Buffer
	cmd.SetOut(&output)
	cmd.SetArgs([]string{"batch"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(output.String(), `"xlsx-tool batch input"`) {
		t.Fatalf("unexpected schema: %s", output.String())
	}
}

func TestVersionJSON(t *testing.T) {
	cmd := NewVersionCmd("1.2.3")
	var output bytes.Buffer
	cmd.SetOut(&output)
	cmd.SetArgs([]string{"--json"})
	if err := cmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(output.String(), `"version": "1.2.3"`) {
		t.Fatalf("unexpected version output: %s", output.String())
	}
}

func TestHumanReadableCapabilitiesVersionAndDoctor(t *testing.T) {
	capabilitiesCmd := NewCapabilitiesCmd("1.2.3")
	var capabilitiesOutput bytes.Buffer
	capabilitiesCmd.SetOut(&capabilitiesOutput)
	if err := capabilitiesCmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(capabilitiesOutput.String(), "agent contract v1") {
		t.Fatalf("unexpected capabilities output: %s", capabilitiesOutput.String())
	}

	versionCmd := NewVersionCmd("1.2.3")
	var versionOutput bytes.Buffer
	versionCmd.SetOut(&versionOutput)
	if err := versionCmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if strings.TrimSpace(versionOutput.String()) != "xlsx-tool 1.2.3" {
		t.Fatalf("unexpected version output: %s", versionOutput.String())
	}

	doctorCmd := NewDoctorCmd("dev")
	var doctorOutput bytes.Buffer
	doctorCmd.SetOut(&doctorOutput)
	if err := doctorCmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(doctorOutput.String(), "go: ok") {
		t.Fatalf("unexpected doctor output: %s", doctorOutput.String())
	}
}

func TestSchemaListAndUnknownSchema(t *testing.T) {
	listCmd := NewSchemaCmd()
	var listOutput bytes.Buffer
	listCmd.SetOut(&listOutput)
	if err := listCmd.Execute(); err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(listOutput.String(), `"version"`) {
		t.Fatalf("unexpected schema list: %s", listOutput.String())
	}

	unknownCmd := NewSchemaCmd()
	unknownCmd.SetArgs([]string{"ghost"})
	if err := unknownCmd.Execute(); err == nil {
		t.Fatal("expected unknown schema to fail")
	}
}

func TestDoctorRejectsUnknownRequiredDependency(t *testing.T) {
	cmd := NewDoctorCmd("dev")
	cmd.SetArgs([]string{"--require", "ghost"})
	if err := cmd.Execute(); err == nil {
		t.Fatal("expected unknown dependency to fail")
	}
}
