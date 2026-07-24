package commands

import (
	"embed"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"os/exec"
	"runtime"
	"sort"
	"strings"

	"github.com/spf13/cobra"
)

const agentContractVersion = "1"

//go:embed schemas/*.json
var contractSchemas embed.FS

type capabilityCommand struct {
	Name     string `json:"name"`
	Mutates  bool   `json:"mutates"`
	JSON     bool   `json:"json"`
	Optional string `json:"optional_dependency,omitempty"`
}

type capabilitiesPayload struct {
	SchemaVersion        string              `json:"schema_version"`
	Tool                 string              `json:"tool"`
	ToolVersion          string              `json:"tool_version"`
	ContractVersion      string              `json:"contract_version"`
	Addressing           string              `json:"addressing"`
	NetworkAccess        bool                `json:"runtime_network_access"`
	AtomicWrites         bool                `json:"atomic_writes"`
	StructuredStdout     bool                `json:"structured_stdout"`
	DiagnosticsStderr    bool                `json:"diagnostics_stderr"`
	Commands             []capabilityCommand `json:"commands"`
	OptionalDependencies []string            `json:"optional_dependencies"`
	Schemas              []string            `json:"schemas"`
	Limits               map[string]int64    `json:"limits"`
}

type doctorCheck struct {
	Name     string `json:"name"`
	Required bool   `json:"required"`
	OK       bool   `json:"ok"`
	Detail   string `json:"detail"`
}

type doctorPayload struct {
	SchemaVersion string        `json:"schema_version"`
	Tool          string        `json:"tool"`
	ToolVersion   string        `json:"tool_version"`
	Status        string        `json:"status"`
	Checks        []doctorCheck `json:"checks"`
}

func capabilities(version string) capabilitiesPayload {
	return capabilitiesPayload{
		SchemaVersion:     agentContractVersion,
		Tool:              "xlsx-tool",
		ToolVersion:       version,
		ContractVersion:   agentContractVersion,
		Addressing:        "sheet names and A1 cell references",
		NetworkAccess:     false,
		AtomicWrites:      true,
		StructuredStdout:  true,
		DiagnosticsStderr: true,
		Commands: []capabilityCommand{
			{Name: "sheets", JSON: true},
			{Name: "info", JSON: true},
			{Name: "find", JSON: true},
			{Name: "read", JSON: true},
			{Name: "write", Mutates: true, JSON: true},
			{Name: "batch", Mutates: true, JSON: true},
			{Name: "validate", Mutates: true, JSON: true, Optional: "libreoffice"},
			{Name: "capabilities", JSON: true},
			{Name: "doctor", JSON: true},
			{Name: "schema", JSON: true},
			{Name: "version", JSON: true},
		},
		OptionalDependencies: []string{"libreoffice"},
		Schemas:              []string{"batch", "capabilities", "doctor", "version"},
		Limits: map[string]int64{
			"batch_json_bytes":        maxBatchJSONBytes,
			"batch_changes":           maxBatchChanges,
			"expanded_workbook_bytes": maxWorkbookExpandedBytes,
			"in_memory_xml_bytes":     maxWorkbookXMLBytes,
		},
	}
}

func writeSystemJSON(w io.Writer, value any) error {
	encoder := json.NewEncoder(w)
	encoder.SetIndent("", "  ")
	return encoder.Encode(value)
}

func NewCapabilitiesCmd(version string) *cobra.Command {
	var outputJSON bool
	cmd := &cobra.Command{
		Use:   "capabilities",
		Short: "Describe the stable automation contract",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			payload := capabilities(version)
			if outputJSON {
				return writeSystemJSON(cmd.OutOrStdout(), payload)
			}
			fmt.Fprintf(cmd.OutOrStdout(), "xlsx-tool %s (agent contract v%s)\n", version, agentContractVersion)
			fmt.Fprintln(cmd.OutOrStdout(), "Addressing: sheet names and A1 cell references")
			fmt.Fprintln(cmd.OutOrStdout(), "Writes: validated, fsynced, and atomically published")
			fmt.Fprintln(cmd.OutOrStdout(), "Runtime network access: none")
			fmt.Fprintln(cmd.OutOrStdout(), "Optional dependency: LibreOffice (validate only)")
			return nil
		},
	}
	cmd.Flags().BoolVar(&outputJSON, "json", false, "Output the contract as JSON")
	return cmd
}

func NewVersionCmd(version string) *cobra.Command {
	var outputJSON bool
	cmd := &cobra.Command{
		Use:   "version",
		Short: "Show version and contract information",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			if outputJSON {
				return writeSystemJSON(cmd.OutOrStdout(), map[string]string{
					"schema_version":   agentContractVersion,
					"tool":             "xlsx-tool",
					"version":          version,
					"contract_version": agentContractVersion,
				})
			}
			fmt.Fprintf(cmd.OutOrStdout(), "xlsx-tool %s\n", version)
			return nil
		},
	}
	cmd.Flags().BoolVar(&outputJSON, "json", false, "Output version information as JSON")
	return cmd
}

func NewDoctorCmd(version string) *cobra.Command {
	var outputJSON bool
	var required []string
	cmd := &cobra.Command{
		Use:   "doctor",
		Short: "Check runtime requirements without opening a workbook",
		Args:  cobra.NoArgs,
		RunE: func(cmd *cobra.Command, _ []string) error {
			requiredSet := make(map[string]bool, len(required))
			for _, name := range required {
				requiredSet[strings.ToLower(strings.TrimSpace(name))] = true
			}
			for name := range requiredSet {
				if name != "libreoffice" {
					return fmt.Errorf("unknown optional dependency %q", name)
				}
			}

			checks := []doctorCheck{{
				Name:     "go",
				Required: true,
				OK:       true,
				Detail:   runtime.Version(),
			}}
			tempCheck := doctorCheck{Name: "temporary-directory", Required: true}
			file, err := os.CreateTemp("", "xlsx-tool-doctor-*")
			if err == nil {
				tempCheck.OK = true
				tempCheck.Detail = "writable"
				name := file.Name()
				_ = file.Close()
				_ = os.Remove(name)
			} else {
				tempCheck.Detail = err.Error()
			}
			checks = append(checks, tempCheck)

			libreOffice := doctorCheck{
				Name:     "libreoffice",
				Required: requiredSet["libreoffice"],
			}
			for _, candidate := range []string{"soffice", "libreoffice"} {
				if path, lookupErr := exec.LookPath(candidate); lookupErr == nil {
					libreOffice.OK = true
					libreOffice.Detail = path
					break
				}
			}
			if !libreOffice.OK {
				libreOffice.Detail = "not found; required only for validate"
			}
			checks = append(checks, libreOffice)

			status := "ok"
			failedRequired := false
			for _, check := range checks {
				if check.Required && !check.OK {
					status = "error"
					failedRequired = true
				} else if !check.OK && status == "ok" {
					status = "degraded"
				}
			}
			payload := doctorPayload{
				SchemaVersion: agentContractVersion,
				Tool:          "xlsx-tool",
				ToolVersion:   version,
				Status:        status,
				Checks:        checks,
			}
			if outputJSON {
				if err := writeSystemJSON(cmd.OutOrStdout(), payload); err != nil {
					return err
				}
			} else {
				for _, check := range checks {
					state := "ok"
					if !check.OK {
						state = "missing"
					}
					fmt.Fprintf(cmd.OutOrStdout(), "%s: %s (%s)\n", check.Name, state, check.Detail)
				}
			}
			if failedRequired {
				return fmt.Errorf("one or more required runtime checks failed")
			}
			return nil
		},
	}
	cmd.Flags().BoolVar(&outputJSON, "json", false, "Output checks as JSON")
	cmd.Flags().StringSliceVar(&required, "require", nil, "Require optional dependencies (libreoffice)")
	return cmd
}

func schemaNames() []string {
	entries, _ := contractSchemas.ReadDir("schemas")
	names := make([]string, 0, len(entries))
	for _, entry := range entries {
		names = append(names, strings.TrimSuffix(entry.Name(), ".schema.json"))
	}
	sort.Strings(names)
	return names
}

func NewSchemaCmd() *cobra.Command {
	cmd := &cobra.Command{
		Use:   "schema [name]",
		Short: "Print a bundled JSON Schema",
		Args:  cobra.MaximumNArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if len(args) == 0 {
				return writeSystemJSON(cmd.OutOrStdout(), map[string]any{
					"schema_version": agentContractVersion,
					"schemas":        schemaNames(),
				})
			}
			name := strings.ToLower(args[0])
			data, err := contractSchemas.ReadFile("schemas/" + name + ".schema.json")
			if err != nil {
				return fmt.Errorf("unknown schema %q (available: %s)", name, strings.Join(schemaNames(), ", "))
			}
			_, err = cmd.OutOrStdout().Write(append(data, '\n'))
			return err
		},
	}
	return cmd
}
