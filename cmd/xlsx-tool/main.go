package main

import (
	"errors"
	"fmt"
	"os"
	"runtime/debug"

	"github.com/evrenverse/xlsx-tool/internal/commands"
	"github.com/spf13/cobra"
)

var version = "dev" // Set by GoReleaser at build time

// resolveVersion reports the version to expose through --version, the agent
// contract, and doctor. A GoReleaser build injects it with -ldflags, but
// `go install github.com/evrenverse/xlsx-tool/cmd/xlsx-tool@latest` links
// without them, so fall back to the module version the Go toolchain records
// in the binary. A local `go build` records "(devel)" and keeps "dev".
func resolveVersion() string {
	if version != "dev" {
		return version
	}
	if info, ok := debug.ReadBuildInfo(); ok {
		if v := info.Main.Version; v != "" && v != "(devel)" {
			return v
		}
	}
	return version
}

var rootCmd = &cobra.Command{
	Use:   "xlsx-tool",
	Short: "Inspect and edit XLSX workbooks from scripts and AI agents",
	Long: `xlsx-tool provides focused, machine-readable operations for Excel
(.xlsx) workbooks while preserving workbook features supported by Excelize.

Writes are crash-safe: changes are written to a temporary file and atomically
replace the destination only after a successful save.

Typical AI agent workflow:
  1. xlsx-tool info   file.xlsx                       -- discover sheets
  2. xlsx-tool find   file.xlsx "Total" --json        -- locate label cells
  3. xlsx-tool read   file.xlsx Sheet1 --cells B5,C5  -- read target cells
  4. xlsx-tool batch  file.xlsx changes.json --json   -- apply changes
  5. xlsx-tool read   file.xlsx Sheet1 --cells B5,C5  -- verify result`,
	// Errors are reported as one line on stderr without a usage dump so
	// callers can parse stdout independently.
	SilenceUsage:  true,
	SilenceErrors: true,
}

func init() {
	resolved := resolveVersion()
	rootCmd.Version = resolved
	rootCmd.AddCommand(commands.NewSheetsCmd())
	rootCmd.AddCommand(commands.NewInfoCmd())
	rootCmd.AddCommand(commands.NewReadCmd())
	rootCmd.AddCommand(commands.NewFindCmd())
	rootCmd.AddCommand(commands.NewWriteCmd())
	rootCmd.AddCommand(commands.NewBatchCmd())
	rootCmd.AddCommand(commands.NewRecalcCmd())
	rootCmd.AddCommand(commands.NewValidateCmd())
	rootCmd.AddCommand(commands.NewCapabilitiesCmd(resolved))
	rootCmd.AddCommand(commands.NewDoctorCmd(resolved))
	rootCmd.AddCommand(commands.NewSchemaCmd())
	rootCmd.AddCommand(commands.NewVersionCmd(resolved))
}

func main() {
	if err := rootCmd.Execute(); err != nil {
		// Zero-match find already reported its result on stdout — exit
		// non-zero without an extra error line.
		if !errors.Is(err, commands.ErrNoMatches) {
			fmt.Fprintln(os.Stderr, "Error:", err)
		}
		os.Exit(1)
	}
}
