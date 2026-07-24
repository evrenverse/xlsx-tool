package commands

import (
	"fmt"
	"io"
	"os"

	"github.com/spf13/cobra"
)

func NewSheetsCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "sheets <file>",
		Short: "List sheet names in an Excel file",
		Long: `Print each sheet name in the workbook, one per line.

Useful for scripting — pipe the output to other commands or use it
to discover which sheet names to pass to 'read' or 'write'.`,
		Example: `  # List all sheets
  xlsx-tool sheets report.xlsx

  # Use in a script to read every sheet
  for sheet in $(xlsx-tool sheets report.xlsx); do
    xlsx-tool read report.xlsx "$sheet" --json > "${sheet}.json"
  done`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			return runSheets(args[0], os.Stdout)
		},
	}
}

func runSheets(filePath string, w io.Writer) error {
	f, err := openWorkbook(filePath, false)
	if err != nil {
		return fmt.Errorf("open file: %w", err)
	}
	defer f.Close()

	for _, name := range f.GetSheetList() {
		fmt.Fprintln(w, name)
	}
	return nil
}
