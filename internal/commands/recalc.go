package commands

import (
	"fmt"
	"os"

	"github.com/spf13/cobra"
)

// NewRecalcCmd exposes the LibreOffice recalculation step on its own.
//
// The exit code reports only whether recalculation itself worked. A workbook
// whose formulas evaluate to errors, or whose required inputs are still empty,
// recalculates successfully and exits 0 — judging cell contents is 'validate'.
func NewRecalcCmd() *cobra.Command {
	return &cobra.Command{
		Use:   "recalc <file>",
		Short: "Recalculate formulas in place with LibreOffice",
		Long: `Recalculate every formula in the workbook and write the refreshed
cached values back to the same file.

LibreOffice does the calculation in a private, throwaway profile; the result
replaces the original atomically, preserving its permissions. Readers that
only consume cached values — including this tool's own 'read' — see the fresh
numbers afterwards.

The exit code covers recalculation only: a missing LibreOffice, a timeout, or
a copy or write failure. Formula errors such as #VALUE! and unfilled required
cells are content, not failure, and leave the exit code at 0. Use 'validate'
to judge them.`,
		Example: `  # Refresh cached formula results in place
  xlsx-tool recalc report.xlsx

  # Recalculate, then read the updated values
  xlsx-tool recalc report.xlsx && xlsx-tool read report.xlsx Sheet1 --cell B5`,
		Args: cobra.ExactArgs(1),
		RunE: func(cmd *cobra.Command, args []string) error {
			if _, err := os.Stat(args[0]); err != nil {
				return fmt.Errorf("open file: %w", err)
			}
			return runExternalRecalc(args[0])
		},
	}
}
