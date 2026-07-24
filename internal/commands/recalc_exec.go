package commands

import (
	"context"
	"fmt"
	"io"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"strings"
	"time"
)

const recalcTimeout = 120 * time.Second

// runExternalRecalc refreshes formula caches in xlsxPath with LibreOffice.
//
// XLSX_RECALC_BIN remains as a narrow testing and integration hook. Without
// it, the command invokes LibreOffice directly with an isolated, private user
// profile. No shared profile, generated native code, or LD_PRELOAD shim is
// used.
func runExternalRecalc(xlsxPath string) error {
	if override := os.Getenv("XLSX_RECALC_BIN"); override != "" {
		out, err := exec.Command(override, xlsxPath).CombinedOutput()
		if err != nil {
			return fmt.Errorf(
				"custom recalculation command failed: %w: %s",
				err,
				strings.TrimSpace(string(out)),
			)
		}
		return nil
	}
	return recalcWithLibreOffice(xlsxPath)
}

func recalcWithLibreOffice(xlsxPath string) error {
	bin, err := findLibreOffice()
	if err != nil {
		return err
	}

	absolute, err := filepath.Abs(xlsxPath)
	if err != nil {
		return fmt.Errorf("resolve workbook path: %w", err)
	}
	info, err := os.Stat(absolute)
	if err != nil {
		return fmt.Errorf("stat workbook: %w", err)
	}

	parent := filepath.Dir(absolute)
	workDir, err := os.MkdirTemp(parent, ".xlsx-tool-recalc-")
	if err != nil {
		return fmt.Errorf("create private recalculation directory: %w", err)
	}
	defer os.RemoveAll(workDir)

	inputDir := filepath.Join(workDir, "input")
	outputDir := filepath.Join(workDir, "output")
	profileDir := filepath.Join(workDir, "profile")
	for _, dir := range []string{inputDir, outputDir, profileDir} {
		if err := os.Mkdir(dir, 0o700); err != nil {
			return fmt.Errorf("create private recalculation directory: %w", err)
		}
	}

	inputPath := filepath.Join(inputDir, "workbook.xlsx")
	if err := copyFile(inputPath, absolute, 0o600); err != nil {
		return fmt.Errorf("copy workbook for recalculation: %w", err)
	}

	profileURL := (&url.URL{Scheme: "file", Path: filepath.ToSlash(profileDir)}).String()
	ctx, cancel := context.WithTimeout(context.Background(), recalcTimeout)
	defer cancel()
	cmd := exec.CommandContext(
		ctx,
		bin,
		"-env:UserInstallation="+profileURL,
		"--headless",
		"--nologo",
		"--nodefault",
		"--nofirststartwizard",
		"--convert-to",
		"xlsx",
		"--outdir",
		outputDir,
		inputPath,
	)
	output, runErr := cmd.CombinedOutput()
	if ctx.Err() == context.DeadlineExceeded {
		return fmt.Errorf("LibreOffice recalculation timed out after %s", recalcTimeout)
	}
	if runErr != nil {
		return fmt.Errorf(
			"LibreOffice recalculation failed: %w: %s",
			runErr,
			strings.TrimSpace(string(output)),
		)
	}

	recalculated := filepath.Join(outputDir, "workbook.xlsx")
	if _, err := os.Stat(recalculated); err != nil {
		return fmt.Errorf("LibreOffice did not produce a recalculated workbook: %w", err)
	}

	replacement, err := os.CreateTemp(parent, ".xlsx-tool-recalc-*.tmp")
	if err != nil {
		return fmt.Errorf("create replacement file: %w", err)
	}
	replacementPath := replacement.Name()
	defer os.Remove(replacementPath)

	source, err := os.Open(recalculated)
	if err != nil {
		replacement.Close()
		return fmt.Errorf("open recalculated workbook: %w", err)
	}
	_, copyErr := io.Copy(replacement, source)
	closeSourceErr := source.Close()
	syncErr := replacement.Sync()
	closeReplacementErr := replacement.Close()
	if copyErr != nil {
		return fmt.Errorf("copy recalculated workbook: %w", copyErr)
	}
	if closeSourceErr != nil {
		return fmt.Errorf("close recalculated workbook: %w", closeSourceErr)
	}
	if syncErr != nil {
		return fmt.Errorf("sync recalculated workbook: %w", syncErr)
	}
	if closeReplacementErr != nil {
		return fmt.Errorf("close replacement workbook: %w", closeReplacementErr)
	}
	if err := os.Chmod(replacementPath, info.Mode().Perm()); err != nil {
		return fmt.Errorf("preserve workbook permissions: %w", err)
	}
	if err := os.Rename(replacementPath, absolute); err != nil {
		return fmt.Errorf("replace workbook after recalculation: %w", err)
	}
	if err := syncDirectory(parent); err != nil {
		return fmt.Errorf("sync workbook directory: %w", err)
	}
	return nil
}

func findLibreOffice() (string, error) {
	candidates := []string{"soffice", "libreoffice"}
	if runtime.GOOS == "windows" {
		candidates = append(candidates,
			filepath.Join(os.Getenv("ProgramFiles"), "LibreOffice", "program", "soffice.exe"),
		)
	}
	for _, candidate := range candidates {
		if candidate == "" {
			continue
		}
		if path, err := exec.LookPath(candidate); err == nil {
			return path, nil
		}
	}
	return "", fmt.Errorf("LibreOffice not found; install LibreOffice or set XLSX_RECALC_BIN")
}

func copyFile(destinationPath, sourcePath string, mode os.FileMode) error {
	source, err := os.Open(sourcePath)
	if err != nil {
		return err
	}
	defer source.Close()

	destination, err := os.OpenFile(destinationPath, os.O_CREATE|os.O_EXCL|os.O_WRONLY, mode)
	if err != nil {
		return err
	}
	if _, err := io.Copy(destination, source); err != nil {
		destination.Close()
		return err
	}
	if err := destination.Sync(); err != nil {
		destination.Close()
		return err
	}
	return destination.Close()
}
