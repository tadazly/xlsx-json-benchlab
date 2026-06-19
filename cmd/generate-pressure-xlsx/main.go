package main

import (
	"flag"
	"fmt"
	"log"
	"math"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

const (
	bytesPerMB       = 1024 * 1024
	excelMaxRows     = 1048576
	defaultSheet     = "PressureData"
	defaultCellChars = 48
)

func main() {
	outDir := flag.String("out", filepath.Join("..", "pressure-table"), "output directory")
	targetsText := flag.String("targets", "10,50,100", "target XLSX sizes in MB, comma separated")
	cols := flag.Int("cols", 24, "number of filled columns")
	cellChars := flag.Int("cell-chars", defaultCellChars, "Chinese characters per data cell")
	tolerance := flag.Float64("tolerance", 0.05, "accepted file size tolerance ratio")
	maxAttempts := flag.Int("max-attempts", 6, "maximum row-count calibration attempts per target")
	flag.Parse()

	targets, err := parseTargets(*targetsText)
	if err != nil {
		log.Fatal(err)
	}
	if *cols <= 0 {
		log.Fatal("cols must be greater than 0")
	}
	if *cellChars <= 0 {
		log.Fatal("cell-chars must be greater than 0")
	}
	if *tolerance <= 0 {
		log.Fatal("tolerance must be greater than 0")
	}
	if *maxAttempts <= 0 {
		log.Fatal("max-attempts must be greater than 0")
	}

	if err := os.MkdirAll(*outDir, 0o755); err != nil {
		log.Fatalf("create output directory: %v", err)
	}

	for _, targetMB := range targets {
		targetBytes := int64(targetMB * bytesPerMB)
		outputPath := filepath.Join(*outDir, fmt.Sprintf("pressure-%dmb.xlsx", targetMB))
		if err := generateTarget(outputPath, targetBytes, *cols, *cellChars, *tolerance, *maxAttempts); err != nil {
			log.Fatal(err)
		}
	}
}

func generateTarget(outputPath string, targetBytes int64, cols int, cellChars int, tolerance float64, maxAttempts int) error {
	start := time.Now()
	rows := estimateRows(targetBytes, cols, cellChars)
	var bestPath string
	var bestSize int64
	bestDelta := math.MaxFloat64

	for attempt := 1; attempt <= maxAttempts; attempt++ {
		tempPath := fmt.Sprintf("%s.attempt-%d.tmp.xlsx", outputPath, attempt)
		_ = os.Remove(tempPath)

		if err := writeWorkbook(tempPath, rows, cols, cellChars); err != nil {
			return err
		}

		info, err := os.Stat(tempPath)
		if err != nil {
			return fmt.Errorf("stat generated file: %w", err)
		}

		size := info.Size()
		delta := math.Abs(float64(size-targetBytes)) / float64(targetBytes)
		fmt.Printf(
			"%s attempt %d: rows=%d cols=%d size=%.2fMB target=%.2fMB delta=%.2f%%\n",
			filepath.Base(outputPath),
			attempt,
			rows-1,
			cols,
			float64(size)/bytesPerMB,
			float64(targetBytes)/bytesPerMB,
			delta*100,
		)

		if delta < bestDelta {
			if bestPath != "" && bestPath != tempPath {
				_ = os.Remove(bestPath)
			}
			bestPath = tempPath
			bestSize = size
			bestDelta = delta
		} else {
			_ = os.Remove(tempPath)
		}

		if delta <= tolerance {
			break
		}

		nextRows := int(math.Round(float64(rows) * float64(targetBytes) / float64(size)))
		if size < targetBytes {
			nextRows = int(math.Ceil(float64(nextRows) * 1.02))
		}
		rows = clampRows(nextRows)
	}

	_ = os.Remove(outputPath)
	if err := os.Rename(bestPath, outputPath); err != nil {
		return fmt.Errorf("move generated file: %w", err)
	}

	fmt.Printf(
		"done %s: size=%.2fMB delta=%.2f%% elapsed=%s\n",
		outputPath,
		float64(bestSize)/bytesPerMB,
		bestDelta*100,
		time.Since(start).Round(time.Millisecond),
	)
	return nil
}

func writeWorkbook(path string, rows int, cols int, cellChars int) error {
	workbook := excelize.NewFile()
	defer func() {
		if err := workbook.Close(); err != nil {
			log.Printf("close workbook: %v", err)
		}
	}()

	defaultSheetName := workbook.GetSheetName(0)
	if err := workbook.SetSheetName(defaultSheetName, defaultSheet); err != nil {
		return fmt.Errorf("rename sheet: %w", err)
	}

	streamWriter, err := workbook.NewStreamWriter(defaultSheet)
	if err != nil {
		return fmt.Errorf("create stream writer: %w", err)
	}

	header := make([]interface{}, cols)
	for col := range header {
		header[col] = fmt.Sprintf("\u5b57\u6bb5_%03d", col+1)
	}
	if err := streamWriter.SetRow("A1", header); err != nil {
		return fmt.Errorf("write header row: %w", err)
	}

	for row := 2; row <= rows; row++ {
		values := make([]interface{}, cols)
		for col := range values {
			values[col] = makeCellValue(row, col+1, cellChars)
		}

		cell, err := excelize.CoordinatesToCellName(1, row)
		if err != nil {
			return fmt.Errorf("make row cell name: %w", err)
		}
		if err := streamWriter.SetRow(cell, values); err != nil {
			return fmt.Errorf("write row %d: %w", row, err)
		}
	}

	if err := streamWriter.Flush(); err != nil {
		return fmt.Errorf("flush stream writer: %w", err)
	}
	if err := workbook.SaveAs(path); err != nil {
		return fmt.Errorf("save workbook: %w", err)
	}
	return nil
}

func makeCellValue(row int, col int, length int) string {
	seed := uint64(row)*11400714819323198485 ^ uint64(col)*14029467366897019727
	runes := make([]rune, 0, length+24)
	runes = append(runes, '\u884c')
	runes = append(runes, []rune(strconv.Itoa(row-1))...)
	runes = append(runes, '\u5217')
	runes = append(runes, []rune(strconv.Itoa(col))...)
	runes = append(runes, '_')

	for len(runes) < length {
		seed = nextSeed(seed)
		runes = append(runes, chineseRunes[seed%uint64(len(chineseRunes))])
	}
	return string(runes)
}

func nextSeed(seed uint64) uint64 {
	seed ^= seed << 13
	seed ^= seed >> 7
	seed ^= seed << 17
	return seed
}

func estimateRows(targetBytes int64, cols int, cellChars int) int {
	estimatedCompressedCellBytes := math.Max(16, float64(cellChars)*1.25)
	rows := int(math.Ceil(float64(targetBytes) / (float64(cols) * estimatedCompressedCellBytes)))
	return clampRows(rows + 1)
}

func clampRows(rows int) int {
	if rows < 2 {
		return 2
	}
	if rows > excelMaxRows {
		return excelMaxRows
	}
	return rows
}

func parseTargets(text string) ([]int, error) {
	parts := strings.Split(text, ",")
	targets := make([]int, 0, len(parts))
	for _, part := range parts {
		part = strings.TrimSpace(part)
		if part == "" {
			continue
		}
		target, err := strconv.Atoi(part)
		if err != nil {
			return nil, fmt.Errorf("parse target %q: %w", part, err)
		}
		if target <= 0 {
			return nil, fmt.Errorf("target must be greater than 0: %d", target)
		}
		targets = append(targets, target)
	}
	if len(targets) == 0 {
		return nil, fmt.Errorf("no targets provided")
	}
	return targets, nil
}

var chineseRunes = []rune{
	'\u6570', '\u636e', '\u538b', '\u529b', '\u6d4b', '\u8bd5', '\u8868', '\u683c',
	'\u4e2d', '\u6587', '\u5185', '\u5bb9', '\u57ce', '\u5e02', '\u59d3', '\u540d',
	'\u5730', '\u5740', '\u5355', '\u4f4d', '\u90e8', '\u95e8', '\u72b6', '\u6001',
	'\u8ba2', '\u5355', '\u91d1', '\u989d', '\u65f6', '\u95f4', '\u5907', '\u6ce8',
	'\u7f16', '\u53f7', '\u7c7b', '\u578b', '\u8d28', '\u91cf', '\u5ba2', '\u6237',
	'\u4ea7', '\u54c1', '\u9879', '\u76ee', '\u7edf', '\u8ba1', '\u5206', '\u6790',
	'\u5b57', '\u6bb5', '\u8bb0', '\u5f55', '\u6e05', '\u5355', '\u8d44', '\u6599',
	'\u5317', '\u4eac', '\u4e0a', '\u6d77', '\u5e7f', '\u5dde', '\u6df1', '\u5733',
}
