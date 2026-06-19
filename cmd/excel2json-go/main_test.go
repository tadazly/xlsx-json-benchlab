package main

import (
	"encoding/json"
	"os"
	"path/filepath"
	"testing"

	"github.com/xuri/excelize/v2"
)

func TestConvertChineseRows(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "\u4e2d\u6587\u8868\u683c.xlsx")
	outputPath := filepath.Join(tempDir, "output.json")

	writeWorkbook(t, inputPath, [][]any{
		{"\u59d3\u540d", "\u57ce\u5e02", "\u5907\u6ce8"},
		{"\u5f20\u4e09", "\u5317\u4eac", "\u4e2d\u6587\u6b63\u5e38"},
		{"\u674e\u56db", "\u4e0a\u6d77", "\u7b2c\u4e8c\u884c"},
	})

	rowCount, err := convert(inputPath, outputPath, "")
	if err != nil {
		t.Fatal(err)
	}
	if rowCount != 2 {
		t.Fatalf("row count = %d, want 2", rowCount)
	}

	records := readRecords(t, outputPath)
	if records[0]["\u59d3\u540d"] != "\u5f20\u4e09" || records[1]["\u57ce\u5e02"] != "\u4e0a\u6d77" {
		t.Fatalf("unexpected records: %#v", records)
	}
}

func TestConvertOnlyHeaderColumnsWithValues(t *testing.T) {
	tempDir := t.TempDir()
	inputPath := filepath.Join(tempDir, "sparse.xlsx")
	outputPath := filepath.Join(tempDir, "output.json")

	writeWorkbook(t, inputPath, [][]any{
		{"\u59d3\u540d", "", "\u57ce\u5e02", ""},
		{"\u5f20\u4e09", "ignored", "\u5317\u4eac", "ignored"},
		{"", "ignored", "", "ignored"},
	})

	rowCount, err := convert(inputPath, outputPath, "")
	if err != nil {
		t.Fatal(err)
	}
	if rowCount != 1 {
		t.Fatalf("row count = %d, want 1", rowCount)
	}

	records := readRecords(t, outputPath)
	if len(records[0]) != 2 {
		t.Fatalf("record column count = %d, want 2: %#v", len(records[0]), records[0])
	}
	if _, ok := records[0]["column_2"]; ok {
		t.Fatalf("unexpected empty header column: %#v", records[0])
	}
}

func writeWorkbook(t *testing.T, path string, rows [][]any) {
	t.Helper()

	workbook := excelize.NewFile()
	sheet := workbook.GetSheetName(0)
	for rowIndex, row := range rows {
		cell, err := excelize.CoordinatesToCellName(1, rowIndex+1)
		if err != nil {
			t.Fatal(err)
		}
		if err := workbook.SetSheetRow(sheet, cell, &row); err != nil {
			t.Fatal(err)
		}
	}
	if err := workbook.SaveAs(path); err != nil {
		t.Fatal(err)
	}
	if err := workbook.Close(); err != nil {
		t.Fatal(err)
	}
}

func readRecords(t *testing.T, path string) []map[string]string {
	t.Helper()

	content, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}

	var records []map[string]string
	if err := json.Unmarshal(content, &records); err != nil {
		t.Fatal(err)
	}
	return records
}
