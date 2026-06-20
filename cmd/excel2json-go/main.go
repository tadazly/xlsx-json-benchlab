package main

import (
	"bufio"
	"errors"
	"flag"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

func main() {
	input := flag.String("input", "", "input .xlsx file")
	output := flag.String("output", "", "output .json file")
	sheet := flag.String("sheet", "", "sheet name, defaults to the first sheet")
	flag.Parse()

	if *input == "" && flag.NArg() > 0 {
		*input = flag.Arg(0)
	}
	if *output == "" && flag.NArg() > 1 {
		*output = flag.Arg(1)
	}
	if *input == "" {
		log.Fatal("missing input file: use -input data.xlsx [-output data.json] [-sheet Sheet1]")
	}
	if *output == "" {
		*output = defaultOutputPath(*input)
	}

	start := time.Now()
	rows, err := convert(*input, *output, *sheet)
	if err != nil {
		log.Fatal(err)
	}

	fmt.Printf("Converted %d data rows to %s in %s\n", rows, *output, time.Since(start).Round(time.Millisecond))
}

func convert(inputPath, outputPath, sheetName string) (int, error) {
	if strings.ToLower(filepath.Ext(inputPath)) != ".xlsx" {
		return 0, fmt.Errorf("input must be an .xlsx file: %s", inputPath)
	}

	workbook, err := excelize.OpenFile(inputPath)
	if err != nil {
		return 0, fmt.Errorf("open excel file: %w", err)
	}
	defer func() {
		if closeErr := workbook.Close(); closeErr != nil {
			log.Printf("close workbook: %v", closeErr)
		}
	}()

	if sheetName == "" {
		sheets := workbook.GetSheetList()
		if len(sheets) == 0 {
			return 0, errors.New("workbook has no sheets")
		}
		sheetName = sheets[0]
	}

	streamRows, err := workbook.Rows(sheetName)
	if err != nil {
		return 0, fmt.Errorf("open sheet %q: %w", sheetName, err)
	}
	defer func() {
		if closeErr := streamRows.Close(); closeErr != nil {
			log.Printf("close sheet rows: %v", closeErr)
		}
	}()

	if err := os.MkdirAll(filepath.Dir(outputPath), 0o755); err != nil && filepath.Dir(outputPath) != "." {
		return 0, fmt.Errorf("create output directory: %w", err)
	}

	outputFile, err := os.Create(outputPath)
	if err != nil {
		return 0, fmt.Errorf("create output file: %w", err)
	}
	defer func() {
		if closeErr := outputFile.Close(); closeErr != nil {
			log.Printf("close output file: %v", closeErr)
		}
	}()

	writer := bufio.NewWriterSize(outputFile, 1024*1024)
	defer func() {
		if err := writer.Flush(); err != nil {
			log.Printf("flush output file: %v", err)
		}
	}()

	if _, err := writer.WriteString("[\n"); err != nil {
		return 0, fmt.Errorf("write json start: %w", err)
	}

	var headerColumns []headerColumn
	rowBuffer := make([]byte, 0, 4096)
	rowCount := 0
	firstObject := true

	for streamRows.Next() {
		rowValues, err := streamRows.Columns()
		if err != nil {
			return rowCount, fmt.Errorf("read row: %w", err)
		}

		if headerColumns == nil {
			headerColumns = selectHeaderColumns(rowValues)
			if len(headerColumns) == 0 {
				return 0, errors.New("header row has no non-empty columns")
			}
			continue
		}
		if isEmptyRow(rowValues, headerColumns) {
			continue
		}

		if !firstObject {
			if _, err := writer.WriteString(",\n"); err != nil {
				return rowCount, fmt.Errorf("write json comma: %w", err)
			}
		}
		firstObject = false

		var writeErr error
		rowBuffer, writeErr = writeJSONObject(writer, headerColumns, rowValues, rowBuffer)
		if writeErr != nil {
			return rowCount, fmt.Errorf("write json row: %w", writeErr)
		}
		rowCount++
	}
	if err := streamRows.Error(); err != nil {
		return rowCount, fmt.Errorf("iterate sheet rows: %w", err)
	}
	if headerColumns == nil {
		return 0, errors.New("sheet is empty")
	}

	if _, err := writer.WriteString("\n]\n"); err != nil {
		return rowCount, fmt.Errorf("write json end: %w", err)
	}
	if err := writer.Flush(); err != nil {
		return rowCount, fmt.Errorf("flush output file: %w", err)
	}

	return rowCount, nil
}

type headerColumn struct {
	index        int
	header       string
	quotedHeader string
}

func selectHeaderColumns(row []string) []headerColumn {
	columns := make([]headerColumn, 0, len(row))
	seen := make(map[string]int, len(row))

	for index, value := range row {
		header := strings.TrimSpace(value)
		if header == "" {
			continue
		}

		seen[header]++
		if seen[header] > 1 {
			header = fmt.Sprintf("%s_%d", header, seen[header])
		}
		columns = append(columns, headerColumn{
			index:        index,
			header:       header,
			quotedHeader: string(appendJSONString(nil, header)),
		})
	}

	return columns
}

func isEmptyRow(row []string, columns []headerColumn) bool {
	for _, column := range columns {
		if column.index < len(row) && strings.TrimSpace(row[column.index]) != "" {
			return false
		}
	}
	return true
}

func writeJSONObject(writer *bufio.Writer, columns []headerColumn, row []string, buffer []byte) ([]byte, error) {
	buffer = buffer[:0]
	buffer = append(buffer, '{')
	for index, column := range columns {
		if index > 0 {
			buffer = append(buffer, ',')
		}

		buffer = append(buffer, column.quotedHeader...)
		buffer = append(buffer, ':')

		value := ""
		if column.index < len(row) {
			value = row[column.index]
		}
		buffer = appendJSONString(buffer, value)
	}
	buffer = append(buffer, '}')

	_, err := writer.Write(buffer)
	return buffer, err
}

func defaultOutputPath(inputPath string) string {
	extension := filepath.Ext(inputPath)
	return strings.TrimSuffix(inputPath, extension) + ".json"
}

func appendJSONString(buffer []byte, value string) []byte {
	buffer = append(buffer, '"')
	start := 0
	for index := 0; index < len(value); index++ {
		var escaped string
		switch value[index] {
		case '\\':
			escaped = `\\`
		case '"':
			escaped = `\"`
		case '\b':
			escaped = `\b`
		case '\f':
			escaped = `\f`
		case '\n':
			escaped = `\n`
		case '\r':
			escaped = `\r`
		case '\t':
			escaped = `\t`
		default:
			if value[index] < 0x20 {
				buffer = append(buffer, value[start:index]...)
				buffer = append(buffer, `\u00`...)
				const hex = "0123456789abcdef"
				buffer = append(buffer, hex[value[index]>>4], hex[value[index]&0x0f])
				start = index + 1
			}
			continue
		}

		buffer = append(buffer, value[start:index]...)
		buffer = append(buffer, escaped...)
		start = index + 1
	}
	buffer = append(buffer, value[start:]...)
	buffer = append(buffer, '"')
	return buffer
}
