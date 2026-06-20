package main

import (
	"archive/zip"
	"bufio"
	"bytes"
	"errors"
	"flag"
	"fmt"
	"html"
	"io"
	"log"
	"os"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

func main() {
	input := flag.String("input", "", "input .xlsx file or directory")
	output := flag.String("output", "", "output .json file or directory")
	sheet := flag.String("sheet", "", "sheet name, defaults to the first sheet")
	flat := flag.Bool("flat", false, "when input is a directory, write all JSON files directly under the output directory")
	maxConcurrency := flag.Int("concurrency", 0, "maximum concurrent conversions for directory input, defaults to logical CPU cores")
	flag.IntVar(maxConcurrency, "j", 0, "shorthand for -concurrency")
	flag.Parse()

	if *input == "" && flag.NArg() > 0 {
		*input = flag.Arg(0)
	}
	if *output == "" && flag.NArg() > 1 {
		*output = flag.Arg(1)
	}
	if *input == "" {
		log.Fatal("missing input: use -input data.xlsx|dir [-output data.json|dir] [-sheet Sheet1] [-flat] [-concurrency N]")
	}

	start := time.Now()
	results, err := runConversions(*input, *output, *sheet, *flat, *maxConcurrency)
	if err != nil {
		log.Fatal(err)
	}

	failures := 0
	totalRows := 0
	for _, result := range results {
		if result.err != nil {
			failures++
			fmt.Printf("FAILED %s -> %s in %s: %v\n", result.input, result.output, result.elapsed.Round(time.Millisecond), result.err)
			continue
		}
		totalRows += result.rows
		fmt.Printf("Converted %d data rows: %s -> %s in %s\n", result.rows, result.input, result.output, result.elapsed.Round(time.Millisecond))
	}

	fmt.Printf("Finished %d file(s), %d failed, %d total data rows in %s\n", len(results), failures, totalRows, time.Since(start).Round(time.Millisecond))
	if failures > 0 {
		os.Exit(1)
	}
}

type conversionTask struct {
	input  string
	output string
}

type conversionResult struct {
	input   string
	output  string
	rows    int
	elapsed time.Duration
	err     error
}

func runConversions(inputPath, outputPath, sheetName string, flat bool, maxConcurrency int) ([]conversionResult, error) {
	info, err := os.Stat(inputPath)
	if err != nil {
		return nil, fmt.Errorf("stat input: %w", err)
	}

	if !info.IsDir() {
		output := outputPath
		if output == "" {
			output = defaultOutputPath(inputPath)
		} else if outputInfo, err := os.Stat(output); err == nil && outputInfo.IsDir() {
			output = filepath.Join(output, defaultOutputPath(filepath.Base(inputPath)))
		}

		start := time.Now()
		rows, err := convertFast(inputPath, output, sheetName)
		return []conversionResult{{
			input:   inputPath,
			output:  output,
			rows:    rows,
			elapsed: time.Since(start),
			err:     err,
		}}, nil
	}

	tasks, err := collectDirectoryTasks(inputPath, outputPath, flat)
	if err != nil {
		return nil, err
	}
	if len(tasks) == 0 {
		return nil, fmt.Errorf("no .xlsx files found under %s", inputPath)
	}

	concurrency := optimalConcurrency(maxConcurrency, len(tasks))
	fmt.Printf("Discovered %d .xlsx file(s), concurrency=%d\n", len(tasks), concurrency)

	taskCh := make(chan conversionTask)
	resultCh := make(chan conversionResult, len(tasks))

	var wg sync.WaitGroup
	for index := 0; index < concurrency; index++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for task := range taskCh {
				start := time.Now()
				rows, err := convertFast(task.input, task.output, sheetName)
				resultCh <- conversionResult{
					input:   task.input,
					output:  task.output,
					rows:    rows,
					elapsed: time.Since(start),
					err:     err,
				}
			}
		}()
	}

	go func() {
		for _, task := range tasks {
			taskCh <- task
		}
		close(taskCh)
		wg.Wait()
		close(resultCh)
	}()

	results := make([]conversionResult, 0, len(tasks))
	for result := range resultCh {
		results = append(results, result)
	}
	sort.Slice(results, func(i, j int) bool {
		return results[i].input < results[j].input
	})
	return results, nil
}

func collectDirectoryTasks(inputDir, outputDir string, flat bool) ([]conversionTask, error) {
	if outputDir == "" {
		outputDir = inputDir
	}

	var inputs []string
	if err := filepath.WalkDir(inputDir, func(path string, entry os.DirEntry, err error) error {
		if err != nil {
			return err
		}
		if entry.IsDir() {
			return nil
		}
		if isExcelFile(path) {
			inputs = append(inputs, path)
		}
		return nil
	}); err != nil {
		return nil, fmt.Errorf("walk input directory: %w", err)
	}
	sort.Strings(inputs)

	tasks := make([]conversionTask, 0, len(inputs))
	usedFlatNames := make(map[string]int, len(inputs))
	for _, input := range inputs {
		relative, err := filepath.Rel(inputDir, input)
		if err != nil {
			return nil, fmt.Errorf("make relative path for %s: %w", input, err)
		}

		var output string
		if flat {
			output = filepath.Join(outputDir, uniqueFlatOutputName(defaultOutputPath(filepath.Base(input)), usedFlatNames))
		} else {
			output = filepath.Join(outputDir, defaultOutputPath(relative))
		}
		tasks = append(tasks, conversionTask{input: input, output: output})
	}
	return tasks, nil
}

func isExcelFile(path string) bool {
	name := filepath.Base(path)
	return strings.EqualFold(filepath.Ext(name), ".xlsx") && !strings.HasPrefix(name, "~$")
}

func uniqueFlatOutputName(name string, used map[string]int) string {
	count := used[name]
	used[name] = count + 1
	if count == 0 {
		return name
	}

	ext := filepath.Ext(name)
	base := strings.TrimSuffix(name, ext)
	return fmt.Sprintf("%s_%d%s", base, count+1, ext)
}

func optimalConcurrency(maxConcurrency int, taskCount int) int {
	concurrency := runtime.NumCPU()
	if maxConcurrency > 0 && maxConcurrency < concurrency {
		concurrency = maxConcurrency
	}
	if concurrency < 1 {
		concurrency = 1
	}
	if taskCount > 0 && concurrency > taskCount {
		concurrency = taskCount
	}
	return concurrency
}

func convertFast(inputPath, outputPath, sheetName string) (int, error) {
	if strings.ToLower(filepath.Ext(inputPath)) != ".xlsx" {
		return 0, fmt.Errorf("input must be an .xlsx file: %s", inputPath)
	}

	zipReader, err := zip.OpenReader(inputPath)
	if err != nil {
		return 0, fmt.Errorf("open xlsx: %w", err)
	}
	defer func() {
		if closeErr := zipReader.Close(); closeErr != nil {
			log.Printf("close xlsx: %v", closeErr)
		}
	}()

	entries := make(map[string]*zip.File, len(zipReader.File))
	for _, entry := range zipReader.File {
		entries[filepath.ToSlash(entry.Name)] = entry
	}

	workbookXML, err := readZipEntry(entries, "xl/workbook.xml")
	if err != nil {
		return 0, err
	}
	relsXML, err := readZipEntry(entries, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return 0, err
	}

	sheets := parseSheets(workbookXML)
	if len(sheets) == 0 {
		return 0, errors.New("workbook has no sheets")
	}

	sheetInfo, err := selectSheet(sheets, sheetName)
	if err != nil {
		return 0, err
	}

	relationships := parseRelationships(relsXML)
	target, ok := relationships[sheetInfo.relationshipID]
	if !ok {
		return 0, fmt.Errorf("worksheet relationship not found: %s", sheetInfo.relationshipID)
	}

	sheetPath := normalizeZipPath("xl", target)
	sheetEntry, ok := entries[sheetPath]
	if !ok {
		return 0, fmt.Errorf("xlsx entry not found: %s", sheetPath)
	}

	var sharedStrings []string
	if _, ok := entries["xl/sharedStrings.xml"]; ok {
		sharedXML, err := readZipEntry(entries, "xl/sharedStrings.xml")
		if err != nil {
			return 0, err
		}
		sharedStrings = parseSharedStrings(sharedXML)
	}

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

	writer := bufio.NewWriterSize(outputFile, 4*1024*1024)
	defer func() {
		if flushErr := writer.Flush(); flushErr != nil {
			log.Printf("flush output file: %v", flushErr)
		}
	}()

	sheetReader, err := sheetEntry.Open()
	if err != nil {
		return 0, fmt.Errorf("open worksheet entry %s: %w", sheetPath, err)
	}
	defer func() {
		if closeErr := sheetReader.Close(); closeErr != nil {
			log.Printf("close worksheet entry %s: %v", sheetPath, closeErr)
		}
	}()

	rowCount, err := parseWorksheetToJSON(sheetReader, sharedStrings, writer)
	if err != nil {
		return rowCount, err
	}
	if err := writer.Flush(); err != nil {
		return rowCount, fmt.Errorf("flush output file: %w", err)
	}

	return rowCount, nil
}

func readZipEntry(entries map[string]*zip.File, name string) ([]byte, error) {
	entry, ok := entries[name]
	if !ok {
		return nil, fmt.Errorf("xlsx entry not found: %s", name)
	}

	reader, err := entry.Open()
	if err != nil {
		return nil, fmt.Errorf("open xlsx entry %s: %w", name, err)
	}
	defer func() {
		if closeErr := reader.Close(); closeErr != nil {
			log.Printf("close xlsx entry %s: %v", name, closeErr)
		}
	}()

	data, err := io.ReadAll(reader)
	if err != nil {
		return nil, fmt.Errorf("read xlsx entry %s: %w", name, err)
	}
	return data, nil
}

type sheetInfo struct {
	name           string
	relationshipID string
}

func parseSheets(xml []byte) []sheetInfo {
	var sheets []sheetInfo
	offset := 0
	for {
		start := bytes.Index(xml[offset:], []byte("<sheet "))
		if start < 0 {
			break
		}
		start += offset
		end := bytes.IndexByte(xml[start:], '>')
		if end < 0 {
			break
		}
		tag := xml[start : start+end+1]
		sheets = append(sheets, sheetInfo{
			name:           attrValue(tag, "name"),
			relationshipID: attrValue(tag, "r:id"),
		})
		offset = start + end + 1
	}
	return sheets
}

func selectSheet(sheets []sheetInfo, requestedName string) (sheetInfo, error) {
	if requestedName == "" {
		return sheets[0], nil
	}
	for _, sheet := range sheets {
		if sheet.name == requestedName {
			return sheet, nil
		}
	}
	return sheetInfo{}, fmt.Errorf("sheet not found: %s", requestedName)
}

func parseRelationships(xml []byte) map[string]string {
	relationships := make(map[string]string)
	offset := 0
	for {
		start := bytes.Index(xml[offset:], []byte("<Relationship "))
		if start < 0 {
			break
		}
		start += offset
		end := bytes.IndexByte(xml[start:], '>')
		if end < 0 {
			break
		}
		tag := xml[start : start+end+1]
		id := attrValue(tag, "Id")
		target := attrValue(tag, "Target")
		if id != "" && target != "" {
			relationships[id] = target
		}
		offset = start + end + 1
	}
	return relationships
}

func parseSharedStrings(xml []byte) []string {
	var values []string
	offset := 0
	for {
		start := bytes.Index(xml[offset:], []byte("<si"))
		if start < 0 {
			break
		}
		start += offset
		openEnd := bytes.IndexByte(xml[start:], '>')
		if openEnd < 0 {
			break
		}
		contentStart := start + openEnd + 1
		closeStart := bytes.Index(xml[contentStart:], []byte("</si>"))
		if closeStart < 0 {
			break
		}
		contentEnd := contentStart + closeStart
		values = append(values, collectTextTags(xml[contentStart:contentEnd]))
		offset = contentEnd + len("</si>")
	}
	return values
}

func parseWorksheetToJSON(reader io.Reader, sharedStrings []string, writer *bufio.Writer) (int, error) {
	if _, err := writer.WriteString("[\n"); err != nil {
		return 0, fmt.Errorf("write json start: %w", err)
	}

	var headerColumns []headerColumn
	rowCount := 0
	firstObject := true
	rowBuffer := make([]byte, 0, 4096)
	scanner := newRowScanner(reader)

	for {
		rowXML, err := scanner.next()
		if err != nil {
			return rowCount, err
		}
		if rowXML == nil {
			break
		}
		row := parseRowElement(rowXML, sharedStrings)

		if headerColumns == nil {
			headerColumns = selectHeaderColumns(row)
			if len(headerColumns) == 0 {
				return 0, errors.New("header row has no non-empty columns")
			}
		} else if !isEmptyRow(row, headerColumns) {
			if !firstObject {
				if _, err := writer.WriteString(",\n"); err != nil {
					return rowCount, fmt.Errorf("write json comma: %w", err)
				}
			}
			firstObject = false

			var err error
			rowBuffer, err = writeJSONObject(writer, headerColumns, row, rowBuffer)
			if err != nil {
				return rowCount, fmt.Errorf("write json row: %w", err)
			}
			rowCount++
		}
	}

	if headerColumns == nil {
		return 0, errors.New("sheet is empty")
	}
	if _, err := writer.WriteString("\n]\n"); err != nil {
		return rowCount, fmt.Errorf("write json end: %w", err)
	}

	return rowCount, nil
}

type rowScanner struct {
	reader io.Reader
	buffer []byte
	chunk  []byte
	eof    bool
}

func newRowScanner(reader io.Reader) *rowScanner {
	return &rowScanner{
		reader: reader,
		buffer: make([]byte, 0, 2*1024*1024),
		chunk:  make([]byte, 256*1024),
	}
}

func (scanner *rowScanner) next() ([]byte, error) {
	for {
		if row := scanner.popRow(); row != nil {
			return row, nil
		}
		if scanner.eof {
			return nil, nil
		}

		n, err := scanner.reader.Read(scanner.chunk)
		if n > 0 {
			scanner.buffer = append(scanner.buffer, scanner.chunk[:n]...)
		}
		if err != nil {
			if errors.Is(err, io.EOF) {
				scanner.eof = true
				continue
			}
			return nil, fmt.Errorf("read worksheet XML: %w", err)
		}
	}
}

func (scanner *rowScanner) popRow() []byte {
	rowStart := bytes.Index(scanner.buffer, []byte("<row"))
	if rowStart < 0 {
		scanner.trimPrefixForNeedle([]byte("<row"))
		return nil
	}
	if rowStart > 0 {
		scanner.buffer = scanner.buffer[rowStart:]
	}

	rowEnd := bytes.Index(scanner.buffer, []byte("</row>"))
	if rowEnd < 0 {
		return nil
	}
	rowEnd += len("</row>")
	row := append([]byte(nil), scanner.buffer[:rowEnd]...)
	scanner.buffer = scanner.buffer[rowEnd:]
	return row
}

func (scanner *rowScanner) trimPrefixForNeedle(needle []byte) {
	keep := len(needle) - 1
	if keep <= 0 || len(scanner.buffer) <= keep {
		return
	}
	copy(scanner.buffer, scanner.buffer[len(scanner.buffer)-keep:])
	scanner.buffer = scanner.buffer[:keep]
}

func parseRowElement(rowXML []byte, sharedStrings []string) []string {
	openEnd := bytes.IndexByte(rowXML, '>')
	if openEnd < 0 {
		return nil
	}
	closeStart := bytes.LastIndex(rowXML, []byte("</row>"))
	if closeStart < 0 || closeStart <= openEnd {
		return nil
	}
	return parseRow(rowXML[openEnd+1:closeStart], sharedStrings)
}

func parseRow(rowXML []byte, sharedStrings []string) []string {
	var values []string
	nextImplicitIndex := 0
	offset := 0

	for {
		cellStart := bytes.Index(rowXML[offset:], []byte("<c"))
		if cellStart < 0 {
			break
		}
		cellStart += offset
		cellOpenEnd := bytes.IndexByte(rowXML[cellStart:], '>')
		if cellOpenEnd < 0 {
			break
		}
		cellTag := rowXML[cellStart : cellStart+cellOpenEnd+1]
		index := cellRefToColumnIndex(attrValue(cellTag, "r"))
		if index < 0 {
			index = nextImplicitIndex
		}
		nextImplicitIndex = index + 1

		if index >= len(values) {
			values = append(values, make([]string, index-len(values)+1)...)
		}

		if isSelfClosingTag(cellTag) {
			values[index] = ""
			offset = cellStart + cellOpenEnd + 1
			continue
		}

		cellContentStart := cellStart + cellOpenEnd + 1
		cellClose := bytes.Index(rowXML[cellContentStart:], []byte("</c>"))
		if cellClose < 0 {
			break
		}
		cellContentEnd := cellContentStart + cellClose
		cellContent := rowXML[cellContentStart:cellContentEnd]

		values[index] = resolveCellValue(attrValue(cellTag, "t"), cellContent, sharedStrings)
		offset = cellContentEnd + len("</c>")
	}

	return values
}

func isSelfClosingTag(tag []byte) bool {
	for index := len(tag) - 2; index >= 0; index-- {
		switch tag[index] {
		case ' ', '\t', '\n', '\r':
			continue
		case '/':
			return true
		default:
			return false
		}
	}
	return false
}
func resolveCellValue(cellType string, content []byte, sharedStrings []string) string {
	switch cellType {
	case "inlineStr":
		return collectTextTags(content)
	case "s":
		indexText := textInTag(content, "v")
		index, err := strconv.Atoi(strings.TrimSpace(indexText))
		if err != nil || index < 0 || index >= len(sharedStrings) {
			return ""
		}
		return sharedStrings[index]
	default:
		return textInTag(content, "v")
	}
}

type headerColumn struct {
	index        int
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

func collectTextTags(content []byte) string {
	var builder strings.Builder
	offset := 0
	for {
		start := bytes.Index(content[offset:], []byte("<t"))
		if start < 0 {
			break
		}
		start += offset
		openEnd := bytes.IndexByte(content[start:], '>')
		if openEnd < 0 {
			break
		}
		textStart := start + openEnd + 1
		textEnd := bytes.Index(content[textStart:], []byte("</t>"))
		if textEnd < 0 {
			break
		}
		raw := content[textStart : textStart+textEnd]
		builder.WriteString(xmlText(raw))
		offset = textStart + textEnd + len("</t>")
	}
	return builder.String()
}

func textInTag(content []byte, tag string) string {
	open := []byte("<" + tag)
	close := []byte("</" + tag + ">")
	start := bytes.Index(content, open)
	if start < 0 {
		return ""
	}
	openEnd := bytes.IndexByte(content[start:], '>')
	if openEnd < 0 {
		return ""
	}
	textStart := start + openEnd + 1
	textEnd := bytes.Index(content[textStart:], close)
	if textEnd < 0 {
		return ""
	}
	return xmlText(content[textStart : textStart+textEnd])
}

func xmlText(raw []byte) string {
	if !bytes.Contains(raw, []byte("&")) {
		return string(raw)
	}
	return html.UnescapeString(string(raw))
}

func attrValue(tag []byte, name string) string {
	needle := []byte(" " + name + "=\"")
	start := bytes.Index(tag, needle)
	if start < 0 {
		return ""
	}
	start += len(needle)
	end := bytes.IndexByte(tag[start:], '"')
	if end < 0 {
		return ""
	}
	return xmlText(tag[start : start+end])
}

func cellRefToColumnIndex(cellRef string) int {
	index := 0
	seenLetter := false
	for _, char := range strings.ToUpper(cellRef) {
		if char < 'A' || char > 'Z' {
			break
		}
		seenLetter = true
		index = index*26 + int(char-'A'+1)
	}
	if !seenLetter {
		return -1
	}
	return index - 1
}

func normalizeZipPath(baseDir, target string) string {
	target = filepath.ToSlash(target)
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}

	parts := strings.Split(filepath.ToSlash(baseDir+"/"+target), "/")
	normalized := make([]string, 0, len(parts))
	for _, part := range parts {
		switch part {
		case "", ".":
			continue
		case "..":
			if len(normalized) > 0 {
				normalized = normalized[:len(normalized)-1]
			}
		default:
			normalized = append(normalized, part)
		}
	}
	return strings.Join(normalized, "/")
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
