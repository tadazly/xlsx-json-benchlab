package main

import (
	"archive/zip"
	"bufio"
	"bytes"
	"encoding/json"
	"errors"
	"flag"
	"fmt"
	"html"
	"io"
	"log"
	"math"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

type field struct {
	key   string
	value any
}

type rowObject []field

type attrInfo struct {
	name string
	key  string
	typ  string
	des  string
}

type tableData struct {
	content []rowObject
	attrs   []attrInfo
	file    string
}

type itemInfo struct {
	name      string
	fields    []field
	index     map[string]int
	merge     int
	sensitive map[string]bool
}

type task struct {
	name  string
	path  string
	mtime int64
}

type taskResult struct {
	name    string
	path    string
	changed bool
	rows    int
	json    []byte
	err     error
}

type options struct {
	workspace     string
	tableRoot     string
	outputDir     string
	hashPath      string
	nodePath      string
	xlsxModule    string
	all           bool
	concurrency   int
	skipSensitive bool
}

func main() {
	opts := parseFlags()
	start := time.Now()

	if err := run(opts); err != nil {
		log.Fatal(err)
	}
	fmt.Printf("Finished in %s\n", time.Since(start).Round(time.Millisecond))
}

func parseFlags() options {
	var opts options
	flag.StringVar(&opts.workspace, "workspace", "", "workspace-splan root, defaults to WORKSPACE_PATH or current directory")
	flag.StringVar(&opts.tableRoot, "table-root", "", "table excel/common directory, defaults to <workspace>/table/excel/common")
	flag.StringVar(&opts.outputDir, "output", "", "output json directory, defaults to <workspace>/client/publish/resource/json/xls/cn")
	flag.StringVar(&opts.hashPath, "hash", "", "incremental hash json path, defaults to <workspace>/.temp/table/xls_hash_fast.json")
	flag.StringVar(&opts.nodePath, "node", "node", "node executable used only for .xls fallback")
	flag.StringVar(&opts.xlsxModule, "xlsx-module", "", "xlsx module path used by .xls fallback, defaults to <workspace>/node_modules/xlsx")
	flag.BoolVar(&opts.all, "all", false, "rebuild all tables")
	flag.BoolVar(&opts.all, "a", false, "shorthand for -all")
	flag.IntVar(&opts.concurrency, "concurrency", 0, "maximum concurrent table conversions")
	flag.IntVar(&opts.concurrency, "j", 0, "shorthand for -concurrency")
	flag.BoolVar(&opts.skipSensitive, "skip-sensitive", true, "skip fields listed in $items sensitiveField")
	flag.Parse()

	if opts.workspace == "" {
		opts.workspace = os.Getenv("WORKSPACE_PATH")
	}
	if opts.workspace == "" {
		if cwd, err := os.Getwd(); err == nil {
			opts.workspace = cwd
		}
	}
	opts.workspace = filepath.Clean(opts.workspace)
	if opts.tableRoot == "" {
		opts.tableRoot = filepath.Join(opts.workspace, "table", "excel", "common")
	}
	if opts.outputDir == "" {
		opts.outputDir = filepath.Join(opts.workspace, "client", "publish", "resource", "json", "xls", "cn")
	}
	if opts.hashPath == "" {
		opts.hashPath = filepath.Join(opts.workspace, ".temp", "table", "xls_hash_fast.json")
	}
	if opts.xlsxModule == "" {
		opts.xlsxModule = filepath.Join(opts.workspace, "node_modules", "xlsx")
	}
	return opts
}

func run(opts options) error {
	if err := os.MkdirAll(opts.outputDir, 0o755); err != nil {
		return fmt.Errorf("create output dir: %w", err)
	}

	itemsPath := filepath.Join(opts.tableRoot, "$items.xls")
	itemsData, err := xlsToJSON(itemsPath, opts)
	if err != nil {
		return fmt.Errorf("load $items: %w", err)
	}
	items, itemOrder, err := createItems(itemsData)
	if err != nil {
		return err
	}

	hashes := map[string]int64{}
	if !opts.all {
		hashes = readHashFile(opts.hashPath)
	}

	tasks, err := collectTasks(opts.tableRoot, items, hashes, opts.outputDir, opts.all)
	if err != nil {
		return err
	}
	for _, t := range tasks.all {
		item := items[t.name]
		item.set("modification", t.mtime)
	}

	results := processTasks(tasks.changed, opts, items)
	failures := 0
	changed := 0
	mergeJSON := make(map[string]json.RawMessage)
	for _, result := range results {
		if result.err != nil {
			failures++
			fmt.Printf("FAILED %s: %v\n", result.path, result.err)
			continue
		}
		changed++
		hashes[result.name] = fileMTime(result.path)
		if items[result.name].merge == 1 {
			mergeJSON[result.name] = json.RawMessage(result.json)
		}
		fmt.Printf("Converted %d rows: %s\n", result.rows, result.path)
	}
	if failures > 0 {
		return fmt.Errorf("%d table(s) failed", failures)
	}

	for _, name := range itemOrder {
		item := items[name]
		if item.merge != 1 {
			continue
		}
		if _, ok := mergeJSON[name]; ok {
			continue
		}
		path := filepath.Join(opts.outputDir, name+".json")
		data, err := os.ReadFile(path)
		if err != nil {
			return fmt.Errorf("read merge table %s: %w", path, err)
		}
		mergeJSON[name] = json.RawMessage(data)
	}
	for name, raw := range mergeJSON {
		items[name].set("values", raw)
	}

	if err := writeItems(filepath.Join(opts.outputDir, "items.json"), itemOrder, items); err != nil {
		return err
	}
	if err := writeHashFile(opts.hashPath, hashes); err != nil {
		return err
	}

	fmt.Printf("Processed %d/%d configured table(s), rebuilt %d\n", len(tasks.all), len(itemOrder), changed)
	return nil
}

type collectedTasks struct {
	all     []task
	changed []task
}

func collectTasks(root string, items map[string]*itemInfo, hashes map[string]int64, outputDir string, all bool) (collectedTasks, error) {
	var out collectedTasks
	err := filepath.WalkDir(root, func(path string, entry os.DirEntry, walkErr error) error {
		if walkErr != nil {
			return walkErr
		}
		if entry.IsDir() || !isExcelFile(path) {
			return nil
		}
		name := strings.TrimSuffix(filepath.Base(path), filepath.Ext(path))
		if name == "$items" || strings.HasPrefix(name, "$") {
			return nil
		}
		if _, ok := items[name]; !ok {
			return nil
		}
		mtime := fileMTime(path)
		t := task{name: name, path: path, mtime: mtime}
		out.all = append(out.all, t)
		outputPath := filepath.Join(outputDir, name+".json")
		_, statErr := os.Stat(outputPath)
		if all || hashes[name] != mtime || statErr != nil {
			out.changed = append(out.changed, t)
		}
		return nil
	})
	if err != nil {
		return out, fmt.Errorf("walk table root: %w", err)
	}
	sort.Slice(out.all, func(i, j int) bool { return out.all[i].name < out.all[j].name })
	sort.Slice(out.changed, func(i, j int) bool { return out.changed[i].name < out.changed[j].name })
	return out, nil
}

func processTasks(tasks []task, opts options, items map[string]*itemInfo) []taskResult {
	if len(tasks) == 0 {
		return nil
	}
	concurrency := opts.concurrency
	if concurrency <= 0 {
		concurrency = runtime.NumCPU()
	}
	if concurrency > len(tasks) {
		concurrency = len(tasks)
	}
	if concurrency < 1 {
		concurrency = 1
	}

	taskCh := make(chan task)
	resultCh := make(chan taskResult, len(tasks))
	var wg sync.WaitGroup
	for i := 0; i < concurrency; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			for t := range taskCh {
				resultCh <- convertTask(t, opts, items[t.name])
			}
		}()
	}
	go func() {
		for _, t := range tasks {
			taskCh <- t
		}
		close(taskCh)
		wg.Wait()
		close(resultCh)
	}()

	results := make([]taskResult, 0, len(tasks))
	for result := range resultCh {
		results = append(results, result)
	}
	sort.Slice(results, func(i, j int) bool { return results[i].name < results[j].name })
	return results
}

func convertTask(t task, opts options, item *itemInfo) taskResult {
	data, err := xlsToJSON(t.path, opts)
	if err != nil {
		return taskResult{name: t.name, path: t.path, err: err}
	}
	if opts.skipSensitive && len(item.sensitive) > 0 {
		data = filterSensitive(data, item.sensitive)
	}
	for rowIndex := range data.content {
		for fieldIndex := range data.content[rowIndex] {
			if s, ok := data.content[rowIndex][fieldIndex].value.(string); ok {
				data.content[rowIndex][fieldIndex].value = encodeHTML(s)
			}
		}
	}
	jsonBytes, err := marshalRows(data.content)
	if err != nil {
		return taskResult{name: t.name, path: t.path, err: err}
	}
	outputPath := filepath.Join(opts.outputDir, t.name+".json")
	if err := os.WriteFile(outputPath, jsonBytes, 0o644); err != nil {
		return taskResult{name: t.name, path: t.path, err: fmt.Errorf("write %s: %w", outputPath, err)}
	}
	return taskResult{name: t.name, path: t.path, changed: true, rows: len(data.content), json: jsonBytes}
}

func filterSensitive(data tableData, sensitive map[string]bool) tableData {
	filteredAttrs := make([]attrInfo, 0, len(data.attrs))
	for _, attr := range data.attrs {
		if sensitive[attr.key] || sensitive[attr.name] {
			continue
		}
		filteredAttrs = append(filteredAttrs, attr)
	}
	filteredRows := make([]rowObject, 0, len(data.content))
	for _, row := range data.content {
		next := make(rowObject, 0, len(row))
		for _, f := range row {
			if sensitive[f.key] {
				continue
			}
			next = append(next, f)
		}
		filteredRows = append(filteredRows, next)
	}
	data.attrs = filteredAttrs
	data.content = filteredRows
	return data
}

func xlsToJSON(path string, opts options) (tableData, error) {
	sheets, err := readWorkbookRows(path, opts)
	if err != nil {
		return tableData{}, err
	}
	if len(sheets) < 2 {
		return tableData{}, fmt.Errorf("%s has fewer than 2 sheets", path)
	}
	return rowsToTable(path, sheets[0], sheets[1])
}

func rowsToTable(path string, dataRows [][]any, typeRows [][]any) (tableData, error) {
	attrs := make([]attrInfo, 0, len(typeRows))
	seen := map[string]bool{}
	firstAttrName := ""
	for _, row := range typeRows {
		name := cellString(rowValue(row, 0))
		key := cellString(rowValue(row, 1))
		typ := cellString(rowValue(row, 2))
		des := cellString(rowValue(row, 3))
		if name == "" || key == "" {
			continue
		}
		if firstAttrName == "" {
			firstAttrName = name
		}
		if typ == "uint" {
			typ = "number"
		} else if typ == "" {
			typ = "string"
		}
		if seen[name] {
			fmt.Printf("WARN duplicate attr column in %s: %s\n", path, name)
			continue
		}
		seen[name] = true
		attrs = append(attrs, attrInfo{name: name, key: key, typ: typ, des: des})
	}
	if firstAttrName == "" {
		return tableData{}, fmt.Errorf("%s type sheet has no attrs", path)
	}
	if len(dataRows) == 0 {
		return tableData{attrs: attrs, file: path}, nil
	}

	headers := make(map[string]int, len(dataRows[0]))
	for index, value := range dataRows[0] {
		name := cellString(value)
		if name != "" {
			headers[name] = index
		}
	}
	firstIndex, hasFirst := headers[firstAttrName]
	if !hasFirst {
		return tableData{}, fmt.Errorf("%s data sheet missing first attr column %q", path, firstAttrName)
	}

	content := make([]rowObject, 0, len(dataRows)-1)
	for _, row := range dataRows[1:] {
		if isMissing(rowValue(row, firstIndex)) {
			continue
		}
		obj := make(rowObject, 0, len(attrs))
		for _, attr := range attrs {
			columnIndex, ok := headers[attr.name]
			var raw any
			if ok {
				raw = rowValue(row, columnIndex)
			}
			obj = append(obj, field{key: attr.key, value: convertValue(raw, attr, path)})
		}
		content = append(content, obj)
	}
	return tableData{content: content, attrs: attrs, file: path}, nil
}

func convertValue(raw any, attr attrInfo, path string) any {
	if !isMissing(raw) {
		switch attr.typ {
		case "number":
			return parseNumber(raw)
		case "string":
			return normalizeString(raw)
		default:
			return normalizeLooseValue(raw)
		}
	}
	switch attr.typ {
	case "number":
		return float64(0)
	case "string":
		return ""
	default:
		fmt.Printf("WARN unhandled empty attr in %s: name=%s key=%s type=%s\n", path, attr.name, attr.key, attr.typ)
		return ""
	}
}

func parseNumber(raw any) any {
	switch v := raw.(type) {
	case float64:
		if math.IsNaN(v) || math.IsInf(v, 0) {
			return nil
		}
		return v
	case int:
		return float64(v)
	case string:
		text := strings.TrimSpace(v)
		if text == "" {
			return float64(0)
		}
		n, err := strconv.ParseFloat(text, 64)
		if err != nil || math.IsNaN(n) || math.IsInf(n, 0) {
			return nil
		}
		return n
	default:
		text := strings.TrimSpace(fmt.Sprint(v))
		if text == "" {
			return float64(0)
		}
		n, err := strconv.ParseFloat(text, 64)
		if err != nil {
			return nil
		}
		return n
	}
}

func normalizeString(raw any) string {
	text := cellString(raw)
	if strings.TrimSpace(text) == "" {
		return ""
	}
	return text
}

func normalizeLooseValue(raw any) any {
	switch v := raw.(type) {
	case nil:
		return ""
	case float64, bool:
		return v
	default:
		return cellString(raw)
	}
}

func createItems(data tableData) (map[string]*itemInfo, []string, error) {
	items := make(map[string]*itemInfo, len(data.content))
	order := make([]string, 0, len(data.content))
	for _, row := range data.content {
		item := newItemInfo()
		for _, f := range row {
			if !shouldKeepItemField(f) {
				continue
			}
			item.set(f.key, normalizeItemValue(f.key, f.value))
		}
		name, _ := item.getString("name")
		if name == "" {
			continue
		}
		item.name = name
		item.merge = intValue(item.value("merge"))
		item.sensitive = parseCSVSet(item.stringValue("sensitiveField"))
		items[name] = item
		order = append(order, name)
	}
	return items, order, nil
}

func newItemInfo() *itemInfo {
	item := &itemInfo{index: map[string]int{}}
	item.set("modification", float64(0))
	item.set("values", json.RawMessage("[]"))
	item.set("name", "")
	item.set("id", float64(0))
	item.set("subkey", json.RawMessage("{}"))
	item.set("typeName", "")
	item.set("merge", float64(0))
	return item
}

func shouldKeepItemField(f field) bool {
	if f.key == "path" || f.key == "constNamespace" {
		return true
	}
	switch v := f.value.(type) {
	case string:
		return v != ""
	case float64:
		return true
	case nil:
		return false
	default:
		return true
	}
}

func normalizeItemValue(key string, value any) any {
	if key == "constField" || key == "htmlField" || key == "multidimensional" {
		if s, ok := value.(string); ok && strings.TrimSpace(s) != "" {
			var raw json.RawMessage
			if json.Unmarshal([]byte(s), &raw) == nil {
				return raw
			}
		}
	}
	return value
}

func (item *itemInfo) set(key string, value any) {
	if index, ok := item.index[key]; ok {
		item.fields[index].value = value
		return
	}
	item.index[key] = len(item.fields)
	item.fields = append(item.fields, field{key: key, value: value})
}

func (item *itemInfo) value(key string) any {
	if index, ok := item.index[key]; ok {
		return item.fields[index].value
	}
	return nil
}

func (item *itemInfo) stringValue(key string) string {
	value, _ := item.getString(key)
	return value
}

func (item *itemInfo) getString(key string) (string, bool) {
	value := item.value(key)
	if value == nil {
		return "", false
	}
	return cellString(value), true
}

func intValue(value any) int {
	switch v := value.(type) {
	case float64:
		return int(v)
	case string:
		n, _ := strconv.Atoi(strings.TrimSpace(v))
		return n
	default:
		n, _ := strconv.Atoi(fmt.Sprint(v))
		return n
	}
}

func parseCSVSet(value string) map[string]bool {
	out := map[string]bool{}
	for _, part := range strings.Split(value, ",") {
		part = strings.TrimSpace(part)
		if part != "" {
			out[part] = true
		}
	}
	return out
}

func marshalRows(rows []rowObject) ([]byte, error) {
	var buf bytes.Buffer
	buf.WriteString("[\n")
	for i, row := range rows {
		if i > 0 {
			buf.WriteString(",\n")
		}
		buf.WriteString("\t{\n")
		for j, f := range row {
			if j > 0 {
				buf.WriteString(",\n")
			}
			key, _ := json.Marshal(f.key)
			value, err := marshalValue(f.value)
			if err != nil {
				return nil, err
			}
			buf.WriteString("\t\t")
			buf.Write(key)
			buf.WriteByte(':')
			buf.Write(value)
		}
		buf.WriteString("\n\t}")
	}
	buf.WriteString("\n]")
	return buf.Bytes(), nil
}

func writeItems(path string, order []string, items map[string]*itemInfo) error {
	var buf bytes.Buffer
	buf.WriteString("{\n\t\"tables\":")
	tableNames := make([]string, 0, len(order))
	for _, name := range order {
		if _, ok := items[name]; ok {
			tableNames = append(tableNames, name)
		}
	}
	tableJSON, _ := json.Marshal(tableNames)
	buf.Write(tableJSON)

	for _, name := range order {
		item, ok := items[name]
		if !ok {
			continue
		}
		key, _ := json.Marshal(name)
		buf.WriteString(",\n\t")
		buf.Write(key)
		buf.WriteString(":{\n")
		for i, f := range item.fields {
			if i > 0 {
				buf.WriteString(",\n")
			}
			fieldKey, _ := json.Marshal(f.key)
			value, err := marshalValue(f.value)
			if err != nil {
				return fmt.Errorf("marshal item %s.%s: %w", name, f.key, err)
			}
			buf.WriteString("\t\t")
			buf.Write(fieldKey)
			buf.WriteByte(':')
			buf.Write(value)
		}
		buf.WriteString("\n\t}")
	}
	buf.WriteString("\n}")
	return os.WriteFile(path, buf.Bytes(), 0o644)
}

func marshalValue(value any) ([]byte, error) {
	switch v := value.(type) {
	case json.RawMessage:
		if len(v) == 0 {
			return []byte("null"), nil
		}
		return v, nil
	default:
		return json.Marshal(v)
	}
}

func readHashFile(path string) map[string]int64 {
	data, err := os.ReadFile(path)
	if err != nil {
		return map[string]int64{}
	}
	out := map[string]int64{}
	if err := json.Unmarshal(data, &out); err != nil {
		fmt.Printf("WARN ignore invalid hash file %s: %v\n", path, err)
		return map[string]int64{}
	}
	return out
}

func writeHashFile(path string, hashes map[string]int64) error {
	if err := os.MkdirAll(filepath.Dir(path), 0o755); err != nil {
		return err
	}
	data, err := json.Marshal(hashes)
	if err != nil {
		return err
	}
	return os.WriteFile(path, data, 0o644)
}

func readWorkbookRows(path string, opts options) ([][][]any, error) {
	if strings.EqualFold(filepath.Ext(path), ".xlsx") {
		sheets, err := readXLSXRowsFast(path)
		if err == nil {
			return sheets, nil
		}
		fmt.Printf("WARN fast xlsx reader fallback to node for %s: %v\n", path, err)
		return readXLSRowsWithNode(path, opts)
	}
	return readXLSRowsWithNode(path, opts)
}

func readXLSRowsWithNode(path string, opts options) ([][][]any, error) {
	script := `
const file = process.argv[1];
const modulePath = process.argv[2];
const XLSX = require(modulePath);
const wb = XLSX.readFile(file, { cellDates: false });
const out = wb.SheetNames.map((name) => XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: true }));
process.stdout.write(JSON.stringify(out));
`
	cmd := exec.Command(opts.nodePath, "-e", script, path, opts.xlsxModule)
	var stderr bytes.Buffer
	cmd.Stderr = &stderr
	output, err := cmd.Output()
	if err != nil {
		return nil, fmt.Errorf("read .xls with node failed; pass -node and -xlsx-module if node is not in PATH: %w\n%s", err, stderr.String())
	}
	var sheets [][][]any
	if err := json.Unmarshal(output, &sheets); err != nil {
		return nil, fmt.Errorf("decode node xls output: %w", err)
	}
	return sheets, nil
}

func readXLSXRowsFast(path string) ([][][]any, error) {
	reader, err := zip.OpenReader(path)
	if err != nil {
		return nil, fmt.Errorf("open xlsx: %w", err)
	}
	defer reader.Close()

	entries := make(map[string]*zip.File, len(reader.File))
	for _, entry := range reader.File {
		entries[filepath.ToSlash(entry.Name)] = entry
	}

	workbookXML, err := readZipEntry(entries, "xl/workbook.xml")
	if err != nil {
		return nil, err
	}
	relsXML, err := readZipEntry(entries, "xl/_rels/workbook.xml.rels")
	if err != nil {
		return nil, err
	}
	sheets := parseSheets(workbookXML)
	if len(sheets) == 0 {
		return nil, errors.New("workbook has no sheets")
	}
	relationships := parseRelationships(relsXML)
	var sharedStrings []string
	if _, ok := entries["xl/sharedStrings.xml"]; ok {
		sharedXML, err := readZipEntry(entries, "xl/sharedStrings.xml")
		if err != nil {
			return nil, err
		}
		sharedStrings = parseSharedStrings(sharedXML)
	}

	out := make([][][]any, 0, len(sheets))
	for _, sheet := range sheets {
		target, ok := relationships[sheet.relationshipID]
		if !ok {
			return nil, fmt.Errorf("worksheet relationship not found: %s", sheet.relationshipID)
		}
		sheetPath := normalizeZipPath("xl", target)
		entry, ok := entries[sheetPath]
		if !ok {
			return nil, fmt.Errorf("xlsx entry not found: %s", sheetPath)
		}
		rows, err := readSheetRows(entry, sheetPath, sharedStrings)
		if err != nil {
			return nil, err
		}
		out = append(out, rows)
	}
	return out, nil
}

func readSheetRows(entry *zip.File, sheetPath string, sharedStrings []string) ([][]any, error) {
	reader, err := entry.Open()
	if err != nil {
		return nil, fmt.Errorf("open worksheet entry %s: %w", sheetPath, err)
	}
	defer reader.Close()
	scanner := newRowScanner(reader)
	rows := [][]any{}
	for {
		rowXML, err := scanner.next()
		if err != nil {
			return nil, err
		}
		if rowXML == nil {
			break
		}
		row := parseRowElement(rowXML, sharedStrings)
		anyRow := make([]any, len(row))
		for i, v := range row {
			if v != "" {
				anyRow[i] = v
			}
		}
		rows = append(rows, anyRow)
	}
	return rows, nil
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
	defer reader.Close()
	return io.ReadAll(reader)
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
		sheets = append(sheets, sheetInfo{name: attrValue(tag, "name"), relationshipID: attrValue(tag, "r:id")})
		offset = start + end + 1
	}
	return sheets
}

func parseRelationships(xml []byte) map[string]string {
	relationships := map[string]string{}
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

type rowScanner struct {
	reader io.Reader
	buffer []byte
	chunk  []byte
	eof    bool
}

func newRowScanner(reader io.Reader) *rowScanner {
	return &rowScanner{reader: reader, buffer: make([]byte, 0, 2*1024*1024), chunk: make([]byte, 256*1024)}
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
		values[index] = resolveCellValue(attrValue(cellTag, "t"), rowXML[cellContentStart:cellContentEnd], sharedStrings)
		offset = cellContentEnd + len("</c>")
	}
	return values
}

func isSelfClosingTag(tag []byte) bool {
	for i := len(tag) - 2; i >= 0; i-- {
		switch tag[i] {
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
	case "str":
		if text := textInTag(content, "v"); text != "" {
			return text
		}
		return collectTextTags(content)
	default:
		return textInTag(content, "v")
	}
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
		builder.WriteString(xmlText(content[textStart : textStart+textEnd]))
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

func xmlText(raw []byte) string {
	if !bytes.Contains(raw, []byte("&")) {
		return string(raw)
	}
	return html.UnescapeString(string(raw))
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

func rowValue(row []any, index int) any {
	if index < 0 || index >= len(row) {
		return nil
	}
	return row[index]
}

func isMissing(value any) bool {
	if value == nil {
		return true
	}
	if s, ok := value.(string); ok && s == "" {
		return true
	}
	return false
}

func cellString(value any) string {
	switch v := value.(type) {
	case nil:
		return ""
	case string:
		return v
	case float64:
		if math.Trunc(v) == v {
			return strconv.FormatInt(int64(v), 10)
		}
		return strconv.FormatFloat(v, 'f', -1, 64)
	case bool:
		if v {
			return "true"
		}
		return "false"
	default:
		return fmt.Sprint(v)
	}
}

func encodeHTML(str string) string {
	if str == "" {
		return ""
	}
	s := strings.ReplaceAll(str, "&", "&amp;")
	s = strings.ReplaceAll(s, "<", "&lt;")
	s = strings.ReplaceAll(s, ">", "&gt;")
	s = strings.ReplaceAll(s, `"`, "&quot;")
	return s
}

func fileMTime(path string) int64 {
	info, err := os.Stat(path)
	if err != nil {
		return 0
	}
	return info.ModTime().UnixMilli()
}

func isExcelFile(path string) bool {
	name := filepath.Base(path)
	if strings.HasPrefix(name, "~$") {
		return false
	}
	ext := strings.ToLower(filepath.Ext(name))
	return ext == ".xls" || ext == ".xlsx"
}

func init() {
	log.SetFlags(0)
	_ = bufio.ErrInvalidUnreadByte
}
