# 低内存高速 XLSX 转 JSON 实现说明

这份文档整理当前项目中两套“低内存、高吞吐、中文安全”的 `.xlsx` 转 JSON 实现，方便后续用 Agent 创建完整打表工具时复用。

当前实现：

- Go fast：`cmd/excel2json-fast/main.go`
- Node.js：`nodejs/excel2json.js`

## 输出契约

两套实现必须输出完全一致的 JSON，结构如下：

```json
[
{"字段_001":"行1列1_中文内容","字段_002":"行1列2_中文内容"},
{"字段_001":"行2列1_中文内容","字段_002":"行2列2_中文内容"}
]
```

规则：

- 默认读取第一个 worksheet。
- 支持通过 `-sheet` 或 `--sheet` 指定 sheet 名称。
- 第一行作为表头。
- 只输出第一行表头非空的列。
- 第一行表头为空的列，即使后续行有数据，也不输出。
- 重复表头按 `name_2`、`name_3` 递增重命名。
- 空数据行跳过，判断范围只包括需要导出的列。
- JSON 字段顺序必须和表头顺序一致。
- JSON 使用 compact 格式，一行一个对象。
- 输出必须是 UTF-8。
- 中文文本不能出现 `�` replacement character。
- 对同一个受支持的输入，Go fast 和 Node.js 输出应当逐字节一致。
- 输入是目录时，需要递归转换所有 `.xlsx` 文件。
- 目录输出默认保留相对路径。
- 扁平化输出时，所有 JSON 写入同一目录；同名文件使用 `_2`、`_3` 后缀避免覆盖。
- 批处理需要并发执行，默认并发数按 CPU 逻辑核心数，允许用户指定最高并发数。

## XLSX 解析模型

`.xlsx` 本质是 ZIP 包。当前高速实现只读取这些文件：

- `xl/workbook.xml`：sheet 名称和 relationship id。
- `xl/_rels/workbook.xml.rels`：relationship id 到 worksheet XML 路径的映射。
- `xl/sharedStrings.xml`：可选的共享字符串表。
- `xl/worksheets/sheetN.xml`：真正的行列数据。

实现目标不是完整 Excel 兼容，而是面向“普通大表转 JSON”的高性能路径。

## Go Fast 实现关键

文件：

```text
cmd/excel2json-fast/main.go
```

核心策略：

- 使用标准库 `archive/zip` 打开 `.xlsx`。
- 只把很小的元数据 XML 读进内存。
- worksheet XML 不整体读入内存。
- 通过 `zip.File.Open()` 流式读取 worksheet entry。
- 自己实现 `rowScanner`，按 `</row>` 边界切出一行。
- 只解析当前行里的 `<c>`、`<v>`、`<is>`、`<t>`。
- 只保留当前行、小扫描缓冲、输出缓冲、shared strings。
- 使用 `bufio.Writer` 批量写 JSON。
- 使用 `strconv.AppendQuote` 生成 JSON 字符串，避免 `encoding/json` 的 map 排序和反射开销。

关键函数：

- `convertFast`：入口流程，打开 ZIP，定位 sheet，准备 writer。
- `runConversions`：区分单文件和目录输入，调度并发批处理。
- `collectDirectoryTasks`：递归收集 `.xlsx` 文件并生成输出路径。
- `optimalConcurrency`：根据 CPU 逻辑核心数和用户上限计算并发数。
- `parseSheets`：从 `workbook.xml` 读取 sheet 名称和 `r:id`。
- `parseRelationships`：从 `workbook.xml.rels` 建立 id 到 target 的映射。
- `parseSharedStrings`：读取 `sharedStrings.xml`。
- `newRowScanner` / `rowScanner.next`：流式读取 worksheet XML，返回完整 `<row>...</row>`。
- `parseRowElement`：去掉外层 `<row>`。
- `parseRow`：解析当前行中的 cell。
- `resolveCellValue`：处理 `inlineStr`、`t="s"` shared string、普通 `<v>`。
- `selectHeaderColumns`：处理表头过滤、重复表头重命名。
- `writeJSONObject`：按固定字段顺序写一行 compact JSON object。

为什么快：

- 不走 `excelize.Rows()` 的通用 Excel 解析链路。
- 不用 `encoding/xml` 做通用 XML 解析。
- 不为每行创建 `map[string]string`。
- 不让 `encoding/json` 对 map key 排序。
- 每行只构造一个 JSON buffer，然后写入。
- worksheet XML 流式处理，不会因为解压 XML 体积放大而爆内存。

内存注意：

- 当前 Go fast 已经避免整张 worksheet XML 入内存。
- `sharedStrings.xml` 仍会读成 `[]string`。
- 如果遇到巨大 shared strings 表，可以后续改成 shared string 临时文件索引或 mmap/offset 索引。

## Node.js 实现关键

文件：

```text
nodejs/excel2json.js
```

依赖：

```json
{
  "unzipper": "...",
  "saxes": "..."
}
```

核心策略：

- 使用 `unzipper.Open.file(input)` 读取 `.xlsx` ZIP 结构。
- 使用 `saxes` 做 XML 流式解析。
- 使用 `StringDecoder("utf8")` 处理 XML chunk。
- 先解析 workbook 和 relationships，定位目标 worksheet。
- 如存在 `sharedStrings.xml`，先解析 shared strings。
- worksheet 通过 SAX 事件逐行读取。
- JSON 通过 `fs.createWriteStream` 流式写入。
- 写入使用队列保持顺序，并处理 backpressure。

关键函数：

- `convert`：入口流程，打开 ZIP，定位 worksheet，读取 shared strings。
- `runConversions`：区分单文件和目录输入，调度并发批处理。
- `collectDirectoryTasks`：递归收集 `.xlsx` 文件并生成输出路径。
- `optimalConcurrency`：根据 CPU 逻辑核心数和用户上限计算并发数。
- `readWorkbook`：解析 sheet 元数据和 relationships。
- `readSharedStrings`：构建共享字符串表。
- `streamWorksheetToJSON`：流式读取 worksheet 并写 JSON。
- `parseXMLStream`：把 stream chunk 通过 `StringDecoder` 送入 `SaxesParser`。
- `resolveCellValue`：处理 shared string、inline string、普通值。
- `selectHeaderColumns`：处理表头规则。
- `buildJSONObject`：生成和 Go fast 一致的一行 JSON object。
- `writeChunk`：写入文件并处理 backpressure。

中文安全关键点：

- 不要使用 ExcelJS streaming reader 处理这类中文大表。
- 之前实测 ExcelJS streaming reader 会在部分中文文本中产生 `�`。
- Node.js 版本必须用 `StringDecoder("utf8")` 处理 XML chunk，否则 UTF-8 多字节字符跨 chunk 时可能损坏。

## 验证流程

生成压力表：

```powershell
go run ./cmd/generate-pressure-xlsx
```

运行 Go fast：

```powershell
go run ./cmd/excel2json-fast -input ..\pressure-table\pressure-100mb.xlsx -output .\100mb-fast.json
```

运行 Node.js：

```powershell
cd nodejs
npm run convert -- --input ..\..\pressure-table\pressure-100mb.xlsx --output ..\100mb-node.json
cd ..
```

比较输出：

```powershell
Get-Item .\100mb-fast.json, .\100mb-node.json | Select-Object Name,Length
Get-FileHash .\100mb-fast.json, .\100mb-node.json -Algorithm SHA256
```

验收标准：

- 行数一致。
- 文件大小一致。
- SHA256 一致。
- 抽样检查中文没有乱码。
- 输出中不出现 `�`。

## 当前基准参考

当前压力表上的参考结果：

```text
Go fast streaming:
10MB   ~278ms
50MB   ~1.352s
100MB  ~2.695s

Node custom streaming:
10MB   ~533ms
50MB   ~2.49s
100MB  ~4.88s
```

这些数字和机器、磁盘缓存、输入结构有关，后续应以目标生产数据重新压测。

## 已知边界

适合：

- `.xlsx`
- 普通二维表
- 第一行表头
- 大量中文文本
- `inlineStr`
- shared strings
- 简单 `<v>` 数字或字符串值
- 大文件批量转 JSON

不完整支持：

- 样式驱动的日期格式还原。
- 公式重新计算。
- 富文本格式保留。
- 合并单元格语义。
- 样式影响展示值的场景。
- 一次导出多个 worksheet。
- 超大 `sharedStrings.xml` 的极限低内存场景。

## 给 Agent 的实现提示词

可以把下面这段给后续 Agent：

```text
请实现一个低内存高速 .xlsx 转 JSON 命令行工具。要求：

1. 输入 .xlsx，输出 .json。
2. 默认读取第一个 worksheet，支持按名称指定 sheet。
3. 第一行作为表头，只输出第一行非空表头列。
4. 重复表头按 name_2、name_3 重命名。
5. 空数据行跳过，只判断导出列。
6. 输出 compact JSON array，一行一个 object，字段顺序和表头一致。
7. 必须保留中文，UTF-8 chunk 边界不能损坏文本。
8. 不能把大型 worksheet XML 整体读入内存。
9. 直接解析 .xlsx ZIP 中的 workbook.xml、workbook.xml.rels、sharedStrings.xml、worksheet XML。
10. 支持 inlineStr、t="s" shared string、普通 <v>。
11. 提供 10MB、50MB、100MB 压测和 SHA256 对齐验证。
12. 明确说明不完整支持日期格式、公式、合并单元格、样式语义。
13. 支持目录输入递归转换。
14. 支持保留相对路径输出和扁平化输出。
15. 支持按 CPU 逻辑核心数并发，并允许指定最高并发数。
16. 输出每个文件的转换耗时和总耗时。
```

## 推荐生产形态

建议完整工具保留三种模式：

- `fast`：默认路径，低内存、高吞吐，处理普通大表。
- `safe`：通用 Excel 库兜底，处理复杂 workbook。
- `verify`：抽样或全量双实现转换，对比 SHA256 或行级 checksum。

对于“中文普通大表批量打 JSON”的场景，优先使用 `fast`。
