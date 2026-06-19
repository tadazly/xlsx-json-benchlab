# 高性能打表工具实验项目

这个项目用于探索和沉淀“高性能、低内存、中文安全”的 Excel `.xlsx` 转 JSON 打表方案。它不是一个完整业务产品，而是一个实现参考：通过 Go 和 Node.js 两条路线对比不同解析策略，启发后续构建更完整的打表工具。

项目重点关注：

- 大体积 `.xlsx` 表格转换。
- 大量中文文本不乱码。
- 低内存流式处理。
- 输出 JSON 结构稳定，便于 hash / diff 校验。
- 在通用 Excel 兼容性和极致性能之间做清晰取舍。

## 输出规则

所有转换器都遵循同一套 JSON 输出契约：

- 默认读取第一个 worksheet。
- 支持通过 sheet 名称指定目标 worksheet。
- 第一行作为表头。
- 只输出第一行表头非空的列。
- 第一行表头为空的列，即使后续行有数据，也不会输出。
- 重复表头会重命名为 `name_2`、`name_3`。
- 空数据行会被跳过，判断范围只包含导出列。
- 字段顺序与表头顺序一致。
- JSON 为 compact array，一行一个 object。
- 转换结束后输出数据行数和耗时。

输出示例：

```json
[
{"字段_001":"行1列1_中文内容","字段_002":"行1列2_中文内容"},
{"字段_001":"行2列1_中文内容","字段_002":"行2列2_中文内容"}
]
```

## 实现路线

### Go 通用版

路径：

```text
cmd/excel2json-go
```

特点：

- 使用 `excelize` 流式读取 Excel。
- 使用有序 JSON 写入，避免 map key 排序。
- 兼容性更好，适合作为兜底实现。
- 性能不如专用解析器，但行为更接近通用 Excel reader。

运行：

```powershell
go run ./cmd/excel2json-go -input .\data.xlsx -output .\data.json
```

指定 sheet：

```powershell
go run ./cmd/excel2json-go -input .\data.xlsx -output .\data.json -sheet Sheet1
```

### Go fast 低内存高速版

路径：

```text
cmd/excel2json-fast
```

特点：

- 不走通用 Excel reader。
- 直接读取 `.xlsx` ZIP 内部 XML。
- 流式读取 worksheet XML，不把整张 XML 解压进内存。
- 按 `<row>...</row>` 切行，只保留当前行。
- 只解析打表需要的结构：`workbook.xml`、relationships、`sharedStrings.xml`、worksheet XML。
- 对普通大表吞吐最高，是当前推荐的高性能路线。

运行：

```powershell
go run ./cmd/excel2json-fast -input .\data.xlsx -output .\data-fast.json
```

### Node.js 低内存流式版

路径：

```text
nodejs/excel2json.js
```

特点：

- 使用 `unzipper` 读取 `.xlsx` ZIP 结构。
- 使用 `saxes` 流式解析 XML。
- 使用 `StringDecoder("utf8")` 处理 UTF-8 chunk 边界，避免中文被拆坏。
- 不使用 ExcelJS streaming reader，因为实测中文场景可能出现 `�`。
- 输出结构与 Go fast 对齐，可用于跨实现校验。

安装依赖：

```powershell
cd nodejs
npm install
```

运行：

```powershell
npm run convert -- --input ..\data.xlsx --output ..\data-node.json
```

指定 sheet：

```powershell
npm run convert -- --input ..\data.xlsx --output ..\data-node.json --sheet Sheet1
```

## 压力测试表

生成 10MB、50MB、100MB 左右的中文填充 `.xlsx`：

```powershell
go run ./cmd/generate-pressure-xlsx
```

默认输出到：

```text
..\pressure-table
```

自定义参数：

```powershell
go run ./cmd/generate-pressure-xlsx -out ..\pressure-table -targets 10,50,100 -cols 24 -cell-chars 48
```

## 校验方式

建议用 Go fast 和 Node.js 互相校验输出：

```powershell
go run ./cmd/excel2json-fast -input ..\pressure-table\pressure-100mb.xlsx -output .\100mb-fast.json
cd nodejs
npm run convert -- --input ..\..\pressure-table\pressure-100mb.xlsx --output ..\100mb-node.json
cd ..
```

比较文件大小和 hash：

```powershell
Get-Item .\100mb-fast.json, .\100mb-node.json | Select-Object Name,Length
Get-FileHash .\100mb-fast.json, .\100mb-node.json -Algorithm SHA256
```

如果大小和 SHA256 一致，说明两边输出逐字节一致。

## 设计文档

低内存 Go fast 和 Node.js 转换器的实现关键整理在：

[docs/low-memory-xlsx-json.md](docs/low-memory-xlsx-json.md)

这份文档适合直接交给后续 Agent，用来创建完整的打表工具。

## 适用边界

适合：

- 普通 `.xlsx` 二维表。
- 第一行表头。
- 大量中文文本。
- 大文件批量转 JSON。
- 需要低内存和高吞吐的打表流程。

不完整支持：

- 样式驱动的日期格式还原。
- 公式重新计算。
- 富文本格式保留。
- 合并单元格语义。
- 样式影响展示值的复杂 workbook。

推荐生产形态：

- `fast`：默认高性能路径。
- `safe`：通用 Excel 库兜底。
- `verify`：抽样或全量双实现校验。
