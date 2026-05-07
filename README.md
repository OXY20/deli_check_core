# deli_check_core

**版本：v1.1.0**

Deli 打卡机考勤数据导出处理工具。支持从 `.xls` 和 `.xlsx` 格式的员工刷卡记录表中自动提取员工信息、打卡日期和具体时间，并输出为结构化的 JSON 文件。同时提供 Web 可视化界面，支持在线查看打卡记录、用餐统计及导出 Excel。

## 功能特性

- **自动识别员工块**：动态扫描「工号」「姓名」「部门」信息行，无需硬编码坐标。
- **支持跨行数据**：同一员工打卡数据可能分布在连续多行，自动合并解析。
- **兼容双格式**：同时支持旧版 `.xls`（`extrame/xls`）和新版 `.xlsx`（`excelize/v2`）。
- **修复 xls 库 Bug**：原库 `row.LastCol()` 在某些行返回 0 导致数据大量丢失，已改为固定列扫描（0~49）解决。
- **支持任意日期表头**：日期表头不一定从 1 开始（支持如 `15,16,17...` 的半月考勤），单天考勤（仅一个日期数字）也能正确识别。
- **单文件 / 目录扫描两种模式**：
  - 不传参数时，默认扫描 `data/origin/` 目录下所有 `.xls` 文件并合并输出。
  - 使用 `-f` 参数可指定单个 `.xls` 或 `.xlsx` 文件进行处理。
- **空记录告警**：员工无任何打卡记录时会在日志中输出提示，但输出中仍保留空记录结构。
- **Web 可视化界面**（`deli_check_web` 分支）：
  - 浏览器上传考勤文件，在线查看打卡记录。
  - 支持多表信息合并模式：同时上传多个考勤表，自定义地点名称，保留全部原始记录并汇总统计。
  - 按人员显示打卡详情，支持用餐统计（早餐 / 午餐 / 晚餐）。
  - 一键导出用餐数据为 `.xlsx` 表格。
  - 支持导出人员明细数据（多表模式下含地点字段）。

## 快速开始

### 编译

```bash
go build -o deli_check_core.exe .
```

### 使用方式

#### 1. 目录扫描模式（默认）

扫描 `data/origin/` 下所有 `.xls` 文件，结果写入 `data/output/`：

```bash
go run .
# 或
.\deli_check_core.exe
```

#### 2. 单文件模式

使用 `-f` 指定任意路径的 `.xls` 或 `.xlsx`：

```bash
go run . -f data\origin\12_(1月)员工刷卡记录表.xls
# 或
.\deli_check_core.exe -f C:\Users\ERSHI\Desktop\考勤表.xlsx
```

### 输出文件

运行后会在 `data/output/` 生成以下 JSON：

- `records.json`：处理结果汇总对象，内部包含明细记录数组 `records`（AttendanceRecord 数组）以及 `total_files`、`total_records`、`employee_count`、`generated_at` 等统计字段。
- `summary.json`：按员工汇总的统计（工号、姓名、部门、记录数）。

## 数据格式

### AttendanceRecord（单条考勤记录）

```json
{
  "employee_id": "1",
  "employee_name": "林如海",
  "department": "公司",
  "date": "2026-01-08",
  "time": "08:15"
}
```

> **注意**：工具不对「上班 / 下班」做区分，单元格中的每一个时间都会生成一条独立记录。

## 项目结构

```
deli_check_core/
├── main.go              # CLI 入口，支持 -f 参数
├── core/
│   └── compose.go       # 编排逻辑：聚合多文件 / 单文件，排序并写 JSON
├── tools/
│   ├── excel.go         # 核心解析器（统一矩阵 + 员工块识别）
│   ├── excel_test.go    # 自动化测试（36 个用例，覆盖解析全流程）
│   └── xlsx.go          # .xlsx 读取适配器（复用同一套解析逻辑）
├── cmd/                 # 开发调试工具（testparser / runall / analyze）
├── docs/                # 文档目录（待补充）
├── data/
│   ├── origin/          # 原始考勤文件（Git-ignored）
│   └── output/          # 生成的 JSON（Git-ignored）
├── go.mod
├── go.sum
├── .gitignore
└── README.md
```

## 关键实现说明

### 1. 动态识别逻辑

- **员工信息行**：遍历行内单元格，命中 `工号：`、`姓名：`、`部门：` 后，分别读取其右侧单元格的值。
- **日期表头行**：扫描行中连续递增的整数序列，支持从非 1 开始的序列（如 `15,16,17...`）；单天考勤仅含一个整数（如 `31`）也正确识别；行中若存在任何非数字内容则排除（防数据行误判）。
- **数据行**：日期表头后的非空行均视为该员工的数据行，按列偏移映射到具体日期。

### 2. 时间拆分规则

单个单元格可能出现多个打卡时间，以下分隔符均支持拆分：

- 空格、Tab
- 换行符（`\n` / `\r`）
- 英文逗号 `,` / 中文逗号 `，`
- 英文分号 `;` / 中文分号 `；`

## 测试

运行自动化测试：

```bash
go test ./tools/ -v
```

测试覆盖 `detectDayHeaderRow`、`findBaseDate`、`splitTimes`、`scanEmployeeInfo`、`isBlankRow`、`buildRecords`、`parseMatrix` 等核心函数，共 36 个用例。

## 依赖

```
github.com/extrame/xls     v0.0.1   # .xls 读取
github.com/xuri/excelize/v2 v2.10.1 # .xlsx 读取
```

## 注意事项

- `data/` 目录已加入 `.gitignore`，真实生产数据和生成结果不会被提交到 Git。
- 在 Windows PowerShell 中运行若文件路径含中文括号，建议用单引号包裹路径：
  ```powershell
  .\deli_check_core.exe -f 'data\origin\12_(1月)员工刷卡记录表.xls'
  ```

---

**License**：MIT（如有需要可后续补充）
