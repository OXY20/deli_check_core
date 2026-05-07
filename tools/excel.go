package tools

import (
	"fmt"
	"log"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/extrame/xls"
)

// AttendanceRecord 表示清洗后的单条考勤记录
type AttendanceRecord struct {
	EmployeeID   string `json:"employee_id"`   // 工号
	EmployeeName string `json:"employee_name"` // 姓名
	Department   string `json:"department"`    // 部门
	Date         string `json:"date"`          // 日期 (YYYY-MM-DD)
	Time         string `json:"time"`          // 打卡时间
	Location     string `json:"location"`      // 地点（多表合并模式下使用）
}

// ProcessExcel 根据文件后缀自动选择 .xls 或 .xlsx 解析器
func ProcessExcel(path string) ([]AttendanceRecord, error) {
	lower := strings.ToLower(path)
	if strings.HasSuffix(lower, ".xlsx") {
		return processXlsx(path)
	}
	if strings.HasSuffix(lower, ".xls") {
		return processXls(path)
	}
	return nil, fmt.Errorf("不支持的文件格式: %s", path)
}

// processXls 读取 .xls 文件并解析
func processXls(path string) ([]AttendanceRecord, error) {
	wb, err := xls.Open(path, "utf-8")
	if err != nil {
		return nil, fmt.Errorf("打开 xls 文件失败: %w", err)
	}
	sheet := wb.GetSheet(0)
	if sheet == nil {
		return nil, fmt.Errorf("没有工作表")
	}
	maxRow := int(sheet.MaxRow)
	data := make([][]string, maxRow+1)
	for i := 0; i <= maxRow; i++ {
		row := sheet.Row(i)
		line := make([]string, 50)
		for j := 0; j < 50; j++ {
			line[j] = row.Col(j)
		}
		data[i] = line
	}
	return parseMatrix(data, path)
}

// parseMatrix 是统一的核心解析逻辑，接收二维字符串矩阵
func parseMatrix(data [][]string, path string) ([]AttendanceRecord, error) {
	baseDate, err := findBaseDate(data)
	if err != nil {
		return nil, fmt.Errorf("无法确定考勤基准日期: %w", err)
	}

	var records []AttendanceRecord
	var current *employeeBlock
	maxRow := len(data)

	for i := 0; i < maxRow; i++ {
		row := data[i]

		// 1) 尝试识别新的员工信息行
		if empID, empName, dept, ok := scanEmployeeInfo(row); ok {
			if current != nil {
				recs := current.buildRecords(baseDate)
				if len(recs) == 0 {
					log.Printf("[提示] 员工 %s(%s) 在 %s 中无任何打卡记录", current.name, current.id, path)
				}
				records = append(records, recs...)
			}
			current = &employeeBlock{
				id:   empID,
				name: empName,
				dept: dept,
			}
			continue
		}

		if current == nil {
			continue
		}

		// 2) 尝试识别日期数字行（位于信息行之后）
		if !current.dayHeaderFound {
			if startCol, dayCount, ok := detectDayHeaderRow(row); ok {
				current.dayStartCol = startCol
				current.dayCount = dayCount
				current.dayHeaderFound = true
			}
			continue
		}

		// 3) 已找到日期表头，后续有内容的行都视为该员工的数据行
		if !isBlankRow(row) {
			current.dataRows = append(current.dataRows, row)
		}
	}

	// finalize 最后一个员工块
	if current != nil {
		recs := current.buildRecords(baseDate)
		if len(recs) == 0 {
			log.Printf("[提示] 员工 %s(%s) 在 %s 中无任何打卡记录", current.name, current.id, path)
		}
		records = append(records, recs...)
	}

	log.Printf("文件 %s 中共读取到 %d 条数据", path, len(records))
	return records, nil
}

// employeeBlock 保存单个员工的相关信息
type employeeBlock struct {
	id             string
	name           string
	dept           string
	dayHeaderFound bool
	dayStartCol    int
	dayCount       int
	dataRows       [][]string
}

func (b *employeeBlock) buildRecords(baseDate time.Time) []AttendanceRecord {
	if !b.dayHeaderFound || b.dayCount == 0 {
		return nil
	}
	var records []AttendanceRecord
	for _, row := range b.dataRows {
		for d := 0; d < b.dayCount; d++ {
			col := b.dayStartCol + d
			if col >= len(row) {
				continue
			}
			cell := strings.TrimSpace(row[col])
			if cell == "" {
				continue
			}
			recordDate := baseDate.AddDate(0, 0, d)
			for _, t := range splitTimes(cell) {
				records = append(records, AttendanceRecord{
					EmployeeID:   b.id,
					EmployeeName: b.name,
					Department:   b.dept,
					Date:         recordDate.Format("2006-01-02"),
					Time:         t,
				})
			}
		}
	}
	return records
}

func scanEmployeeInfo(row []string) (id, name, dept string, ok bool) {
	if row == nil {
		return "", "", "", false
	}
	idCol := -1
	nameCol := -1
	deptCol := -1

	for j := 0; j < 50 && j < len(row); j++ {
		switch strings.TrimSpace(row[j]) {
		case "工号：":
			idCol = j
		case "姓名：":
			nameCol = j
		case "部门：":
			deptCol = j
		}
	}

	if idCol == -1 {
		return "", "", "", false
	}
	if idCol != -1 && idCol+1 < len(row) {
		id = strings.TrimSpace(row[idCol+1])
	}
	if nameCol != -1 && nameCol+1 < len(row) {
		name = strings.TrimSpace(row[nameCol+1])
	}
	if deptCol != -1 && deptCol+1 < len(row) {
		dept = strings.TrimSpace(row[deptCol+1])
	}
	return id, name, dept, true
}

func detectDayHeaderRow(row []string) (startCol, dayCount int, ok bool) {
	if row == nil {
		return 0, 0, false
	}

	// 收集行中所有数字及其列号，同时检查是否存在非空非数字内容
	// 日期表头行特征：所有非空单元格均为纯数字，不含任何文字或时间格式
	var nums []struct{ col, val int }
	for j := 0; j < 50 && j < len(row); j++ {
		v := strings.TrimSpace(row[j])
		if v == "" {
			continue
		}
		n, err := strconv.Atoi(v)
		if err != nil {
			// 行中存在非数字内容（如时间 "07:18" 或文字），排除数据行误判
			return 0, 0, false
		}
		nums = append(nums, struct{ col, val int }{j, n})
	}

	if len(nums) == 0 {
		return 0, 0, false
	}

	// 单天模式：行中只有一个数字（如只有 31），视为仅一天的日期表头
	if len(nums) == 1 {
		return nums[0].col, 1, true
	}

	// 多天模式：数字连续递增即可（不一定从 1 开始，支持如 15,16,17... 的部分月考勤）
	startCol = nums[0].col
	dayCount = 1
	expected := nums[0].val + 1
	for i := 1; i < len(nums); i++ {
		if nums[i].val != expected {
			return 0, 0, false
		}
		dayCount++
		expected++
	}
	return startCol, dayCount, dayCount > 0
}

func isBlankRow(row []string) bool {
	if row == nil {
		return true
	}
	for j := 0; j < 50 && j < len(row); j++ {
		if strings.TrimSpace(row[j]) != "" {
			return false
		}
	}
	return true
}

func splitTimes(val string) []string {
	parts := strings.FieldsFunc(val, func(r rune) bool {
		return r == '\n' || r == '\r' || r == ' ' || r == '\t' ||
			r == ',' || r == '，' || r == ';' || r == '；'
	})
	var res []string
	for _, p := range parts {
		p = strings.TrimSpace(p)
		if p != "" {
			res = append(res, p)
		}
	}
	return res
}

// baseDateRe 匹配带范围的考勤日期，例如：考勤日期：2026-01-04～2026-01-31
var baseDateRe = regexp.MustCompile(`考勤日期[：:]\s*(\d{4})-(\d{2})-(\d{2})\s*[～~]\s*(\d{4})-(\d{2})-(\d{2})`)

// baseDateSingleRe 匹配仅一天的考勤日期，例如：考勤日期：2026-05-07
var baseDateSingleRe = regexp.MustCompile(`考勤日期[：:]\s*(\d{4})-(\d{2})-(\d{2})`)

func findBaseDate(data [][]string) (time.Time, error) {
	for i := 0; i < len(data) && i < 10; i++ {
		row := data[i]
		if row == nil {
			continue
		}
		for j := 0; j < 50 && j < len(row); j++ {
			cell := strings.TrimSpace(row[j])
			if cell == "" {
				continue
			}

			// 优先匹配日期范围格式：考勤日期：2026-01-04～2026-01-31
			if matches := baseDateRe.FindStringSubmatch(cell); len(matches) == 7 {
				y, _ := strconv.Atoi(matches[1])
				m, _ := strconv.Atoi(matches[2])
				d, _ := strconv.Atoi(matches[3])
				return time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.Local), nil
			}

			// 匹配单日期格式：考勤日期：2026-05-07（仅一天考勤）
			if matches := baseDateSingleRe.FindStringSubmatch(cell); len(matches) == 4 {
				y, _ := strconv.Atoi(matches[1])
				m, _ := strconv.Atoi(matches[2])
				d, _ := strconv.Atoi(matches[3])
				return time.Date(y, time.Month(m), d, 0, 0, 0, 0, time.Local), nil
			}
		}
	}
	return time.Time{}, fmt.Errorf("未找到考勤日期")
}
