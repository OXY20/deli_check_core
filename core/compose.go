package core

import (
	"encoding/json"
	"fmt"
	"log"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"time"

	"deli_check_core/tools"
)

// ComposeResult 表示 core 处理后的结果
type ComposeResult struct {
	TotalFiles    int                      `json:"total_files"`
	TotalRecords  int                      `json:"total_records"`
	EmployeeCount int                      `json:"employee_count"`
	GeneratedAt   string                   `json:"generated_at"`
	Records       []tools.AttendanceRecord `json:"records"`
}

// SummaryItem 表示单个员工的汇总统计
type SummaryItem struct {
	EmployeeID   string `json:"employee_id"`
	EmployeeName string `json:"employee_name"`
	Department   string `json:"department"`
	RecordCount  int    `json:"record_count"`
}

// Compose 扫描 inputDir 下所有 .xls 文件，解析并输出到 outputDir
func Compose(inputDir, outputDir string) (*ComposeResult, error) {
	entries, err := os.ReadDir(inputDir)
	if err != nil {
		return nil, fmt.Errorf("读取输入目录失败: %w", err)
	}

	var files []string
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		lower := strings.ToLower(e.Name())
		if strings.HasSuffix(lower, ".xls") || strings.HasSuffix(lower, ".xlsx") {
			files = append(files, filepath.Join(inputDir, e.Name()))
		}
	}

	sort.Strings(files)

	var allRecords []tools.AttendanceRecord
	for _, f := range files {
		recs, err := tools.ProcessExcel(f)
		if err != nil {
			log.Printf("解析文件失败 %s: %v", f, err)
			continue
		}
		allRecords = append(allRecords, recs...)
	}

	// 按日期、时间、工号排序，便于查看
	sort.Slice(allRecords, func(i, j int) bool {
		if allRecords[i].Date != allRecords[j].Date {
			return allRecords[i].Date < allRecords[j].Date
		}
		if allRecords[i].Time != allRecords[j].Time {
			return allRecords[i].Time < allRecords[j].Time
		}
		return allRecords[i].EmployeeID < allRecords[j].EmployeeID
	})

	// 统计员工数
	empMap := make(map[string]struct{})
	for _, r := range allRecords {
		key := fmt.Sprintf("%s|%s", r.EmployeeID, r.EmployeeName)
		empMap[key] = struct{}{}
	}

	result := &ComposeResult{
		TotalFiles:    len(files),
		TotalRecords:  len(allRecords),
		EmployeeCount: len(empMap),
		GeneratedAt:   time.Now().Format("2006-01-02 15:04:05"),
		Records:       allRecords,
	}

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return nil, fmt.Errorf("创建输出目录失败: %w", err)
	}

	// 写入 records.json
	recordsPath := filepath.Join(outputDir, "records.json")
	if err := writeJSON(recordsPath, result); err != nil {
		return nil, fmt.Errorf("写入 records.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", recordsPath)

	// 写入 summary.json（员工汇总）
	summary := buildSummary(allRecords)
	summaryPath := filepath.Join(outputDir, "summary.json")
	if err := writeJSON(summaryPath, summary); err != nil {
		return nil, fmt.Errorf("写入 summary.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", summaryPath)

	return result, nil
}

// ProcessMultipleFiles 处理多个 Excel 文件，保留原始记录并附加地点信息
func ProcessMultipleFiles(files []string, locations []string, outputDir string) (*ComposeResult, error) {
	if len(files) == 0 {
		return nil, fmt.Errorf("没有需要处理的文件")
	}

	var allRecords []tools.AttendanceRecord
	for i, f := range files {
		recs, err := tools.ProcessExcel(f)
		if err != nil {
			log.Printf("解析文件失败 %s: %v", f, err)
			continue
		}
		loc := ""
		if i < len(locations) {
			loc = locations[i]
		}
		for j := range recs {
			recs[j].Location = loc
		}
		allRecords = append(allRecords, recs...)
	}

	// 按日期、时间、工号排序
	sort.Slice(allRecords, func(i, j int) bool {
		if allRecords[i].Date != allRecords[j].Date {
			return allRecords[i].Date < allRecords[j].Date
		}
		if allRecords[i].Time != allRecords[j].Time {
			return allRecords[i].Time < allRecords[j].Time
		}
		return allRecords[i].EmployeeID < allRecords[j].EmployeeID
	})

	empMap := make(map[string]struct{})
	for _, r := range allRecords {
		key := fmt.Sprintf("%s|%s", r.EmployeeID, r.EmployeeName)
		empMap[key] = struct{}{}
	}

	result := &ComposeResult{
		TotalFiles:    len(files),
		TotalRecords:  len(allRecords),
		EmployeeCount: len(empMap),
		GeneratedAt:   time.Now().Format("2006-01-02 15:04:05"),
		Records:       allRecords,
	}

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return nil, fmt.Errorf("创建输出目录失败: %w", err)
	}

	recordsPath := filepath.Join(outputDir, "records.json")
	if err := writeJSON(recordsPath, result); err != nil {
		return nil, fmt.Errorf("写入 records.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", recordsPath)

	summary := buildSummary(allRecords)
	summaryPath := filepath.Join(outputDir, "summary.json")
	if err := writeJSON(summaryPath, summary); err != nil {
		return nil, fmt.Errorf("写入 summary.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", summaryPath)

	return result, nil
}

// ProcessSingleFile 处理单个 Excel 文件（.xls 或 .xlsx）
func ProcessSingleFile(inputFile, outputDir string) (*ComposeResult, error) {
	recs, err := tools.ProcessExcel(inputFile)
	if err != nil {
		return nil, fmt.Errorf("解析文件失败 %s: %w", inputFile, err)
	}

	// 按日期、时间、工号排序
	sort.Slice(recs, func(i, j int) bool {
		if recs[i].Date != recs[j].Date {
			return recs[i].Date < recs[j].Date
		}
		if recs[i].Time != recs[j].Time {
			return recs[i].Time < recs[j].Time
		}
		return recs[i].EmployeeID < recs[j].EmployeeID
	})

	empMap := make(map[string]struct{})
	for _, r := range recs {
		key := fmt.Sprintf("%s|%s", r.EmployeeID, r.EmployeeName)
		empMap[key] = struct{}{}
	}

	result := &ComposeResult{
		TotalFiles:    1,
		TotalRecords:  len(recs),
		EmployeeCount: len(empMap),
		GeneratedAt:   time.Now().Format("2006-01-02 15:04:05"),
		Records:       recs,
	}

	if err := os.MkdirAll(outputDir, 0755); err != nil {
		return nil, fmt.Errorf("创建输出目录失败: %w", err)
	}

	recordsPath := filepath.Join(outputDir, "records.json")
	if err := writeJSON(recordsPath, result); err != nil {
		return nil, fmt.Errorf("写入 records.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", recordsPath)

	summary := buildSummary(recs)
	summaryPath := filepath.Join(outputDir, "summary.json")
	if err := writeJSON(summaryPath, summary); err != nil {
		return nil, fmt.Errorf("写入 summary.json 失败: %w", err)
	}
	log.Printf("输出文件: %s", summaryPath)

	return result, nil
}

func buildSummary(records []tools.AttendanceRecord) []SummaryItem {
	type key struct {
		id   string
		name string
		dept string
	}
	m := make(map[key]int)
	for _, r := range records {
		k := key{id: r.EmployeeID, name: r.EmployeeName, dept: r.Department}
		m[k]++
	}

	var items []SummaryItem
	for k, v := range m {
		items = append(items, SummaryItem{
			EmployeeID:   k.id,
			EmployeeName: k.name,
			Department:   k.dept,
			RecordCount:  v,
		})
	}
	sort.Slice(items, func(i, j int) bool {
		if items[i].EmployeeID != items[j].EmployeeID {
			return items[i].EmployeeID < items[j].EmployeeID
		}
		return items[i].EmployeeName < items[j].EmployeeName
	})
	return items
}

func writeJSON(path string, v any) error {
	f, err := os.Create(path)
	if err != nil {
		return err
	}
	defer f.Close()

	enc := json.NewEncoder(f)
	enc.SetIndent("", "  ")
	return enc.Encode(v)
}
