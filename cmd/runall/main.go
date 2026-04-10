package main

import (
	"fmt"
	"log"

	"deli_check_core/tools"
)

func main() {
	files := []string{
		"data/origin/12_(1月)员工刷卡记录表.xls",
		"data/origin/12_(2月)员工刷卡记录表.xls",
		"data/origin/12_(3月)员工刷卡记录表.xls",
	}

	for _, file := range files {
		fmt.Printf("\n========================================\n")
		fmt.Printf("文件: %s\n", file)
		fmt.Printf("========================================\n")

		records, err := tools.ProcessExcel(file)
		if err != nil {
			log.Printf("解析失败: %v", err)
			continue
		}

		// 统计每个员工的记录数
		empRecords := make(map[string][]tools.AttendanceRecord)
		for _, r := range records {
			key := fmt.Sprintf("%s|%s", r.EmployeeID, r.EmployeeName)
			empRecords[key] = append(empRecords[key], r)
		}

		var withRecords, withoutRecords []string
		for key, recs := range empRecords {
			if len(recs) == 0 {
				withoutRecords = append(withoutRecords, key)
			} else {
				withRecords = append(withRecords, fmt.Sprintf("%s => %d条", key, len(recs)))
			}
		}

		fmt.Printf("总记录数: %d\n", len(records))
		fmt.Printf("有打卡记录的员工: %d 人\n", len(withRecords))
		fmt.Printf("无打卡记录的员工: %d 人\n", len(withoutRecords))

		fmt.Printf("\n--- 有记录的部分员工（前20） ---\n")
		limit := 20
		if len(withRecords) < limit {
			limit = len(withRecords)
		}
		for i := 0; i < limit; i++ {
			fmt.Println(withRecords[i])
		}

		fmt.Printf("\n--- 无记录的部分员工（前15） ---\n")
		limit2 := 15
		if len(withoutRecords) < limit2 {
			limit2 = len(withoutRecords)
		}
		for i := 0; i < limit2; i++ {
			fmt.Println(withoutRecords[i])
		}

		fmt.Printf("\n--- 示例记录（前5条） ---\n")
		for i := 0; i < len(records) && i < 5; i++ {
			r := records[i]
			fmt.Printf("[%s %s] %s %s\n", r.Date, r.Time, r.EmployeeID, r.EmployeeName)
		}
	}
}
