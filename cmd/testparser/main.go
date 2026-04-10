package main

import (
	"encoding/json"
	"fmt"
	"log"

	"deli_check_core/tools"
)

func main() {
	records, err := tools.ProcessExcel("data/origin/12_(1月)员工刷卡记录表.xls")
	if err != nil {
		log.Fatalf("解析失败: %v", err)
	}

	fmt.Printf("========== 1. 总记录数 ==========\n")
	fmt.Printf("Total records: %d\n\n", len(records))

	fmt.Printf("========== 2. 前10条记录 ==========\n")
	limit := 10
	if len(records) < limit {
		limit = len(records)
	}
	for i := 0; i < limit; i++ {
		b, _ := json.MarshalIndent(records[i], "", "  ")
		fmt.Println(string(b))
	}
	fmt.Println()

	fmt.Printf("========== 3. 林如海的记录 ==========\n")
	printEmployeeRecords(records, "林如海")

	fmt.Printf("========== 4. 孙建阳的记录 ==========\n")
	printEmployeeRecords(records, "孙建阳")

	fmt.Printf("========== 5. 吴文平的记录（有多行数据） ==========\n")
	printEmployeeRecords(records, "吴文平")

	fmt.Printf("========== 6. 完整性检查 ==========\n")
	checkIntegrity(records)
}

func printEmployeeRecords(records []tools.AttendanceRecord, name string) {
	count := 0
	for _, r := range records {
		if r.EmployeeName == name {
			b, _ := json.MarshalIndent(r, "", "  ")
			fmt.Println(string(b))
			count++
		}
	}
	fmt.Printf("【%s】共 %d 条记录\n\n", name, count)
}

func checkIntegrity(records []tools.AttendanceRecord) {
	empMap := make(map[string]int)
	emptyName := 0
	emptyID := 0
	emptyTime := 0
	for _, r := range records {
		key := fmt.Sprintf("%s|%s", r.EmployeeID, r.EmployeeName)
		empMap[key]++
		if r.EmployeeName == "" {
			emptyName++
		}
		if r.EmployeeID == "" {
			emptyID++
		}
		if r.Time == "" {
			emptyTime++
		}
	}

	fmt.Printf("发现员工数量: %d\n", len(empMap))
	if emptyName > 0 {
		fmt.Printf("警告: 发现 %d 条姓名为空的记录\n", emptyName)
	}
	if emptyID > 0 {
		fmt.Printf("警告: 发现 %d 条工号为空的记录\n", emptyID)
	}
	if emptyTime > 0 {
		fmt.Printf("警告: 发现 %d 条时间为空的记录\n", emptyTime)
	}
	if emptyName == 0 && emptyID == 0 && emptyTime == 0 {
		fmt.Println("未发现明显异常（所有记录的姓名、工号、时间均非空）")
	}

	fmt.Println("\n员工列表及记录数 (部分):")
	for emp, cnt := range empMap {
		if cnt > 0 {
			fmt.Printf("  %s => %d 条\n", emp, cnt)
		}
	}
}
