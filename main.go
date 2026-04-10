package main

import (
	"flag"
	"fmt"
	"log"

	"deli_check_core/core"
)

func main() {
	var fileFlag string
	flag.StringVar(&fileFlag, "f", "", "指定单个 Excel 文件路径 (.xls 或 .xlsx)")
	flag.Parse()

	outputDir := "data/output"

	if fileFlag != "" {
		fmt.Println("指定文件模式")
		fmt.Printf("输入文件: %s\n", fileFlag)
		fmt.Printf("输出目录: %s\n", outputDir)
		fmt.Println()

		result, err := core.ProcessSingleFile(fileFlag, outputDir)
		if err != nil {
			log.Fatalf("处理失败: %v", err)
		}

		fmt.Println()
		fmt.Println("处理完成")
		fmt.Printf("总记录数:   %d\n", result.TotalRecords)
		fmt.Printf("员工人数:   %d\n", result.EmployeeCount)
		fmt.Printf("生成时间:   %s\n", result.GeneratedAt)
		return
	}

	inputDir := "data/origin"

	fmt.Println("目录扫描模式")
	fmt.Printf("输入目录: %s\n", inputDir)
	fmt.Printf("输出目录: %s\n", outputDir)
	fmt.Println()

	result, err := core.Compose(inputDir, outputDir)
	if err != nil {
		log.Fatalf("处理失败: %v", err)
	}

	fmt.Println()
	fmt.Println("处理完成")
	fmt.Printf("处理文件数: %d\n", result.TotalFiles)
	fmt.Printf("总记录数:   %d\n", result.TotalRecords)
	fmt.Printf("员工人数:   %d\n", result.EmployeeCount)
	fmt.Printf("生成时间:   %s\n", result.GeneratedAt)
}
