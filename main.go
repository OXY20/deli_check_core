package main

import (
	"flag"
	"fmt"
	"log"

	"deli_check_core/core"
	"deli_check_core/web"
)

func main() {
	var fileFlag string
	var portFlag string
	flag.StringVar(&fileFlag, "f", "", "指定单个 Excel 文件路径 (.xls 或 .xlsx)，进入 CLI 模式")
	flag.StringVar(&portFlag, "port", "8080", "Web 服务端口号")
	flag.Parse()

	// 仅当显式传 -f 时保留 CLI 模式；否则默认启动 Web 界面
	if fileFlag != "" {
		outputDir := "data/output"
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

	web.StartWebServer(portFlag)
}
