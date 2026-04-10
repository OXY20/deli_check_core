package main

import (
	"fmt"
	"log"

	"github.com/extrame/xls"
)

func main() {
	files := []string{
		"data/origin/12_(1月)员工刷卡记录表.xls",
		"data/origin/12_(2月)员工刷卡记录表.xls",
		"data/origin/12_(3月)员工刷卡记录表.xls",
	}

	for _, file := range files {
		fmt.Printf("\n========== 文件: %s ==========\n", file)
		xf, err := xls.Open(file, "utf-8")
		if err != nil {
			log.Printf("打开文件失败: %v", err)
			continue
		}

		sheet := xf.GetSheet(0)
		if sheet == nil {
			log.Printf("无法获取Sheet")
			continue
		}

		fmt.Printf("总行数: %d\n", sheet.MaxRow)

		for row := 0; row <= int(sheet.MaxRow); row++ {
			rowData := sheet.Row(row)
			if rowData == nil {
				continue
			}

			hasContent := false
			cells := ""
			for col := 0; col < 50; col++ {
				cell := rowData.Col(col)
				if cell != "" {
					hasContent = true
					cells += fmt.Sprintf("[%d:%s] ", col, cell)
				}
			}
			if hasContent {
				fmt.Printf("行%-3d: %s\n", row, cells)
			}
		}
	}
}
