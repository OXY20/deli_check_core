package tools

import (
	"fmt"

	"github.com/xuri/excelize/v2"
)

// processXlsx 读取 .xlsx 文件并解析
func processXlsx(path string) ([]AttendanceRecord, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开 xlsx 文件失败: %w", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)
	if sheetName == "" {
		return nil, fmt.Errorf("没有工作表")
	}

	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取工作表失败: %w", err)
	}

	maxRow := len(rows)
	data := make([][]string, maxRow)
	for i := 0; i < maxRow; i++ {
		line := make([]string, 50)
		for j := 0; j < len(rows[i]) && j < 50; j++ {
			line[j] = rows[i][j]
		}
		data[i] = line
	}

	return parseMatrix(data, path)
}
