package tools

import (
	"strings"
	"testing"
	"time"
)

// ==================== detectDayHeaderRow 测试 ====================

func TestDetectDayHeaderRow_Nil(t *testing.T) {
	startCol, dayCount, ok := detectDayHeaderRow(nil)
	if ok {
		t.Error("nil 行不应识别为日期表头")
	}
	_ = startCol
	_ = dayCount
}

func TestDetectDayHeaderRow_Empty(t *testing.T) {
	row := make([]string, 50)
	startCol, dayCount, ok := detectDayHeaderRow(row)
	_ = startCol
	_ = dayCount
	if ok {
		t.Error("全空行不应识别为日期表头")
	}
}

func TestDetectDayHeaderRow_SingleDay(t *testing.T) {
	// 单天考勤：表头行只有一个数字，如 "31" 表示当月第31日
	row := make([]string, 50)
	row[1] = "31"
	startCol, dayCount, ok := detectDayHeaderRow(row)
	if !ok {
		t.Fatal("单天模式应正确识别")
	}
	if startCol != 1 {
		t.Errorf("startCol = %d, want 1", startCol)
	}
	if dayCount != 1 {
		t.Errorf("dayCount = %d, want 1", dayCount)
	}
}

func TestDetectDayHeaderRow_MultiDayFrom1(t *testing.T) {
	// 多天考勤：1,2,3,...,31
	row := make([]string, 50)
	for i := 0; i < 31; i++ {
		row[i+3] = strings.TrimSpace(string(rune('1' + i%10))) // simplified
	}
	// 使用实际数字构建
	row2 := make([]string, 50)
	row2[2] = "1"
	row2[3] = "2"
	row2[4] = "3"
	row2[5] = "4"
	row2[6] = "5"
	startCol, dayCount, ok := detectDayHeaderRow(row2)
	if !ok {
		t.Fatal("多天从1开始模式应正确识别")
	}
	if startCol != 2 {
		t.Errorf("startCol = %d, want 2", startCol)
	}
	if dayCount != 5 {
		t.Errorf("dayCount = %d, want 5", dayCount)
	}
}

func TestDetectDayHeaderRow_MultiDayFromNon1(t *testing.T) {
	// 部分月考勤：15,16,17,18
	row := make([]string, 50)
	row[1] = "15"
	row[2] = "16"
	row[3] = "17"
	row[4] = "18"
	startCol, dayCount, ok := detectDayHeaderRow(row)
	if !ok {
		t.Fatal("多天非1开始模式应正确识别")
	}
	if startCol != 1 {
		t.Errorf("startCol = %d, want 1", startCol)
	}
	if dayCount != 4 {
		t.Errorf("dayCount = %d, want 4", dayCount)
	}
}

func TestDetectDayHeaderRow_NonContinuous(t *testing.T) {
	// 数字不连续：1,2,5
	row := make([]string, 50)
	row[0] = "1"
	row[1] = "2"
	row[2] = "5"
	_, _, ok := detectDayHeaderRow(row)
	if ok {
		t.Error("非连续数字不应识别为日期表头")
	}
}

func TestDetectDayHeaderRow_MixedTextAndNumber(t *testing.T) {
	// 包含非数字内容的行（如时间 "07:18"），不应识别为日期表头
	row := make([]string, 50)
	row[1] = "31"
	row[2] = "备注" // 非数字内容
	_, _, ok := detectDayHeaderRow(row)
	if ok {
		t.Error("含非数字内容的行不应识别为日期表头")
	}
}

func TestDetectDayHeaderRow_TimeValue(t *testing.T) {
	// 数据行：含时间 "07:18"
	row := make([]string, 50)
	row[1] = "07:18"
	_, _, ok := detectDayHeaderRow(row)
	if ok {
		t.Error("含时间值的行不应识别为日期表头")
	}
}

func TestDetectDayHeaderRow_MultipleNumbersWithGaps(t *testing.T) {
	// 数字间有空列：1, , , 2（跳过空列继续）
	row := make([]string, 50)
	row[0] = "1"
	row[4] = "2" // 跳过中间空列
	_, _, ok := detectDayHeaderRow(row)
	if !ok {
		t.Error("含空列的数字序列应仍能识别")
	}
}

// ==================== findBaseDate 测试 ====================

func TestFindBaseDate_RangeFormat(t *testing.T) {
	data := [][]string{
		makeRow(50),
		makeRow(50),
	}
	data[0][25] = "考勤日期：2026-01-04～2026-01-31"
	date, err := findBaseDate(data)
	if err != nil {
		t.Fatalf("日期范围格式解析失败: %v", err)
	}
	expected := time.Date(2026, 1, 4, 0, 0, 0, 0, time.Local)
	if !date.Equal(expected) {
		t.Errorf("日期 = %v, want %v", date, expected)
	}
}

func TestFindBaseDate_SingleDateFormat(t *testing.T) {
	data := [][]string{
		makeRow(50),
	}
	data[0][25] = "考勤日期：2026-03-31"
	date, err := findBaseDate(data)
	if err != nil {
		t.Fatalf("单日期格式解析失败: %v", err)
	}
	expected := time.Date(2026, 3, 31, 0, 0, 0, 0, time.Local)
	if !date.Equal(expected) {
		t.Errorf("日期 = %v, want %v", date, expected)
	}
}

func TestFindBaseDate_SameDayRange(t *testing.T) {
	// 单天以范围形式呈现：2026-03-31～2026-03-31
	data := [][]string{
		makeRow(50),
	}
	data[0][25] = "考勤日期：2026-03-31～2026-03-31"
	date, err := findBaseDate(data)
	if err != nil {
		t.Fatalf("同日起止范围格式解析失败: %v", err)
	}
	expected := time.Date(2026, 3, 31, 0, 0, 0, 0, time.Local)
	if !date.Equal(expected) {
		t.Errorf("日期 = %v, want %v", date, expected)
	}
}

func TestFindBaseDate_NotFound(t *testing.T) {
	data := [][]string{
		makeRow(50),
	}
	_, err := findBaseDate(data)
	if err == nil {
		t.Error("无日期时应返回错误")
	}
}

func TestFindBaseDate_ColonVariant(t *testing.T) {
	data := [][]string{
		makeRow(50),
	}
	// 英文冒号变体
	data[0][25] = "考勤日期:2026-04-01～2026-04-30"
	date, err := findBaseDate(data)
	if err != nil {
		t.Fatalf("英文冒号格式解析失败: %v", err)
	}
	expected := time.Date(2026, 4, 1, 0, 0, 0, 0, time.Local)
	if !date.Equal(expected) {
		t.Errorf("日期 = %v, want %v", date, expected)
	}
}

func TestFindBaseDate_TopRowBias(t *testing.T) {
	// 验证只扫描前10行
	data := make([][]string, 15)
	for i := range data {
		data[i] = makeRow(50)
	}
	// 日期在第12行（超出扫描范围）
	data[11][0] = "考勤日期：2026-05-01～2026-05-07"
	_, err := findBaseDate(data)
	if err == nil {
		t.Error("第12行的日期不应被扫描到（仅扫描前10行）")
	}
}

// ==================== splitTimes 测试 ====================

func TestSplitTimes_Single(t *testing.T) {
	result := splitTimes("07:18")
	if len(result) != 1 || result[0] != "07:18" {
		t.Errorf("单个时间拆分错误: %v", result)
	}
}

func TestSplitTimes_Multiple(t *testing.T) {
	result := splitTimes("07:18\n11:38\n17:45")
	if len(result) != 3 {
		t.Fatalf("应返回3个时间, got %d: %v", len(result), result)
	}
	if result[0] != "07:18" || result[1] != "11:38" || result[2] != "17:45" {
		t.Errorf("拆分结果错误: %v", result)
	}
}

func TestSplitTimes_SpaceSeparated(t *testing.T) {
	result := splitTimes("07:18 11:38")
	if len(result) != 2 {
		t.Errorf("空格分隔应返回2个时间: %v", result)
	}
}

func TestSplitTimes_CommaSeparated(t *testing.T) {
	result := splitTimes("07:18,11:38,17:45")
	if len(result) != 3 {
		t.Errorf("逗号分隔应返回3个时间: %v", result)
	}
}

func TestSplitTimes_ChineseComma(t *testing.T) {
	result := splitTimes("07:18，11:38，17:45")
	if len(result) != 3 {
		t.Errorf("中文逗号分隔应返回3个时间: %v", result)
	}
}

func TestSplitTimes_Empty(t *testing.T) {
	result := splitTimes("")
	if len(result) != 0 {
		t.Errorf("空字符串应返回空: %v", result)
	}
}

// ==================== scanEmployeeInfo 测试 ====================

func TestScanEmployeeInfo_Normal(t *testing.T) {
	row := make([]string, 50)
	row[4] = "工号："
	row[5] = "1"
	row[10] = "姓名："
	row[11] = "林如海"
	row[22] = "部门："
	row[23] = "公司"
	id, name, dept, ok := scanEmployeeInfo(row)
	if !ok {
		t.Fatal("应正确识别员工信息")
	}
	if id != "1" || name != "林如海" || dept != "公司" {
		t.Errorf("id=%s name=%s dept=%s", id, name, dept)
	}
}

func TestScanEmployeeInfo_NoID(t *testing.T) {
	row := make([]string, 50)
	row[10] = "姓名："
	row[11] = "林如海"
	_, _, _, ok := scanEmployeeInfo(row)
	if ok {
		t.Error("工号为必填项，缺失时不应通过")
	}
}

func TestScanEmployeeInfo_MissingName(t *testing.T) {
	row := make([]string, 50)
	row[4] = "工号："
	row[5] = "1"
	row[22] = "部门："
	row[23] = "公司"
	id, _, _, ok := scanEmployeeInfo(row)
	if !ok {
		t.Fatal("姓名缺失但工号存在，仍应识别")
	}
	if id != "1" {
		t.Errorf("id = %s, want 1", id)
	}
}

func TestScanEmployeeInfo_NilRow(t *testing.T) {
	_, _, _, ok := scanEmployeeInfo(nil)
	if ok {
		t.Error("nil行不应识别为员工信息")
	}
}

// ==================== isBlankRow 测试 ====================

func TestIsBlankRow_Empty(t *testing.T) {
	row := make([]string, 50)
	if !isBlankRow(row) {
		t.Error("全空行应识别为空白")
	}
}

func TestIsBlankRow_WithContent(t *testing.T) {
	row := make([]string, 50)
	row[10] = "07:18"
	if isBlankRow(row) {
		t.Error("有内容行不应识别为空白")
	}
}

func TestIsBlankRow_Nil(t *testing.T) {
	if !isBlankRow(nil) {
		t.Error("nil行应识别为空白")
	}
}

// ==================== buildRecords 测试 ====================

func TestBuildRecords_SingleDay(t *testing.T) {
	baseDate := time.Date(2026, 3, 31, 0, 0, 0, 0, time.Local)
	block := &employeeBlock{
		id:             "1",
		name:           "测试",
		dept:           "测试部门",
		dayHeaderFound: true,
		dayStartCol:    1,
		dayCount:       1,
		dataRows: [][]string{
			func() []string { r := make([]string, 50); r[1] = "07:18\n11:38"; return r }(),
			func() []string { r := make([]string, 50); r[1] = "17:45"; return r }(),
		},
	}
	records := block.buildRecords(baseDate)
	if len(records) != 3 {
		t.Fatalf("应有3条记录, got %d", len(records))
	}
	for _, r := range records {
		if r.Date != "2026-03-31" {
			t.Errorf("日期应为2026-03-31, got %s", r.Date)
		}
	}
}

func TestBuildRecords_MultiDay(t *testing.T) {
	baseDate := time.Date(2026, 1, 4, 0, 0, 0, 0, time.Local)
	block := &employeeBlock{
		id:             "25",
		name:           "孙建阳",
		dept:           "公司",
		dayHeaderFound: true,
		dayStartCol:    1,
		dayCount:       3, // 3天: 1月4日~6日
		dataRows: [][]string{
			func() []string {
				r := make([]string, 50)
				r[1] = "07:52" // 第1天
				r[2] = "08:10" // 第2天
				r[3] = "07:55" // 第3天
				return r
			}(),
		},
	}
	records := block.buildRecords(baseDate)
	if len(records) != 3 {
		t.Fatalf("应有3条记录, got %d", len(records))
	}
	expectedDates := []string{"2026-01-04", "2026-01-05", "2026-01-06"}
	for i, r := range records {
		if r.Date != expectedDates[i] {
			t.Errorf("记录[%d]日期 = %s, want %s", i, r.Date, expectedDates[i])
		}
	}
}

func TestBuildRecords_NoDayHeader(t *testing.T) {
	block := &employeeBlock{
		id:             "1",
		name:           "测试",
		dayHeaderFound: false,
	}
	records := block.buildRecords(time.Now())
	if len(records) != 0 {
		t.Errorf("无日期表头时应返回空, got %d", len(records))
	}
}

func TestBuildRecords_EmptyDataRows(t *testing.T) {
	block := &employeeBlock{
		id:             "1",
		dayHeaderFound: true,
		dayStartCol:    1,
		dayCount:       1,
		dataRows:       nil,
	}
	records := block.buildRecords(time.Now())
	if len(records) != 0 {
		t.Errorf("无数据行时应返回空, got %d", len(records))
	}
}

// ==================== parseMatrix 集成测试 ====================

func TestParseMatrix_SingleDayIntegration(t *testing.T) {
	// 模拟单天考勤数据结构
	data := [][]string{
		makeRow(50), // Row 0: 标题
		makeRow(50), // Row 1:
		func() []string { // Row 2: 考勤日期
			r := makeRow(50)
			r[25] = "考勤日期：2026-03-31"
			return r
		}(),
		makeRow(50), // Row 3:
		func() []string { // Row 4: 员工1信息
			r := makeRow(50)
			r[4] = "工号："
			r[5] = "1"
			r[10] = "姓名："
			r[11] = "张三"
			return r
		}(),
		func() []string { // Row 5: 日期表头（单天31）
			r := makeRow(50)
			r[1] = "31"
			return r
		}(),
		func() []string { // Row 6: 打卡数据
			r := makeRow(50)
			r[1] = "07:18\n11:38"
			return r
		}(),
		func() []string { // Row 7: 空行
			r := makeRow(50)
			return r
		}(),
		func() []string { // Row 8: 员工2信息
			r := makeRow(50)
			r[4] = "工号："
			r[5] = "2"
			r[10] = "姓名："
			r[11] = "李四"
			return r
		}(),
		func() []string { // Row 9: 日期表头
			r := makeRow(50)
			r[1] = "31"
			return r
		}(),
		// 员工2无打卡数据，应提示但不出错
	}
	records, err := parseMatrix(data, "test.xls")
	if err != nil {
		t.Fatalf("解析失败: %v", err)
	}
	if len(records) != 2 {
		t.Fatalf("应有2条记录, got %d: %+v", len(records), records)
	}
	for _, r := range records {
		if r.Date != "2026-03-31" {
			t.Errorf("日期 = %s, want 2026-03-31", r.Date)
		}
	}
}

func TestParseMatrix_MultiDayIntegration(t *testing.T) {
	data := [][]string{
		makeRow(50),
		func() []string {
			r := makeRow(50)
			r[25] = "考勤日期：2026-01-04～2026-01-06"
			return r
		}(),
		makeRow(50),
		func() []string { // 员工信息
			r := makeRow(50)
			r[4] = "工号："
			r[5] = "25"
			r[10] = "姓名："
			r[11] = "孙建阳"
			r[22] = "部门："
			r[23] = "公司"
			return r
		}(),
		func() []string { // 日期表头 1,2,3
			r := makeRow(50)
			r[1] = "1"
			r[2] = "2"
			r[3] = "3"
			return r
		}(),
		func() []string { // 打卡数据
			r := makeRow(50)
			r[1] = "07:52"
			r[2] = "08:10"
			r[3] = "07:55"
			return r
		}(),
	}
	records, err := parseMatrix(data, "test.xls")
	if err != nil {
		t.Fatalf("解析失败: %v", err)
	}
	if len(records) != 3 {
		t.Fatalf("应有3条记录, got %d", len(records))
	}
	expectedDates := []string{"2026-01-04", "2026-01-05", "2026-01-06"}
	for i, r := range records {
		if r.Date != expectedDates[i] {
			t.Errorf("记录[%d]日期 = %s, want %s", i, r.Date, expectedDates[i])
		}
	}
}

func TestParseMatrix_PartialMonthIntegration(t *testing.T) {
	// 半月考勤：从15日开始，日期表头为 15,16,17
	data := [][]string{
		makeRow(50),
		func() []string {
			r := makeRow(50)
			r[25] = "考勤日期：2026-03-15～2026-03-17"
			return r
		}(),
		makeRow(50),
		func() []string { // 员工信息
			r := makeRow(50)
			r[4] = "工号："
			r[5] = "1"
			r[10] = "姓名："
			r[11] = "测试"
			return r
		}(),
		func() []string { // 日期表头 15,16,17
			r := makeRow(50)
			r[1] = "15"
			r[2] = "16"
			r[3] = "17"
			return r
		}(),
		func() []string { // 打卡数据
			r := makeRow(50)
			r[1] = "07:30"
			r[2] = "07:35"
			r[3] = "07:40"
			return r
		}(),
	}
	records, err := parseMatrix(data, "test.xls")
	if err != nil {
		t.Fatalf("部分月刊勤解析失败: %v", err)
	}
	if len(records) != 3 {
		t.Fatalf("应有3条记录, got %d", len(records))
	}
	expectedDates := []string{"2026-03-15", "2026-03-16", "2026-03-17"}
	for i, r := range records {
		if r.Date != expectedDates[i] {
			t.Errorf("记录[%d]日期 = %s, want %s", i, r.Date, expectedDates[i])
		}
	}
}

func TestParseMatrix_NoDate(t *testing.T) {
	data := [][]string{
		makeRow(50),
		makeRow(50),
	}
	_, err := parseMatrix(data, "test.xls")
	if err == nil {
		t.Error("无日期时应返回错误")
	}
}

// ==================== 辅助函数 ====================

func makeRow(cols int) []string {
	r := make([]string, cols)
	// 确保所有元素为空字符串而非 nil
	for i := range r {
		r[i] = ""
	}
	return r
}
