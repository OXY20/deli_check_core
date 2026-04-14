package web

import (
	"bytes"
	_ "embed"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"time"

	"deli_check_core/core"
	"deli_check_core/tools"

	"github.com/xuri/excelize/v2"
)

//go:embed static/index.html
var indexHTML []byte

//go:embed static/results.html
var resultsHTML []byte

// StartWebServer 启动内嵌 Web 服务并在完成后自动打开浏览器（端口被占用时自动 +1 探测）
func StartWebServer(port string) {
	if port == "" {
		port = "8080"
	}

	mux := http.NewServeMux()
	mux.HandleFunc("/", handleIndex)
	mux.HandleFunc("/results", handleResults)
	mux.HandleFunc("/api/upload", handleUpload)
	mux.HandleFunc("/api/export-meal", handleExportMeal)
	mux.HandleFunc("/api/health", handleHealth)

	listener, addr := findAvailableListener("127.0.0.1", port)
	server := &http.Server{
		Addr:    addr,
		Handler: mux,
	}

	log.Printf("Web 服务启动于 http://%s", addr)
	go func() {
		if err := server.Serve(listener); err != nil && err != http.ErrServerClosed {
			log.Fatalf("Web 服务启动失败: %v", err)
		}
	}()

	// 等待服务就绪后打开浏览器
	time.Sleep(300 * time.Millisecond)
	openBrowser(fmt.Sprintf("http://%s", addr))

	// 保持进程运行
	select {}
}

// findAvailableListener 尝试从 startPort 开始找到一个可用端口，返回 net.Listener 和实际地址
func findAvailableListener(host, startPort string) (net.Listener, string) {
	p, err := strconv.Atoi(startPort)
	if err != nil {
		p = 8080
	}
	for i := 0; i < 100; i++ {
		addr := fmt.Sprintf("%s:%d", host, p+i)
		ln, err := net.Listen("tcp", addr)
		if err == nil {
			return ln, addr
		}
		log.Printf("端口 %d 被占用，尝试端口 %d", p+i, p+i+1)
	}
	log.Fatalf("无法找到可用端口（已尝试 %s ~ %d）", startPort, p+99)
	return nil, ""
}

func handleIndex(w http.ResponseWriter, r *http.Request) {
	if r.URL.Path != "/" {
		http.NotFound(w, r)
		return
	}

	w.Header().Set("Content-Type", "text/html; charset=utf-8")

	// 开发模式：直接从磁盘读取文件，支持热重载
	if os.Getenv("DELI_WEB_DEV") == "true" {
		data, err := os.ReadFile("web/static/index.html")
		if err != nil {
			log.Printf("[dev] 读取文件失败: %v，使用内嵌资源", err)
			w.Write(indexHTML)
			return
		}
		w.Write(data)
		return
	}

	// 生产模式：使用内嵌资源
	w.Write(indexHTML)
}

func handleResults(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "text/html; charset=utf-8")

	if os.Getenv("DELI_WEB_DEV") == "true" {
		data, err := os.ReadFile("web/static/results.html")
		if err != nil {
			log.Printf("[dev] 读取文件失败: %v，使用内嵌资源", err)
			w.Write(resultsHTML)
			return
		}
		w.Write(data)
		return
	}

	w.Write(resultsHTML)
}

func handleHealth(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "application/json")
	json.NewEncoder(w).Encode(map[string]string{"status": "ok"})
}

type exportMealRequest struct {
	Records []tools.AttendanceRecord `json:"records"`
}

type mealStat struct {
	ID        string
	Name      string
	Breakfast int
	Lunch     int
	Dinner    int
	Total     int
}

func handleExportMeal(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method Not Allowed", http.StatusMethodNotAllowed)
		return
	}

	var req exportMealRequest
	if err := json.NewDecoder(r.Body).Decode(&req); err != nil {
		http.Error(w, fmt.Sprintf("解析请求失败: %v", err), http.StatusBadRequest)
		return
	}

	// 按员工分组统计用餐次数（按天去重）
	type dayMealKey struct {
		empKey string
		date   string
		meal   string
	}
	seen := make(map[dayMealKey]struct{})
	empMap := make(map[string]*mealStat)
	for _, rec := range req.Records {
		key := rec.EmployeeID + "|" + rec.EmployeeName
		if _, ok := empMap[key]; !ok {
			empMap[key] = &mealStat{
				ID:   rec.EmployeeID,
				Name: rec.EmployeeName,
			}
		}
		mealType := getMealType(rec.Time)
		if mealType == "" {
			continue
		}
		dm := dayMealKey{empKey: key, date: rec.Date, meal: mealType}
		if _, ok := seen[dm]; ok {
			continue
		}
		seen[dm] = struct{}{}
		switch mealType {
		case "breakfast":
			empMap[key].Breakfast++
		case "lunch":
			empMap[key].Lunch++
		case "dinner":
			empMap[key].Dinner++
		}
	}

	var stats []mealStat
	for _, s := range empMap {
		s.Total = s.Breakfast + s.Lunch + s.Dinner
		stats = append(stats, *s)
	}

	// 按工号数字排序
	sort.Slice(stats, func(i, j int) bool {
		a, errA := strconv.Atoi(stats[i].ID)
		b, errB := strconv.Atoi(stats[j].ID)
		if errA == nil && errB == nil {
			return a < b
		}
		return stats[i].ID < stats[j].ID
	})

	// 生成 xlsx
	f := excelize.NewFile()
	sheet := "用餐统计"
	f.SetSheetName("Sheet1", sheet)

	// 表头
	headers := []string{"工号", "姓名", "早餐", "午餐", "晚餐", "总计"}
	for i, h := range headers {
		cell, _ := excelize.CoordinatesToCellName(i+1, 1)
		f.SetCellValue(sheet, cell, h)
	}

	// 数据
	for i, s := range stats {
		row := i + 2
		f.SetCellValue(sheet, fmt.Sprintf("A%d", row), s.ID)
		f.SetCellValue(sheet, fmt.Sprintf("B%d", row), s.Name)
		f.SetCellValue(sheet, fmt.Sprintf("C%d", row), s.Breakfast)
		f.SetCellValue(sheet, fmt.Sprintf("D%d", row), s.Lunch)
		f.SetCellValue(sheet, fmt.Sprintf("E%d", row), s.Dinner)
		f.SetCellValue(sheet, fmt.Sprintf("F%d", row), s.Total)
	}

	// 写入缓冲区
	var buf bytes.Buffer
	if err := f.Write(&buf); err != nil {
		http.Error(w, fmt.Sprintf("生成 Excel 失败: %v", err), http.StatusInternalServerError)
		return
	}

	now := time.Now().Format("20060102-150405")
	filename := fmt.Sprintf("用餐统计 %s.xlsx", now)

	w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=\"%s\"; filename*=UTF-8''%s", filename, url.PathEscape(filename)))
	w.Write(buf.Bytes())
}

func getMealType(timeStr string) string {
	parts := strings.Split(timeStr, ":")
	if len(parts) < 2 {
		return ""
	}
	hour, _ := strconv.Atoi(parts[0])
	minute, _ := strconv.Atoi(parts[1])
	minutes := hour*60 + minute

	if minutes >= 4*60 && minutes <= 10*60+30 {
		return "breakfast"
	}
	if minutes >= 10*60+31 && minutes <= 15*60 {
		return "lunch"
	}
	if minutes >= 16*60 && minutes <= 21*60 {
		return "dinner"
	}
	return ""
}

func handleUpload(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "Method Not Allowed", http.StatusMethodNotAllowed)
		return
	}

	file, header, err := r.FormFile("file")
	if err != nil {
		http.Error(w, fmt.Sprintf("读取文件失败: %v", err), http.StatusBadRequest)
		return
	}
	defer file.Close()

	lower := strings.ToLower(header.Filename)
	if !strings.HasSuffix(lower, ".xls") && !strings.HasSuffix(lower, ".xlsx") {
		http.Error(w, "仅支持 .xls/.xlsx 文件", http.StatusBadRequest)
		return
	}

	// 获取可执行文件所在目录，将上传文件保存到同目录的 data/upload 下
	exePath, err := os.Executable()
	if err != nil {
		http.Error(w, fmt.Sprintf("获取程序路径失败: %v", err), http.StatusInternalServerError)
		return
	}
	exeDir := filepath.Dir(exePath)
	uploadDir := filepath.Join(exeDir, "data", "upload")
	outputDir := filepath.Join(exeDir, "data", "output")
	_ = os.MkdirAll(uploadDir, 0755)
	_ = os.MkdirAll(outputDir, 0755)

	ext := filepath.Ext(header.Filename)
	base := strings.TrimSuffix(header.Filename, ext)
	timestamp := time.Now().Format("20060102_150405")
	savedName := fmt.Sprintf("%s_%s%s", base, timestamp, ext)
	savePath := filepath.Join(uploadDir, savedName)

	outFile, err := os.Create(savePath)
	if err != nil {
		http.Error(w, fmt.Sprintf("创建保存文件失败: %v", err), http.StatusInternalServerError)
		return
	}

	if _, err := io.Copy(outFile, file); err != nil {
		outFile.Close()
		http.Error(w, fmt.Sprintf("保存文件失败: %v", err), http.StatusInternalServerError)
		return
	}
	outFile.Close()

	// 解析
	result, err := core.ProcessSingleFile(savePath, outputDir)
	if err != nil {
		http.Error(w, fmt.Sprintf("解析失败: %v", err), http.StatusInternalServerError)
		return
	}

	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	json.NewEncoder(w).Encode(result)
}

func openBrowser(url string) {
	var cmd string
	var args []string
	switch runtime.GOOS {
	case "windows":
		cmd = "cmd"
		args = []string{"/c", "start", url}
	case "darwin":
		cmd = "open"
		args = []string{url}
	default:
		cmd = "xdg-open"
		args = []string{url}
	}
	if err := exec.Command(cmd, args...).Start(); err != nil {
		log.Printf("自动打开浏览器失败: %v", err)
	}
}
