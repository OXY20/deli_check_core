package web

import (
	_ "embed"
	"encoding/json"
	"fmt"
	"io"
	"log"
	"net"
	"net/http"
	"os"
	"os/exec"
	"runtime"
	"strconv"
	"strings"
	"time"

	"deli_check_core/core"
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

	// 写入临时文件
	tmpFile, err := os.CreateTemp("", "deli_upload_*.xls")
	if err != nil {
		http.Error(w, fmt.Sprintf("创建临时文件失败: %v", err), http.StatusInternalServerError)
		return
	}
	defer os.Remove(tmpFile.Name())

	if _, err := io.Copy(tmpFile, file); err != nil {
		tmpFile.Close()
		http.Error(w, fmt.Sprintf("保存文件失败: %v", err), http.StatusInternalServerError)
		return
	}
	tmpFile.Close()

	// 解析
	outputDir := "data/output"
	_ = os.MkdirAll(outputDir, 0755)
	result, err := core.ProcessSingleFile(tmpFile.Name(), outputDir)
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
