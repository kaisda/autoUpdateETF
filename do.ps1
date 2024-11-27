

# 設定目標目錄和檔案名稱
$targetDirectory = "C:\Users\kaisda\Desktop\autoUpdateETF"
$pythonScript = "wantgoo-2.py"

# 進入目標目錄
Set-Location -Path $targetDirectory

# 顯示開始訊息
Write-Host "Starting HTTP server..."

# 啟動 HTTP 伺服器
$httpServer = Start-Process -NoNewWindow -FilePath "python" -ArgumentList "-m http.server 8000" -PassThru

# 等待 HTTP 伺服器啟動
Write-Host "Waiting for the server to start..."
Start-Sleep -Seconds 5

# 確認伺服器是否啟動成功
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:8000" -UseBasicParsing
    if ($response.StatusCode -eq 200) {
        Write-Host "HTTP server is running. Ready to execute Python script."
        
        # 提示即將執行 Python 腳本
        Write-Host "Executing Python script: $pythonScript..."
        
        # 執行 Python 腳本
        python $pythonScript
        
        # 提示 Python 腳本執行完畢
        Write-Host "Python script executed successfully. Data written to Google Sheets."
    } else {
        Write-Host "HTTP server did not start correctly. Status code: $($response.StatusCode)"
    }
} catch {
    Write-Host "Failed to connect to HTTP server: $_"
}

# 停止 HTTP 伺服器
Write-Host "Stopping HTTP server..."
Stop-Process -Id $httpServer.Id -Force
Write-Host "HTTP server stopped. Task completed."




