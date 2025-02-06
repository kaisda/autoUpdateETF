# 設定 PowerShell 執行策略
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser -Force

# 設定專案路徑與 Python 腳本
$targetDirectory = "C:\Users\user\Desktop\autoUpdateETF-master"
$pythonScript = "wantgoo-2.py"

# 切換到專案資料夾
Set-Location -Path $targetDirectory

# 啟動 HTTP 伺服器
Write-Host "Starting HTTP server..."
$httpServer = Start-Process -NoNewWindow -FilePath "python" -ArgumentList "-m http.server 8000" -PassThru

# 等待伺服器啟動
Write-Host "Waiting for the server to start..."
Start-Sleep -Seconds 5

# 確認 HTTP 伺服器是否正常運行
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:8000" -UseBasicParsing
    if ($response.StatusCode -eq 200) {
        Write-Host "HTTP server is running. Ready to execute Python script."

        # 執行 Python 腳本
        Write-Host "Executing Python script: $pythonScript..."
        & python $pythonScript
        Write-Host "Python script executed successfully. Data written to Google Sheets."
    }
    else {
        Write-Host "HTTP server did not start correctly. Status code: $($response.StatusCode)"
    }
}
catch {
    Write-Host "Failed to connect to HTTP server: $_"
}

# 關閉 HTTP 伺服器
Write-Host "Stopping HTTP server..."
Stop-Process -Name "python" -Force
Write-Host "HTTP server stopped. Task completed."
