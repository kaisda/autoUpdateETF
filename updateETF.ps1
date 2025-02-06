# 設定 PowerShell 執行策略
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope Process -Force

# 切換到 Python 腳本所在的目錄 (PowerShell 會從這裡執行)
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
Set-Location -Path $scriptPath

# 設定 Python 虛擬環境
$pythonPath = ".\venv\Scripts\python.exe"  # 使用相對路徑

# 顯示開始訊息 (綠色)
Write-Host "🚀 正在執行 ETF 數據更新..." -ForegroundColor Green

# 執行 MoneyDJ 爬取
Write-Host "`n🔍 Running moneydj.py..." -ForegroundColor Cyan
& $pythonPath .\moneydj.py
if ($?) {
    Write-Host "✅ moneydj.py 執行成功！" -ForegroundColor Green
}
else {
    Write-Host "❌ moneydj.py 執行失敗！" -ForegroundColor Red
}

# 執行 WantGoo 爬取
Write-Host "`n🔍 Running wantgoo.py..." -ForegroundColor Cyan
& $pythonPath .\wantgoo.py
if ($?) {
    Write-Host "✅ wantgoo.py 執行成功！" -ForegroundColor Green
}
else {
    Write-Host "❌ wantgoo.py 執行失敗！" -ForegroundColor Red
}

# 顯示完成訊息 (黃色)
Write-Host "`n🎉 所有任務完成！" -ForegroundColor Yellow
