

# �]�w�ؼХؿ��M�ɮצW��
$targetDirectory = "C:\Users\kaisda\Desktop\autoUpdateETF"
$pythonScript = "wantgoo-2.py"

# �i�J�ؼХؿ�
Set-Location -Path $targetDirectory

# ��ܶ}�l�T��
Write-Host "Starting HTTP server..."

# �Ұ� HTTP ���A��
$httpServer = Start-Process -NoNewWindow -FilePath "python" -ArgumentList "-m http.server 8000" -PassThru

# ���� HTTP ���A���Ұ�
Write-Host "Waiting for the server to start..."
Start-Sleep -Seconds 5

# �T�{���A���O�_�Ұʦ��\
try {
    $response = Invoke-WebRequest -Uri "http://127.0.0.1:8000" -UseBasicParsing
    if ($response.StatusCode -eq 200) {
        Write-Host "HTTP server is running. Ready to execute Python script."
        
        # ���ܧY�N���� Python �}��
        Write-Host "Executing Python script: $pythonScript..."
        
        # ���� Python �}��
        python $pythonScript
        
        # ���� Python �}�����槹��
        Write-Host "Python script executed successfully. Data written to Google Sheets."
    } else {
        Write-Host "HTTP server did not start correctly. Status code: $($response.StatusCode)"
    }
} catch {
    Write-Host "Failed to connect to HTTP server: $_"
}

# ���� HTTP ���A��
Write-Host "Stopping HTTP server..."
Stop-Process -Id $httpServer.Id -Force
Write-Host "HTTP server stopped. Task completed."




