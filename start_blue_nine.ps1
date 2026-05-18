# BLUE NINE — 개발 환경 일괄 기동 스크립트 (PowerShell)
# 사용: .\start_blue_nine.ps1

$ErrorActionPreference = "Stop"

Write-Host "==> Activating venv" -ForegroundColor Cyan
. .\.venv\Scripts\Activate.ps1

Write-Host "==> Backend: http://127.0.0.1:8088  (Swagger: /docs)" -ForegroundColor Green
$backend = Start-Process -PassThru -NoNewWindow powershell -ArgumentList "-NoProfile","-Command","python -m backend.run"

Write-Host "==> Frontend (Vite): http://localhost:5173" -ForegroundColor Green
Set-Location frontend
if (-not (Test-Path node_modules)) {
    Write-Host "    npm install (first run)..." -ForegroundColor Yellow
    npm install
}
npm run dev
