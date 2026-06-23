$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
  throw "没有找到 .venv Python，请先安装 requirements.txt"
}

Start-Process -FilePath $Python `
  -ArgumentList @("-m", "uvicorn", "trade_docs.api_server:app", "--host", "127.0.0.1", "--port", "8787") `
  -WorkingDirectory $Root `
  -WindowStyle Hidden

Start-Process -FilePath "cmd.exe" `
  -ArgumentList @("/c", "npm run dev -- --port 5173") `
  -WorkingDirectory (Join-Path $Root "pipl-frontend") `
  -WindowStyle Hidden

Start-Sleep -Seconds 3
Start-Process "http://localhost:5173"
