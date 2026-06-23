$ErrorActionPreference = "Stop"

$Root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
  throw "Missing .venv Python. Install requirements.txt first."
}

Push-Location $Root
try {
  Get-Process | Where-Object { $_.ProcessName -eq "PIPL-YiDianTeng" } | Stop-Process -Force -ErrorAction SilentlyContinue

  Write-Host "1/4 Building frontend..."
  Push-Location (Join-Path $Root "pipl-frontend")
  cmd /c npm install
  if ($LASTEXITCODE -ne 0) { throw "npm install failed" }
  cmd /c npm run build
  if ($LASTEXITCODE -ne 0) { throw "npm build failed" }
  Pop-Location

  Write-Host "2/4 Installing PyInstaller..."
  & $Python -m pip install --upgrade pyinstaller
  if ($LASTEXITCODE -ne 0) { throw "PyInstaller install failed" }

  Write-Host "3/4 Packaging desktop app..."
  & $Python -m PyInstaller --clean --noconfirm (Join-Path $Root "packaging\pipl_yidianteng.spec")
  if ($LASTEXITCODE -ne 0) { throw "PyInstaller packaging failed" }

  Write-Host "4/4 Creating delivery zip..."
  $PackageDir = Join-Path $Root "dist\PIPL-YiDianTeng"
  $ZipPath = Join-Path $Root "dist\PIPL-YiDianTeng-portable.zip"
  if (Test-Path $ZipPath) {
    Remove-Item -LiteralPath $ZipPath -Force
  }
  Compress-Archive -LiteralPath $PackageDir -DestinationPath $ZipPath

  Write-Host ""
  Write-Host "Done: $ZipPath"
  Write-Host "Send this zip to your colleague. Unzip it and double-click PIPL-YiDianTeng.exe."
}
finally {
  Pop-Location
}
