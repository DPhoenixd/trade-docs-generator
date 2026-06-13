$ErrorActionPreference = "Stop"

$projectDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$pythonExe = Join-Path $projectDir ".venv\Scripts\python.exe"
$appFile = Join-Path $projectDir "app.py"
$url = "http://localhost:8501/"

function Test-PortOpen {
    param([string]$HostName, [int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $async = $client.BeginConnect($HostName, $Port, $null, $null)
        $connected = $async.AsyncWaitHandle.WaitOne(500, $false)
        if ($connected) {
            $client.EndConnect($async)
        }
        $client.Close()
        return $connected
    } catch {
        return $false
    }
}

if (-not (Test-Path -LiteralPath $pythonExe)) {
    Write-Host "Virtual environment was not found:" $pythonExe
    Write-Host "Please install dependencies first."
    Read-Host "Press Enter to close"
    exit 1
}

if (-not (Test-Path -LiteralPath $appFile)) {
    Write-Host "app.py was not found:" $appFile
    Read-Host "Press Enter to close"
    exit 1
}

Set-Location -LiteralPath $projectDir

if (Test-PortOpen -HostName "127.0.0.1" -Port 8501) {
    Start-Process $url
    exit 0
}

Start-Process $url
& $pythonExe -m streamlit run $appFile --server.port 8501
