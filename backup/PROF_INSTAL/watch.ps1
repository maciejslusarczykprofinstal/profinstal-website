# watch.ps1 – auto-deploy po zapisie index.html
$ErrorActionPreference = 'Stop'

$here   = Split-Path -Parent $MyInvocation.MyCommand.Path
$deploy = Join-Path $here 'deploy.ps1'

# Poprawne utworzenie FileSystemWatcher (bez pozycyjnych parametrów)
$fsw = New-Object System.IO.FileSystemWatcher -Property @{
    Path                  = $here
    Filter                = 'index.html'
    IncludeSubdirectories = $false
    NotifyFilter          = [IO.NotifyFilters]'FileName, LastWrite, Size'
}
$fsw.EnableRaisingEvents = $true

# Prosty debounce, żeby nie wywoływać deployu kilka razy przy jednym zapisie
$script:busy = $false
$action = {
    if ($script:busy) { return }
    $script:busy = $true
    Start-Sleep -Milliseconds 250
    try {
        Write-Host "$(Get-Date -Format HH:mm:ss) → deploy"
        Start-Process powershell -ArgumentList '-ExecutionPolicy Bypass -File', $deploy -WindowStyle Hidden
    }
    finally { $script:busy = $false }
}

Register-ObjectEvent -InputObject $fsw -EventName Changed -Action $action | Out-Null
Write-Host "👀 Watching index.html…  (Ctrl+C aby przerwać)"
while ($true) { Start-Sleep 1 }
