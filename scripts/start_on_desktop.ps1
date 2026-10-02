# Starts BatchExcel in the logged-on user's desktop session. Anything launched directly from an
# SSH session runs in session 0, where WPF windows are invisible and Excel automation differs.
param([Parameter(Mandatory)][string]$ExePath)

$ErrorActionPreference = 'Stop'
$exe = (Resolve-Path $ExePath).Path
$taskName = 'BatchExcel-Debug'

$action = New-ScheduledTaskAction -Execute $exe -WorkingDirectory (Split-Path $exe)
$user = [Security.Principal.WindowsIdentity]::GetCurrent().Name
$principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Force | Out-Null

$existing = @(Get-Process BatchExcel -ErrorAction SilentlyContinue).Id
Start-ScheduledTask -TaskName $taskName

$deadline = (Get-Date).AddSeconds(30)
do {
    Start-Sleep -Milliseconds 250
    $proc = Get-Process BatchExcel -ErrorAction SilentlyContinue | Where-Object { $_.Id -notin $existing } | Select-Object -First 1
} until ($proc -or (Get-Date) -gt $deadline)

if (-not $proc) { throw "BatchExcel did not start. Is $user logged on to the VM desktop?" }
Write-Host "BatchExcel started on the desktop (PID $($proc.Id), session $($proc.SessionId))."
