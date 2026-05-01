param(
    [switch]$RemoveUserData
)

$ErrorActionPreference = "Stop"

$target = Join-Path $env:LOCALAPPDATA "Programs\Magic-Voice"
$startMenu = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\Magic-Voice.lnk"
$desktop = [Environment]::GetFolderPath("DesktopDirectory")
$desktopLink = Join-Path $desktop "Magic-Voice.lnk"
$runKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$userData = Join-Path $env:LOCALAPPDATA "Magic-Voice"

Remove-Item -Path $startMenu -Force -ErrorAction SilentlyContinue
Remove-Item -Path $desktopLink -Force -ErrorAction SilentlyContinue
Remove-ItemProperty -Path $runKey -Name "Magic-Voice" -ErrorAction SilentlyContinue

if (Test-Path $target) {
    Remove-Item -Path $target -Recurse -Force
}

if ($RemoveUserData -and (Test-Path $userData)) {
    Remove-Item -Path $userData -Recurse -Force
}

Write-Host "Magic-Voice wurde deinstalliert. Nutzerdaten wurden nur mit -RemoveUserData entfernt."
