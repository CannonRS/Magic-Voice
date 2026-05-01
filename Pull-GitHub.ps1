<#
.SYNOPSIS
    Holt Änderungen von GitHub (git pull). Siehe Sync-GitHub.ps1 für alle Parameter.
#>
param(
    [string]$RemoteName = "origin",
    [string]$RemoteUrl = "https://github.com/CannonRS/Magic-Voice.git",
    [switch]$Rebase
)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Pull -RemoteName $RemoteName -RemoteUrl $RemoteUrl -Rebase:$Rebase
