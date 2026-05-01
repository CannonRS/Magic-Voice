<#
.SYNOPSIS
    Sendet den aktuellen Branch zu GitHub (git push). Siehe Sync-GitHub.ps1 für alle Parameter.
#>
param(
    [string]$RemoteName = "origin",
    [string]$RemoteUrl = "https://github.com/CannonRS/Magic-Voice.git",
    [switch]$PushForceWithLease
)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Push -RemoteName $RemoteName -RemoteUrl $RemoteUrl -PushForceWithLease:$PushForceWithLease
