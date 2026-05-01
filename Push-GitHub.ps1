<#
.SYNOPSIS
    git push im Repo-Root (siehe Sync-GitHub.ps1).
#>
param(
    [string]$Branch = "main",
    [switch]$PushForceWithLease
)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Push -Branch $Branch -PushForceWithLease:$PushForceWithLease
