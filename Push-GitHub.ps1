<#
.SYNOPSIS
    git push im Repo-Root (siehe Sync-GitHub.ps1).
#>
param([switch]$PushForceWithLease)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Push -PushForceWithLease:$PushForceWithLease
