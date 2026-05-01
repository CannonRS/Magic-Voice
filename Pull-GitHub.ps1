<#
.SYNOPSIS
    git pull im Repo-Root (siehe Sync-GitHub.ps1).
#>
param([switch]$Rebase)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Pull -Rebase:$Rebase
