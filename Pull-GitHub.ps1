<#
.SYNOPSIS
    git pull im Repo-Root (siehe Sync-GitHub.ps1).
#>
param(
    [string]$Branch = "main",
    [switch]$Rebase
)

$ErrorActionPreference = "Stop"
& (Join-Path $PSScriptRoot "Sync-GitHub.ps1") -Action Pull -Branch $Branch -Rebase:$Rebase
