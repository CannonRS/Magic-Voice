<#
.SYNOPSIS
    Synchronisiert mit dem konfigurierten Git-Remote (z. B. GitHub).

.DESCRIPTION
    Remotes und Upstream kommen aus .git/config (wie bei normaler Git-Nutzung).

    Standard: vor Pull/Push wird auf den Branch gewechselt, den -Branch angibt
    (Vorgabe: main). Mit -UseCurrentBranch bleibt der gerade ausgecheckte Branch.

    Die Skripte enthalten keine Zugangsdaten; sie sind nur Hilfen wie Build.ps1.

.PARAMETER Action
    Pull, Push oder PullPush (Standard: erst pull, dann push).

.PARAMETER Branch
    Ziel-Branch für "git switch" (Vorgabe: main). Nur wirkungsvoll ohne -UseCurrentBranch.

.PARAMETER UseCurrentBranch
    Kein Branch-Wechsel — Pull/Push auf dem Branch, auf dem du schon bist.

.PARAMETER Rebase
    Bei Pull: git pull --rebase.

.PARAMETER PushForceWithLease
    Bei Push: git push --force-with-lease (nur mit Absicht).
#>
param(
    [ValidateSet("Pull", "Push", "PullPush")]
    [string]$Action = "PullPush",
    [string]$Branch = "main",
    [switch]$UseCurrentBranch,
    [switch]$Rebase,
    [switch]$PushForceWithLease
)

$ErrorActionPreference = "Stop"
$repoRoot = $PSScriptRoot

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
    throw "git wurde nicht gefunden. Git for Windows installieren und die Sitzung neu starten."
}

function Invoke-RepoGit {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$Arguments
    )

    Write-Host ("git " + ($Arguments -join " "))
    & git -C $repoRoot @Arguments
    if ($LASTEXITCODE -ne 0) {
        throw "git-Befehl ist fehlgeschlagen (Exit $LASTEXITCODE)."
    }
}

function Assert-GitRepo {
    if (-not (Test-Path -LiteralPath (Join-Path $repoRoot ".git"))) {
        throw "Kein Git-Repository (.git fehlt unter $repoRoot)."
    }
}

function Ensure-TargetBranch {
    if ($UseCurrentBranch) {
        Write-Host "Branch-Wechsel übersprungen (-UseCurrentBranch)."
        return
    }

    if ([string]::IsNullOrWhiteSpace($Branch)) {
        throw "Branch ist leer — -Branch angeben oder -UseCurrentBranch nutzen."
    }

    Write-Host "git switch $Branch"
    Invoke-RepoGit @("switch", $Branch)
}

function Show-Context {
    $current = (& git -C $repoRoot branch --show-current 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($current)) {
        throw "Kein Branch ermittelbar (detached HEAD?). Bitte einen Branch auschecken."
    }

    $current = $current.Trim()
    Write-Host "Aktueller Branch: $current | Repo: $repoRoot"

    $upstream = (& git -C $repoRoot rev-parse --abbrev-ref '@{upstream}' 2>$null)
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($upstream)) {
        Write-Host "Upstream: $upstream"
    }
    else {
        Write-Host "Kein Upstream gesetzt — erster Push z. B.: git push -u origin $current"
    }
}

Assert-GitRepo
Ensure-TargetBranch
Show-Context

switch ($Action) {
    "Pull" {
        if ($Rebase) {
            Invoke-RepoGit @("pull", "--rebase")
        }
        else {
            Invoke-RepoGit @("pull")
        }
    }
    "Push" {
        if ($PushForceWithLease) {
            Invoke-RepoGit @("push", "--force-with-lease")
        }
        else {
            Invoke-RepoGit @("push")
        }
    }
    "PullPush" {
        if ($Rebase) {
            Invoke-RepoGit @("pull", "--rebase")
        }
        else {
            Invoke-RepoGit @("pull")
        }

        if ($PushForceWithLease) {
            Invoke-RepoGit @("push", "--force-with-lease")
        }
        else {
            Invoke-RepoGit @("push")
        }
    }
}

Write-Host "Fertig ($Action)."
