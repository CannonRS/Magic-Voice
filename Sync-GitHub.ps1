<#
.SYNOPSIS
    Synchronisiert mit dem konfigurierten Git-Remote (z. B. GitHub).

.DESCRIPTION
    Nutzt dieselbe Konfiguration wie auf der Kommandozeile: Remotes und Upstream
    stehen in .git/config (einmalig z. B. "git remote add origin ..." und nach dem
    ersten Push "git push -u origin <Branch>").

    Es wird immer der aktuell ausgecheckte Branch verwendet — nicht fest "main".
    Wenn du auf "main" stehst, ist das main; auf "feature/x" entsprechend feature/x.

.PARAMETER Action
    Pull = git pull, Push = git push, PullPush (Standard) = zuerst pull, dann push.

.PARAMETER Rebase
    Bei Pull: git pull --rebase.

.PARAMETER PushForceWithLease
    Bei Push: git push --force-with-lease (nur mit Absicht).
#>
param(
    [ValidateSet("Pull", "Push", "PullPush")]
    [string]$Action = "PullPush",
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

function Show-Context {
    $branch = (& git -C $repoRoot branch --show-current 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($branch)) {
        throw "Kein Branch ermittelbar (detached HEAD?). Bitte einen Branch auschecken."
    }

    $branch = $branch.Trim()
    Write-Host "Aktueller Branch: $branch | Repo: $repoRoot"

    $upstream = (& git -C $repoRoot rev-parse --abbrev-ref '@{upstream}' 2>$null)
    if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($upstream)) {
        Write-Host "Upstream: $upstream"
    }
    else {
        Write-Host "Kein Upstream gesetzt — erster Push z. B.: git push -u origin $branch"
    }
}

Assert-GitRepo
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
