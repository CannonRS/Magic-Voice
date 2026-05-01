<#
.SYNOPSIS
    Synchronisiert mit dem konfigurierten Git-Remote (z. B. GitHub).

.DESCRIPTION
    Fehlt "origin", wird es mit -OriginUrl angelegt (Vorgabe: öffentliches Repo
    CannonRS/Magic-Voice). Danach wie gewohnt Pull/Push; Upstream setzt -u beim
    ersten Push.

    Standard: vor Pull/Push wird auf -Branch gewechselt (Vorgabe: main).
    Mit -UseCurrentBranch bleibt der aktuelle Branch.

    Nur die Repo-URL, keine Zugangsdaten.

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

.PARAMETER OriginUrl
    URL für "git remote add origin", falls origin noch fehlt (HTTPS oder SSH).
#>
param(
    [ValidateSet("Pull", "Push", "PullPush")]
    [string]$Action = "PullPush",
    [string]$Branch = "main",
    [switch]$UseCurrentBranch,
    [switch]$Rebase,
    [switch]$PushForceWithLease,
    [string]$OriginUrl = "https://github.com/CannonRS/Magic-Voice.git"
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

    Invoke-RepoGit @("switch", $Branch)
}

function Get-CurrentBranchName {
    $name = (& git -C $repoRoot branch --show-current 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($name)) {
        throw "Kein Branch ermittelbar (detached HEAD?). Bitte einen Branch auschecken."
    }

    return $name.Trim()
}

function Test-UpstreamConfigured {
    $null = & git -C $repoRoot rev-parse --abbrev-ref '@{upstream}' 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Test-OriginExists {
    $null = & git -C $repoRoot remote get-url origin 2>$null
    return ($LASTEXITCODE -eq 0)
}

function Ensure-Origin {
    param([Parameter(Mandatory = $true)][string]$Url)

    if (Test-OriginExists) {
        $current = (& git -C $repoRoot remote get-url origin).Trim()
        Write-Host "Remote origin: $current"
        return
    }

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw "Kein Remote 'origin' — bitte -OriginUrl setzen (oder einmalig manuell: git remote add origin ...)."
    }

    Invoke-RepoGit @("remote", "add", "origin", $Url.Trim())
}

function Show-Context {
    $current = Get-CurrentBranchName
    Write-Host "Aktueller Branch: $current | Repo: $repoRoot"

    if (Test-UpstreamConfigured) {
        $upstream = (& git -C $repoRoot rev-parse --abbrev-ref '@{upstream}').Trim()
        Write-Host "Upstream: $upstream"
    }
    else {
        Write-Host "Kein Upstream — Pull/Push nutzen origin/$current (Tracking nach erstem Push -u)."
    }
}

function Invoke-RepoPull {
    param([switch]$UseRebase)

    if (Test-UpstreamConfigured) {
        if ($UseRebase) {
            Invoke-RepoGit @("pull", "--rebase")
        }
        else {
            Invoke-RepoGit @("pull")
        }

        return
    }

    if (-not (Test-OriginExists)) {
        throw "Intern: origin fehlt nach Ensure-Origin."
    }

    $b = Get-CurrentBranchName
    if ($UseRebase) {
        Invoke-RepoGit @("pull", "--rebase", "origin", $b)
    }
    else {
        Invoke-RepoGit @("pull", "origin", $b)
    }
}

function Invoke-RepoPush {
    param([switch]$ForceWithLease)

    if (Test-UpstreamConfigured) {
        if ($ForceWithLease) {
            Invoke-RepoGit @("push", "--force-with-lease")
        }
        else {
            Invoke-RepoGit @("push")
        }

        return
    }

    if (-not (Test-OriginExists)) {
        throw "Intern: origin fehlt nach Ensure-Origin."
    }

    $b = Get-CurrentBranchName
    if ($ForceWithLease) {
        Invoke-RepoGit @("push", "--force-with-lease", "-u", "origin", $b)
    }
    else {
        Invoke-RepoGit @("push", "-u", "origin", $b)
    }
}

Assert-GitRepo
Ensure-Origin -Url $OriginUrl
Ensure-TargetBranch
Show-Context

switch ($Action) {
    "Pull" {
        Invoke-RepoPull -UseRebase:$Rebase
    }
    "Push" {
        Invoke-RepoPush -ForceWithLease:$PushForceWithLease
    }
    "PullPush" {
        Invoke-RepoPull -UseRebase:$Rebase
        Invoke-RepoPush -ForceWithLease:$PushForceWithLease
    }
}

Write-Host "Fertig ($Action)."
