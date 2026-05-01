<#
.SYNOPSIS
    Synchronisiert mit dem konfigurierten Git-Remote (z. B. GitHub).

.DESCRIPTION
    Fehlt "origin", wird es mit -OriginUrl angelegt (Vorgabe: öffentliches Repo
    CannonRS/Magic-Voice). Danach wie gewohnt Pull/Push; Upstream setzt -u beim
    ersten Push.

    Es wird immer auf -Branch gewechselt (Vorgabe: main); Pull/Push ohne Upstream
    nutzt ausdrücklich diesen Branch (origin/<Branch>).

    -AllowUnrelatedHistories: einmal nötig, wenn Remote z. B. nur LICENSE/README
    aus der GitHub-Anlage hat und lokal ein anderer Root-Commit existiert.

.PARAMETER Action
    Pull, Push oder PullPush (Standard: erst pull, dann push).

.PARAMETER SkipPull
    Nur bei PullPush: kein pull, nur push (kann fehlschlagen, wenn Remote voraus ist).

.PARAMETER Branch
    Branch für git switch und für explizite pull/push origin/<Branch> (Vorgabe: main).

.PARAMETER Rebase
    Bei Pull: git pull --rebase (wird ignoriert, wenn -AllowUnrelatedHistories gesetzt).

.PARAMETER AllowUnrelatedHistories
    Bei Pull: --allow-unrelated-histories (Merge mit fremder Remote-Historie).

.PARAMETER PushForceWithLease
    Bei Push: git push --force-with-lease (nur mit Absicht).

.PARAMETER OriginUrl
    URL für "git remote add origin", falls origin noch fehlt (HTTPS oder SSH).
#>
param(
    [ValidateSet("Pull", "Push", "PullPush")]
    [string]$Action = "PullPush",
    [string]$Branch = "main",
    [switch]$Rebase,
    [switch]$AllowUnrelatedHistories,
    [switch]$PushForceWithLease,
    [switch]$SkipPull,
    [string]$OriginUrl = "https://github.com/CannonRS/Magic-Voice.git"
)

$ErrorActionPreference = "Stop"
$repoRoot = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($Branch)) {
    throw "Branch darf nicht leer sein (Vorgabe: main)."
}

$syncBranch = $Branch.Trim()

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

function Get-CurrentBranchName {
    $name = (& git -C $repoRoot branch --show-current 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($name)) {
        throw "Kein Branch ermittelbar (detached HEAD?). Bitte einen Branch auschecken."
    }

    return $name.Trim()
}

function Ensure-TargetBranch {
    Invoke-RepoGit @("switch", $syncBranch)
    $current = Get-CurrentBranchName
    if ($current -ne $syncBranch) {
        throw "Erwartet Branch '$syncBranch', ausgecheckt ist '$current'."
    }
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
    Write-Host "Synchronisations-Branch: $syncBranch | Repo: $repoRoot"

    if (Test-UpstreamConfigured) {
        $upstream = (& git -C $repoRoot rev-parse --abbrev-ref '@{upstream}').Trim()
        Write-Host "Upstream: $upstream"
    }
    else {
        Write-Host "Kein Upstream — Pull/Push nutzen origin/$syncBranch (Tracking nach erstem Push -u)."
    }
}

function Invoke-RepoPull {
    param(
        [switch]$UseRebase,
        [switch]$AllowUnrelated
    )

    if ($AllowUnrelated) {
        if ($UseRebase) {
            Write-Host "Hinweis: -Rebase wird bei -AllowUnrelatedHistories ignoriert (Merge)."
        }

        if (Test-UpstreamConfigured) {
            Invoke-RepoGit @("pull", "--allow-unrelated-histories")
        }
        else {
            Invoke-RepoGit @("pull", "origin", $syncBranch, "--allow-unrelated-histories")
        }

        return
    }

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

    if ($UseRebase) {
        Invoke-RepoGit @("pull", "--rebase", "origin", $syncBranch)
    }
    else {
        Invoke-RepoGit @("pull", "origin", $syncBranch)
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

    if ($ForceWithLease) {
        Invoke-RepoGit @("push", "--force-with-lease", "-u", "origin", $syncBranch)
    }
    else {
        Invoke-RepoGit @("push", "-u", "origin", $syncBranch)
    }
}

Assert-GitRepo
Ensure-Origin -Url $OriginUrl
Ensure-TargetBranch
Show-Context

switch ($Action) {
    "Pull" {
        Invoke-RepoPull -UseRebase:$Rebase -AllowUnrelated:$AllowUnrelatedHistories
    }
    "Push" {
        Invoke-RepoPush -ForceWithLease:$PushForceWithLease
    }
    "PullPush" {
        if (-not $SkipPull) {
            Invoke-RepoPull -UseRebase:$Rebase -AllowUnrelated:$AllowUnrelatedHistories
        }

        Invoke-RepoPush -ForceWithLease:$PushForceWithLease
    }
}

Write-Host "Fertig ($Action)."
