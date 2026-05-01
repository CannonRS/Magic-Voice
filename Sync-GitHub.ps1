<#
.SYNOPSIS
    Synchronisiert das lokale Repository mit GitHub (Pull / Push).

.DESCRIPTION
    Legt bei Bedarf den Remote "origin" an (Standard-URL: CannonRS/Magic-Voice).
    Arbeitsverzeichnis ist immer das Verzeichnis dieses Skripts (Repo-Root).

.PARAMETER Action
    Pull = nur holen, Push = nur senden, PullPush = zuerst Pull, dann Push.

.PARAMETER RemoteUrl
    URL für "git remote add", falls der Remote noch fehlt.

.PARAMETER Rebase
    Bei Pull: git pull --rebase (sauberer Verlauf, Vorsicht bei lokalen Merge-Commits).

.PARAMETER PushForceWithLease
    Bei Push: git push --force-with-lease (nur nutzen, wenn du weißt, warum).
#>
param(
    [ValidateSet("Pull", "Push", "PullPush")]
    [string]$Action = "PullPush",
    [string]$RemoteName = "origin",
    [string]$RemoteUrl = "https://github.com/CannonRS/Magic-Voice.git",
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

function Get-CurrentBranch {
    $branch = (& git -C $repoRoot branch --show-current 2>$null)
    if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($branch)) {
        throw "Kein Branch ermittelbar (detached HEAD?). Bitte auf einen Branch wechseln (z. B. main)."
    }

    return $branch.Trim()
}

function Ensure-Remote {
    & git -C $repoRoot remote get-url $RemoteName 2>$null | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Remote '$RemoteName' fehlt — wird angelegt: $RemoteUrl"
        Invoke-RepoGit @("remote", "add", $RemoteName, $RemoteUrl)
    }
    else {
        $current = (& git -C $repoRoot remote get-url $RemoteName).Trim()
        Write-Host "Remote '$RemoteName' -> $current"
    }
}

$branch = Get-CurrentBranch
Write-Host "Branch: $branch | Repo: $repoRoot"

Ensure-Remote

switch ($Action) {
    "Pull" {
        if ($Rebase) {
            Invoke-RepoGit @("pull", "--rebase", $RemoteName, $branch)
        }
        else {
            Invoke-RepoGit @("pull", $RemoteName, $branch)
        }
    }
    "Push" {
        if ($PushForceWithLease) {
            Invoke-RepoGit @("push", "--force-with-lease", "-u", $RemoteName, $branch)
        }
        else {
            Invoke-RepoGit @("push", "-u", $RemoteName, $branch)
        }
    }
    "PullPush" {
        if ($Rebase) {
            Invoke-RepoGit @("pull", "--rebase", $RemoteName, $branch)
        }
        else {
            Invoke-RepoGit @("pull", $RemoteName, $branch)
        }

        if ($PushForceWithLease) {
            Invoke-RepoGit @("push", "--force-with-lease", "-u", $RemoteName, $branch)
        }
        else {
            Invoke-RepoGit @("push", "-u", $RemoteName, $branch)
        }
    }
}

Write-Host "Fertig ($Action)."
