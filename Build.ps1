param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release",
    [switch]$Clean
)

$ErrorActionPreference = "Stop"

$repoRoot = $PSScriptRoot
$project = Join-Path $repoRoot "src\MagicVoice\MagicVoice.csproj"
$publishDir = Join-Path $repoRoot "artifacts\release\Magic-Voice"
$buildOutputDir = Join-Path $repoRoot "src\MagicVoice\bin\x64\$Configuration\net10.0-windows10.0.19041.0\win-x64"

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw "dotnet wurde nicht gefunden. .NET SDK laut global.json installieren und die Sitzung neu starten."
}

if ($Clean) {
    if (Test-Path -LiteralPath $publishDir) {
        Remove-Item -LiteralPath $publishDir -Recurse -Force
    }

    Write-Host "dotnet clean ..."
    & dotnet clean $project -c $Configuration -p:Platform=x64
    if ($LASTEXITCODE -ne 0) {
        throw "dotnet clean ist fehlgeschlagen."
    }
}

New-Item -ItemType Directory -Force -Path $publishDir | Out-Null

Write-Host "dotnet publish -> $publishDir"
& dotnet publish $project -c $Configuration -r win-x64 --self-contained false -p:Platform=x64 -o $publishDir
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish ist fehlgeschlagen."
}

# WinUI: .xbf und .pri (dotnet publish laesst sie oft aus; ohne sie crasht WinUI/XAML):
# sonst fehlen Ressourcen nach publish und die App kann crashen.
if (Test-Path -LiteralPath $buildOutputDir) {
    Get-ChildItem -Path $buildOutputDir -Filter "*.xbf" -ErrorAction SilentlyContinue |
        Copy-Item -Destination $publishDir -Force
    Get-ChildItem -Path $buildOutputDir -Filter "*.pri" -ErrorAction SilentlyContinue |
        Copy-Item -Destination $publishDir -Force
}

$exe = Join-Path $publishDir "MagicVoice.exe"
if (-not (Test-Path -LiteralPath $exe)) {
    throw "Build unvollständig: $exe fehlt."
}

Write-Host ""
Write-Host "Fertig. Den kompletten Ordner auf den Zielrechner kopieren (oder dort .\Install.ps1 -SourceDir <Pfad> nutzen):"
Write-Host "  $publishDir"
