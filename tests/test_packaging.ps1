param([Parameter(Mandatory)][string]$Dotnet)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $projectRoot ('.scratch/packaging-regression-' + [Guid]::NewGuid().ToString('N'))
$oldPublish = Join-Path $outputRoot 'publish'
New-Item -ItemType Directory -Force -Path $oldPublish | Out-Null
$sentinel = Join-Path $oldPublish 'never-package-sentinel.cs'
[IO.File]::WriteAllText($sentinel, 'stale source must never be included')
& (Join-Path $projectRoot 'build/build.ps1') -Dotnet $Dotnet -OutputRoot $outputRoot -SkipInstaller
$first = Get-Content -LiteralPath (Join-Path $outputRoot 'publish-manifest.json') -Raw | ConvertFrom-Json
if ($first.PublishDirectory -eq $oldPublish) { throw 'Build reused the stale publish folder' }
if ($first.Files.Path -contains 'never-package-sentinel.cs') { throw 'Stale source entered manifest' }
foreach ($entry in $first.Files) {
    $actual = Get-FileHash -LiteralPath (Join-Path $first.PublishDirectory $entry.Path) -Algorithm SHA256
    if ($actual.Hash -ne $entry.SHA256) { throw "Manifest hash mismatch: $($entry.Path)" }
}
. (Join-Path $projectRoot 'build/PublishManifest.ps1')
[IO.File]::WriteAllText((Join-Path $first.PublishDirectory 'leaked.py'), 'must be rejected')
$rejected = $false
try { $null = Write-PublishManifest -PublishDirectory $first.PublishDirectory -ManifestPath (Join-Path $outputRoot 'invalid.json') -AllowedRelativePaths @($first.Files.Path) }
catch { if ($_.Exception.Message -like '*Unapproved publish content*') { $rejected = $true } else { throw } }
if (-not $rejected) { throw 'Source in staging was not rejected' }
& (Join-Path $projectRoot 'build/build.ps1') -Dotnet $Dotnet -OutputRoot $outputRoot -SkipInstaller
$second = Get-Content -LiteralPath (Join-Path $outputRoot 'publish-manifest.json') -Raw | ConvertFrom-Json
if ($first.PublishDirectory -eq $second.PublishDirectory) { throw 'Build staging directory was reused' }
if ($second.Files.Path -contains 'leaked.py') { throw 'Previous staging source entered new build' }
if (-not (Test-Path -LiteralPath $sentinel)) { throw 'Build modified the preserved stale directory' }
Write-Output 'PASS: fresh publish staging; approved contents and hashes; source rejection; stale folder preserved.'
Write-Output "Synthetic packaging evidence: $outputRoot"
