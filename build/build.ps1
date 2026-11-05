param(
    [string]$Dotnet = 'dotnet',
    [string]$Iscc = 'ISCC.exe',
    [string]$OutputRoot = '',
    [switch]$SkipInstaller
)
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
if ($OutputRoot -eq '') { $OutputRoot = Join-Path $projectRoot 'artifacts' }
$OutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$publishDir = Join-Path $OutputRoot (Join-Path 'staging' ([Guid]::NewGuid().ToString('N')))
New-Item -ItemType Directory -Path $publishDir -Force | Out-Null
& $Dotnet publish (Join-Path $projectRoot 'src/CashRec/CashRec.csproj') -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -p:DebugSymbols=false -o $publishDir
if ($LASTEXITCODE -ne 0) { throw 'Application publish failed' }
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination (Join-Path $publishDir 'README.md')
Copy-Item -LiteralPath (Join-Path $projectRoot 'THIRD-PARTY-NOTICES.md') -Destination (Join-Path $publishDir 'THIRD-PARTY-NOTICES.md')
Copy-Item -LiteralPath (Join-Path $projectRoot 'third-party') -Destination $publishDir -Recurse -Force
$sdkDirectory = Split-Path -Parent (Get-Command $Dotnet -ErrorAction Stop).Source
foreach ($noticeName in @('LICENSE.txt', 'ThirdPartyNotices.txt')) {
    $noticePath = Join-Path $sdkDirectory $noticeName
    if (-not (Test-Path -LiteralPath $noticePath)) { throw "SDK notice missing: $noticePath" }
    Copy-Item -LiteralPath $noticePath -Destination (Join-Path $publishDir "dotnet-$noticeName")
}
$approvedFiles = @('CashRec.exe', 'README.md', 'THIRD-PARTY-NOTICES.md', 'dotnet-LICENSE.txt', 'dotnet-ThirdPartyNotices.txt')
$approvedFiles += @(Get-ChildItem -LiteralPath (Join-Path $projectRoot 'third-party') -File -Filter '*-LICENSE.txt' | ForEach-Object { 'third-party/' + $_.Name })
. (Join-Path $PSScriptRoot 'PublishManifest.ps1')
$manifestPath = Join-Path $OutputRoot 'publish-manifest.json'
$null = Write-PublishManifest -PublishDirectory $publishDir -ManifestPath $manifestPath -AllowedRelativePaths $approvedFiles
Write-Output "Verified publish manifest: $manifestPath"
if (-not $SkipInstaller) {
    & $Iscc "/DPublishDir=$publishDir" "/DOutputDir=$(Join-Path $OutputRoot 'installer')" (Join-Path $PSScriptRoot 'CashRec.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed' }
}
