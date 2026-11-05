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
$publishDir = Join-Path $OutputRoot 'publish'
& $Dotnet publish (Join-Path $projectRoot 'src/CashRec/CashRec.csproj') -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:DebugType=None -p:DebugSymbols=false -o $publishDir
if ($LASTEXITCODE -ne 0) { throw 'Application publish failed' }
Copy-Item -LiteralPath (Join-Path $projectRoot 'README.md') -Destination (Join-Path $publishDir 'README.md')
Copy-Item -LiteralPath (Join-Path $projectRoot 'THIRD-PARTY-NOTICES.md') -Destination (Join-Path $publishDir 'THIRD-PARTY-NOTICES.md')
$sdkDirectory = Split-Path -Parent ([IO.Path]::GetFullPath($Dotnet))
foreach ($noticeName in @('LICENSE.txt', 'ThirdPartyNotices.txt')) {
    $noticePath = Join-Path $sdkDirectory $noticeName
    if (Test-Path -LiteralPath $noticePath) { Copy-Item -LiteralPath $noticePath -Destination (Join-Path $publishDir "dotnet-$noticeName") }
}
if (-not $SkipInstaller) {
    & $Iscc "/DPublishDir=$publishDir" "/DOutputDir=$(Join-Path $OutputRoot 'installer')" (Join-Path $PSScriptRoot 'CashRec.iss')
    if ($LASTEXITCODE -ne 0) { throw 'Installer build failed' }
}
