function Write-PublishManifest {
    param(
        [Parameter(Mandatory)][string]$PublishDirectory,
        [Parameter(Mandatory)][string]$ManifestPath,
        [Parameter(Mandatory)][string[]]$AllowedRelativePaths
    )
    $publishRoot = [IO.Path]::GetFullPath($PublishDirectory)
    $allowed = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($relativePath in $AllowedRelativePaths) { [void]$allowed.Add($relativePath.Replace('\', '/')) }
    $files = @(Get-ChildItem -LiteralPath $publishRoot -Recurse -File)
    if (-not (Test-Path -LiteralPath (Join-Path $publishRoot 'CashRec.exe') -PathType Leaf)) { throw 'Published application executable is missing' }
    $forbidden = @('.cs', '.py', '.pdb', '.csv', '.xlsx', '.xls', '.wav', '.mp3', '.mp4', '.srt', '.vtt', '.png', '.jpg', '.jpeg')
    $records = foreach ($file in $files) {
        $relative = [IO.Path]::GetRelativePath($publishRoot, $file.FullName).Replace('\', '/')
        if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Publish contains a linked file: $relative" }
        if ($forbidden -contains $file.Extension.ToLowerInvariant() -or -not $allowed.Contains($relative)) {
            throw "Unapproved publish content: $relative"
        }
        [pscustomobject]@{ Path = $relative; Bytes = $file.Length; SHA256 = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash }
    }
    $manifest = [pscustomobject]@{
        Application = 'Cash Rec'
        Runtime = 'win-x64 self-contained'
        PublishDirectory = $publishRoot
        Files = @($records | Sort-Object Path)
    }
    [IO.File]::WriteAllText([IO.Path]::GetFullPath($ManifestPath), ($manifest | ConvertTo-Json -Depth 5) + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
    return $manifest
}
