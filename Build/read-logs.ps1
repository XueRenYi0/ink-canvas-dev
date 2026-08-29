$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$releaseLog  = Join-Path $repo 'Ink Canvas\bin\Release\Log.txt'
$debugLog    = Join-Path $repo 'Ink Canvas\bin\Debug\Log.txt'
$releasesLog = Join-Path $repo 'Releases\Log.txt'
$logs = @($releaseLog,$debugLog,$releasesLog)
foreach ($p in $logs)
{
    Write-Host ('== ' + $p + '  exists=' + (Test-Path -LiteralPath $p))
    if (Test-Path -LiteralPath $p)
    {
        $lines = Get-Content -LiteralPath $p -Tail 80 -ErrorAction SilentlyContinue
        foreach ($line in $lines) { Write-Host $line }
        Write-Host '--- EOF ---'
    }
}
$roots = @(
  [Environment]::GetFolderPath('Desktop'),
  (Join-Path ([Environment]::GetFolderPath('UserProfile')) 'Downloads'),
  [Environment]::GetFolderPath('UserProfile')
)
foreach ($root in $roots)
{
    if (Test-Path -LiteralPath $root)
    {
        Get-ChildItem -LiteralPath $root -Filter 'Log.txt' -Recurse -ErrorAction SilentlyContinue -Depth 5 |
            Select-Object -First 5 |
            ForEach-Object -Process `
            {
                Write-Host ('Found Log.txt: ' + $_.FullName)
                Get-Content -LiteralPath $_.FullName -Tail 40
            }
    }
}
