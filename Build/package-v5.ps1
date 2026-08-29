$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$r = Join-Path $repo 'Releases'
$src = Join-Path $repo 'Ink Canvas\bin\x64\Release'
$stagedir = Join-Path $r 'stage-v5\Ink Canvas v5.0.0 x64'
$sroot = Join-Path $r 'Ink Canvas Setup x64 v5.0.0'
$payload = Join-Path $sroot 'payload'
# remove stale
$rms = @(
    (Join-Path $r 'stage-v5'),
    (Join-Path $r 'Ink Canvas Setup x64 v5.0.0'),
    (Join-Path $r 'stage-portable'),
    (Join-Path $r 'Ink Canvas v2.1.0 x64 免安装版.zip'),
    (Join-Path $r 'Ink Canvas v2.1.0 x64 安装包.zip'),
    (Join-Path $r 'InkCanvas-v2.1.0-x64-Setup.zip')
)
foreach ($p in $rms)
{
    if (Test-Path -LiteralPath $p)
    {
        Remove-Item -LiteralPath $p -Recurse -Force
    }
}
New-Item -ItemType Directory -Path $stagedir -Force | Out-Null
# copy files
Get-ChildItem -LiteralPath $src -File -Force | ForEach-Object -Process `
{
    if ($_.Extension -eq '.pdb') { return }
    if ($_.Extension -eq '.xml') { return }
    if ($_.Name -eq 'Settings.json') { return }
    if ($_.Name -eq 'Log.txt') { return }
    Copy-Item -LiteralPath $_.FullName -Destination $stagedir -Force
}
Get-ChildItem -LiteralPath $src -Directory -Force | ForEach-Object -Process `
{
    if ($_.Name -eq 'CustomShapes') { return }
    $d = Join-Path $stagedir $_.Name
    New-Item -ItemType Directory -Path $d -Force | Out-Null
    Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force | Copy-Item -Destination $d -Force
}
$count = (Get-ChildItem -LiteralPath $stagedir -Recurse -File).Count
Write-Host ('portable files: ' + $count)
# portable zip
$zipP = Join-Path $r 'Ink Canvas v5.0.0 x64 免安装版.zip'
if (Test-Path -LiteralPath $zipP) { Remove-Item -LiteralPath $zipP -Force }
Compress-Archive -LiteralPath $stagedir -DestinationPath $zipP -CompressionLevel Optimal -Force
# setup payload
New-Item -ItemType Directory -Path $payload -Force | Out-Null
Get-ChildItem -LiteralPath $stagedir -File -Force | Copy-Item -Destination $payload -Force
Get-ChildItem -LiteralPath $stagedir -Directory -Force | ForEach-Object -Process `
{
    $d = Join-Path $payload $_.Name
    New-Item -ItemType Directory -Path $d -Force | Out-Null
    Get-ChildItem -LiteralPath $_.FullName -Recurse -File -Force | Copy-Item -Destination $d -Force
}
Copy-Item -LiteralPath (Join-Path $r 'SetupSource.exe') -Destination (Join-Path $sroot 'Setup.exe') -Force
Write-Host ('setup payload files: ' + (Get-ChildItem -LiteralPath $payload -Recurse -File).Count)
# setup zip
$zipS = Join-Path $r 'InkCanvas-v5.0.0-x64-Setup.zip'
if (Test-Path -LiteralPath $zipS) { Remove-Item -LiteralPath $zipS -Force }
Compress-Archive -Path (Join-Path $sroot '*') -DestinationPath $zipS -CompressionLevel Optimal -Force
Write-Host '=== v5 packages ==='
Get-ChildItem -LiteralPath $r -File | ForEach-Object -Process `
{
    if ($_.Name -like '*v5*' -or $_.Name -like '*5.0*' -or $_.Name -eq 'SetupSource.exe')
    {
        '{0,7:N2} MB  {1}' -f ($_.Length / 1MB), $_.Name
    }
}
