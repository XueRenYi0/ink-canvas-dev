$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$r = Join-Path $repo 'Releases'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$check = Join-Path $r 'check-v4'
if (Test-Path -LiteralPath $check)
{
    Remove-Item -LiteralPath $check -Recurse -Force
}
New-Item -ItemType Directory -Path $check -Force | Out-Null
$portable = Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Name -like '*Portable*.zip' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$setupZip = Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Name -like '*Setup*.zip' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $portable -or -not $setupZip)
{
    Write-Host 'Missing zips:'
    Get-ChildItem -LiteralPath $r -File | ForEach-Object { Write-Host ('  ' + $_.Name + '  ' + [math]::Round($_.Length/1MB,2) + ' MB') }
    exit 1
}
Write-Host ('PORTABLE: ' + $portable.Name + '  ' + [math]::Round($portable.Length/1MB,2) + ' MB')
Write-Host ('SETUP   : ' + $setupZip.Name + '  ' + [math]::Round($setupZip.Length/1MB,2) + ' MB')
[IO.Compression.ZipFile]::ExtractToDirectory($portable.FullName, (Join-Path $check 'portable'))
$exe = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter 'Ink Canvas.exe' | Select-Object -First 1
$ini = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter 'VersionInfo.ini' | Select-Object -First 1
$readme = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter '*README*' | Select-Object -First 1
Write-Host ''
Write-Host '--- Portable 版本校验 ---'
$fvi = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe.FullName)
$asmV = [Reflection.AssemblyName]::GetAssemblyName($exe.FullName).Version
'  AssemblyVersion: ' + $asmV
'  FileVersion    : ' + $fvi.FileVersion
'  ProductVersion : ' + $fvi.ProductVersion
'  VersionInfo.ini: ' + [IO.File]::ReadAllText($ini.FullName).Trim()
$readmeText = [IO.File]::ReadAllText($readme.FullName)
$m1 = [regex]::Match($readmeText, 'Ink Canvas · [^\r\n]* v([\d\.]+)').Groups[1].Value
if (-not $m1) { $m1 = [regex]::Match($readmeText, 'v(\d+\.\d+\.\d+)').Groups[1].Value }
Write-Host ('  README head v: ' + $m1)
[IO.Compression.ZipFile]::ExtractToDirectory($setupZip.FullName, (Join-Path $check 'setup'))
$insidePayloadExe = Get-ChildItem -LiteralPath (Join-Path $check 'setup') -Recurse -Filter 'Ink Canvas.exe' | Select-Object -First 1
Write-Host ''
Write-Host '--- Setup 内容版本校验 ---'
$fvi2 = [Diagnostics.FileVersionInfo]::GetVersionInfo($insidePayloadExe.FullName)
$asmV2 = [Reflection.AssemblyName]::GetAssemblyName($insidePayloadExe.FullName).Version
'  payload AssemblyVersion: ' + $asmV2
'  payload FileVersion    : ' + $fvi2.FileVersion
Write-Host ''
$ok = ($asmV.Major -ge 5) -and ($asmV2.Major -ge 5) -and ([IO.File]::ReadAllText($ini.FullName).Trim() -eq '5.0.0') -and ($m1 -like '5*')
Remove-Item -LiteralPath $check -Recurse -Force
Write-Host ('全部校验通过 >=5.0.0:  ' + $ok)
