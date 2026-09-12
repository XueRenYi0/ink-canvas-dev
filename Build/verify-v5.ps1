$ErrorActionPreference = 'Continue'
# 校验 Releases 下最新打包产物版本一致性（Portable.zip 内部 + 文件名 + VersionInfo.ini + 使用说明）
# 产物：InkClass-vX.Y.Z-Portable.zip + InkClass-vX.Y.Z-Setup.exe（Setup 为 Inno 安装器，不内部解包校验）
$repo = Split-Path -Parent $PSScriptRoot
$r = Join-Path $repo 'Releases'
Add-Type -AssemblyName System.IO.Compression.FileSystem

$portable = Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Name -like 'InkClass-*-Portable.zip' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$setup = Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Name -like 'InkClass-*-Setup.exe' } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if (-not $portable -or -not $setup) {
    Write-Host 'Missing artifacts in Releases:'
    Get-ChildItem -LiteralPath $r -File | ForEach-Object { Write-Host ('  ' + $_.Name + '  ' + [math]::Round($_.Length/1MB,2) + ' MB') }
    exit 1
}
Write-Host ('PORTABLE: ' + $portable.Name + '  ' + [math]::Round($portable.Length/1MB,2) + ' MB')
Write-Host ('SETUP   : ' + $setup.Name + '  ' + [math]::Round($setup.Length/1MB,2) + ' MB')

# zip 文件名里的期望版本（如 6.1.0）
$expect = [regex]::Match($portable.Name, 'v(\d+\.\d+\.\d+)').Groups[1].Value
if (-not $expect) { Write-Host '无法从文件名解析版本'; exit 1 }

$check = Join-Path $r 'check-v4'
if (Test-Path -LiteralPath $check) { Remove-Item -LiteralPath $check -Recurse -Force }
New-Item -ItemType Directory -Path $check -Force | Out-Null
[IO.Compression.ZipFile]::ExtractToDirectory($portable.FullName, (Join-Path $check 'portable'))

$exe = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter 'InkClass.exe' | Select-Object -First 1
$ini = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter 'VersionInfo.ini' | Select-Object -First 1
$readme = Get-ChildItem -LiteralPath (Join-Path $check 'portable') -Recurse -Filter '*README*' | Select-Object -First 1

Write-Host ''
Write-Host '--- Portable 版本校验 ---'
$fvi = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe.FullName)
$asmV = [Reflection.AssemblyName]::GetAssemblyName($exe.FullName).Version
$iniV = [IO.File]::ReadAllText($ini.FullName).Trim()
Write-Host ('  AssemblyVersion : ' + $asmV)
Write-Host ('  FileVersion     : ' + $fvi.FileVersion)
Write-Host ('  ProductVersion  : ' + $fvi.ProductVersion)
Write-Host ('  VersionInfo.ini : ' + $iniV)
$readmeV = [regex]::Match([IO.File]::ReadAllText($readme.FullName), 'v(\d+\.\d+\.\d+)').Groups[1].Value
Write-Host ('  使用说明 README : ' + $readmeV)

Remove-Item -LiteralPath $check -Recurse -Force

# 一致性判定：AssemblyVersion / ini / 文件名 三处一致（FileVersion 含日期后缀不比对）
$ok = ($asmV.ToString(3) -eq $expect) -and ($iniV -eq $expect)
$setupOk = ([regex]::Match($setup.Name, 'v(\d+\.\d+\.\d+)').Groups[1].Value -eq $expect)
Write-Host ''
Write-Host ('版本一致性（exe/ini/文件名）: ' + $ok + '    Setup 版本一致: ' + $setupOk)
if (-not ($ok -and $setupOk)) { exit 1 }
