$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$r = Join-Path $repo 'Releases'
$src = Join-Path $repo 'Ink Canvas\bin\Release'

if (-not (Test-Path -LiteralPath (Join-Path $src 'Ink Canvas.exe'))) {
    Write-Host ('未找到编译产物: ' + $src + '\Ink Canvas.exe — 请先运行 rebuild-release-v5.ps1')
    exit 1
}

# 清理旧 zip
Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Extension -eq '.zip' } | ForEach-Object -Process `
{
    Remove-Item -LiteralPath $_.FullName -Force
    Write-Host ('del: ' + $_.Name)
}

# 版本号（与 csproj / InkCanvas.iss / README 保持一致，升级时四处同步改）
$ver = '5.2.0'

# 暂存目录：从 bin\Release 拷贝，排除运行期用户数据与调试文件
$stagedir = Join-Path $r ('stage-v5\Ink Canvas v' + $ver)
$stageRoot = Split-Path -Parent $stagedir
if (Test-Path -LiteralPath $stageRoot) { Remove-Item -LiteralPath $stageRoot -Recurse -Force }
New-Item -ItemType Directory -Path $stagedir -Force | Out-Null
$exclude = @('Log.txt', 'Settings.json', 'Versions.ini', 'custom.json', 'Ink Canvas.pdb')
$excludeDirs = @('CustomShapes', 'StylusTest', 'AutoSavedStrokes', 'History Versions')
Get-ChildItem -LiteralPath $src -Force | Where-Object {
    ($_.PSIsContainer -and $excludeDirs -notcontains $_.Name) -or
    (-not $_.PSIsContainer -and $exclude -notcontains $_.Name)
} | ForEach-Object -Process `
{
    Copy-Item -LiteralPath $_.FullName -Destination (Join-Path $stagedir $_.Name) -Recurse -Force
}

# 打包内使用说明（模板存于 Build/，纳入版本管理，Rebuild 不会丢）
$readmeTpl = Join-Path $PSScriptRoot '使用说明 README.txt'
if (Test-Path -LiteralPath $readmeTpl) {
    $txt = [IO.File]::ReadAllText($readmeTpl)
    [IO.File]::WriteAllText((Join-Path $stagedir '使用说明 README.txt'), $txt, [Text.UTF8Encoding]::new($true))
    Write-Host '使用说明 README.txt: included'
} else {
    Write-Host 'WARN: 未找到 Build\使用说明 README.txt'
}

$exe = Join-Path $stagedir 'Ink Canvas.exe'
$fvi = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
$asm = [Reflection.AssemblyName]::GetAssemblyName($exe).Version
$iniPath = Join-Path $stagedir 'VersionInfo.ini'
Write-Host ('stage: Assembly=' + $asm + '  File=' + $fvi.FileVersion + '  ini=' + [IO.File]::ReadAllText($iniPath).Trim())

# Portable zip（暂存目录整体打包，解压后得到同名文件夹）
$zipP = Join-Path $r ('InkCanvas-v' + $ver + '-Portable.zip')
if (Test-Path -LiteralPath $zipP) { Remove-Item -LiteralPath $zipP -Force }
Compress-Archive -LiteralPath $stagedir -DestinationPath $zipP -CompressionLevel Optimal -Force
Write-Host ('+ portable: ' + $zipP + '   ' + [math]::Round((Get-Item -LiteralPath $zipP).Length/1MB,2) + ' MB')

# 单文件安装包：Inno Setup 编译（脚本 Build\InkCanvas.iss，需已安装 Inno Setup 6）
# 设计说明见 .iss 文件头注释：用户级安装（免 UAC）、正规卸载项、开始菜单/桌面快捷方式
$iscc = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Inno Setup 6\ISCC.exe'),
    (Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'),
    (Join-Path $env:ProgramFiles 'Inno Setup 6\ISCC.exe')
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

if ($iscc) {
    $iss = Join-Path $PSScriptRoot 'InkCanvas.iss'
    $setupExe = Join-Path $r ('InkCanvas-v' + $ver + '-Setup.exe')
    if (Test-Path -LiteralPath $setupExe) { Remove-Item -LiteralPath $setupExe -Force }
    & $iscc $iss | Out-Null
    if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $setupExe)) {
        Write-Host ('+ setup exe: ' + $setupExe + '   ' + [math]::Round((Get-Item -LiteralPath $setupExe).Length/1MB,2) + ' MB')
    } else {
        Write-Host 'WARN: Inno Setup 安装包编译失败（不影响 zip 产物）'
    }
} else {
    Write-Host 'WARN: 未找到 ISCC.exe（Inno Setup 6），跳过安装包。安装：winget install JRSoftware.InnoSetup'
}

# 清理暂存目录（安装包与 zip 已生成）
Remove-Item -LiteralPath $stageRoot -Recurse -Force
# 清理旧安装器产物（csc 时代的 SetupSource.exe / Setup.zip 已废弃）
Remove-Item -LiteralPath (Join-Path $r 'SetupSource.exe') -Force -ErrorAction SilentlyContinue
Get-ChildItem -LiteralPath $r -File | Where-Object { $_.Name -like '*-Setup.zip' } | Remove-Item -Force -ErrorAction SilentlyContinue

Write-Host ''
Write-Host '=== Releases ==='
Get-ChildItem -LiteralPath $r -File | Sort-Object LastWriteTime -Descending | ForEach-Object -Process `
{
    '{0}  {1,7:N2} MB  {2}' -f $_.LastWriteTime.ToString('HH:mm:ss'), ($_.Length/1MB), $_.Name
}
