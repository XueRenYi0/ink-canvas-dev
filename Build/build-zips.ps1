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

# 暂存目录：从 bin\Release 拷贝，排除运行期用户数据与调试文件
$stagedir = Join-Path $r 'stage-v5\Ink Canvas v5.1.0'
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
$zipP = Join-Path $r 'InkCanvas-v5.1.0-Portable.zip'
if (Test-Path -LiteralPath $zipP) { Remove-Item -LiteralPath $zipP -Force }
Compress-Archive -LiteralPath $stagedir -DestinationPath $zipP -CompressionLevel Optimal -Force
Write-Host ('+ portable: ' + $zipP + '   ' + [math]::Round((Get-Item -LiteralPath $zipP).Length/1MB,2) + ' MB')

# 单文件安装包：payload 压成 zip 嵌入 Setup.exe 资源，双击即装、无需解压
$ver = '5.1.0'
$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$setupCs = Join-Path $PSScriptRoot 'InkCanvasSetup.cs'
if ((Test-Path -LiteralPath $csc) -and (Test-Path -LiteralPath $setupCs)) {
    $work = Join-Path $r 'stage-v5\single'
    New-Item -ItemType Directory -Path $work -Force | Out-Null
    $payloadZip = Join-Path $work 'payload.zip'
    Compress-Archive -Path (Join-Path $stagedir '*') -DestinationPath $payloadZip -CompressionLevel Optimal -Force
    $singleExe = Join-Path $r ('InkCanvas-v' + $ver + '-Setup.exe')
    if (Test-Path -LiteralPath $singleExe) { Remove-Item -LiteralPath $singleExe -Force }
    & $csc /nologo /target:winexe /optimize+ /out:$singleExe `
        /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll `
        /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll `
        /resource:$payloadZip,payload.zip $setupCs
    if ($LASTEXITCODE -eq 0 -and (Test-Path -LiteralPath $singleExe)) {
        Write-Host ('+ single exe: ' + $singleExe + '   ' + [math]::Round((Get-Item -LiteralPath $singleExe).Length/1MB,2) + ' MB')
    } else {
        Write-Host 'WARN: 单文件安装包编译失败（不影响其他产物）'
    }
    Remove-Item -LiteralPath $work -Recurse -Force
} else {
    Write-Host 'WARN: 未找到 csc 或 InkCanvasSetup.cs，跳过单文件安装包'
}

# Setup zip（payload + Setup.exe）
$sroot = Join-Path $r 'InkCanvas-v5.1.0-Setup'
$payload = Join-Path $sroot 'payload'
if (Test-Path -LiteralPath $sroot) { Remove-Item -LiteralPath $sroot -Recurse -Force }
New-Item -ItemType Directory -Path $payload -Force | Out-Null
Get-ChildItem -LiteralPath $stagedir -Force | Copy-Item -Destination $payload -Recurse -Force
$setupSrc = Join-Path $r 'SetupSource.exe'
if (-not (Test-Path -LiteralPath $setupSrc)) {
    Write-Host ('未找到 ' + $setupSrc + ' — 请先运行 rebuild-release-v5.ps1')
    exit 1
}
Copy-Item -LiteralPath $setupSrc -Destination (Join-Path $sroot 'Setup.exe') -Force
$zipS = Join-Path $r 'InkCanvas-v5.1.0-Setup.zip'
if (Test-Path -LiteralPath $zipS) { Remove-Item -LiteralPath $zipS -Force }
Compress-Archive -Path (Join-Path $sroot '*') -DestinationPath $zipS -CompressionLevel Optimal -Force
Write-Host ('+ setup:    ' + $zipS + '   ' + [math]::Round((Get-Item -LiteralPath $zipS).Length/1MB,2) + ' MB')

# 清理暂存目录（zip 已生成）
Remove-Item -LiteralPath $stageRoot -Recurse -Force
Remove-Item -LiteralPath $sroot -Recurse -Force

Write-Host ''
Write-Host '=== Releases ==='
Get-ChildItem -LiteralPath $r -File | Sort-Object LastWriteTime -Descending | ForEach-Object -Process `
{
    '{0}  {1,7:N2} MB  {2}' -f $_.LastWriteTime.ToString('HH:mm:ss'), ($_.Length/1MB), $_.Name
}
