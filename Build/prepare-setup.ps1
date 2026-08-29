# prepare-setup.ps1 - 构建安装包目录并输出 zip
param()
$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$rel = Join-Path $repo 'Releases'
$stage = Join-Path $rel 'stage-portable\Ink Canvas v2.1.0 x64'
$setupRoot = Join-Path $rel 'Ink Canvas Setup x64 v2.1.0'
$payload = Join-Path $setupRoot 'payload'
if (Test-Path $setupRoot) { Remove-Item -LiteralPath $setupRoot -Recurse -Force -ErrorAction Stop }
New-Item -ItemType Directory -Path $payload -Force | Out-Null
# 递归拷贝 stage -> payload
Get-ChildItem -LiteralPath $stage -File -Force | Copy-Item -Destination $payload -Force
Get-ChildItem -LiteralPath $stage -Directory -Force | ForEach-Object {
    $d = $_
    $dst = Join-Path $payload $d.Name
    New-Item -ItemType Directory -Path $dst -Force | Out-Null
    Get-ChildItem -LiteralPath $d.FullName -File -Recurse -Force | Copy-Item -Destination $dst -Force
}
# 安装程序放根目录（使用 ASCII 文件名，避免某些压缩 API 出现"路径有非法字符"）
Copy-Item -LiteralPath (Join-Path $rel 'InkCanvasSetup.exe') -Destination (Join-Path $setupRoot 'Setup.exe') -Force
Write-Host "Payload files: $((Get-ChildItem -LiteralPath $payload -Recurse -File -Force).Count)"
Get-ChildItem -LiteralPath $setupRoot | Format-Table Name, Length -AutoSize
# 打包 zip （用 PowerShell Compress-Archive 代替 ZipFile，避免中文/非法字符报错）
$zipSetup = Join-Path $rel 'InkCanvas-v2.1.0-x64-Setup.zip'
if (Test-Path -LiteralPath $zipSetup) { Remove-Item -LiteralPath $zipSetup -Force }
Compress-Archive -LiteralPath $setupRoot -DestinationPath $zipSetup -CompressionLevel Optimal -Force
$sz = Get-Item -LiteralPath $zipSetup -ErrorAction Stop
Write-Host ("SETUP ZIP OK  size={0:N2} MB  path={1}" -f ($sz.Length/1MB), $sz.FullName)
