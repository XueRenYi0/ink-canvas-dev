param(
    [string]$msb = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe',
    [string]$csproj = '',
    [string]$releases = ''
)
# 未显式传参时，路径相对本脚本（Build/）与仓库根解析，任意检出位置可用
# 旧 csc 自制安装器（SetupSource.exe）已废弃：安装包改由 build-zips.ps1 调用 Inno Setup 编译
$repo = Split-Path -Parent $PSScriptRoot
if (-not $csproj)   { $csproj   = Join-Path $repo 'Ink Canvas\Ink Canvas.csproj' }
if (-not $releases) { $releases = Join-Path $repo 'Releases' }
$ErrorActionPreference = 'Continue'
# 统一输出目录（csproj 已把 Debug/Release 都指向 bin\InkClass，见 Ink Canvas.csproj 顶层 OutputPath）
$out = Split-Path $csproj -Parent ; $out = Join-Path $out 'bin\InkClass'
Write-Host ('--- MSBuild Rebuild Release (AnyCPU 32位首选，与 VS Debug 配置一致) ---')
& $msb $csproj /t:Rebuild /p:Configuration=Release /p:Platform=AnyCPU /v:minimal
Write-Host ('MSBuild exit: ' + $LASTEXITCODE)

# 版本号（与 AssemblyInfo.cs / InkCanvas.iss / build-zips.ps1 保持一致，升级时四处同步改）
$ver = '6.6.2'

# 版本 README 刷新（正则自动替换任意历史版本号，发版无需手改模板）
$readme = Join-Path $out '使用说明 README.txt'
if (Test-Path -LiteralPath $readme) {
    $txt = [IO.File]::ReadAllText($readme)
    $txt = [regex]::Replace($txt, 'v\d+\.\d+\.\d+', ('v' + $ver))
    $txt = [regex]::Replace($txt, '\d+\.\d+\.\d{4}\.\d{4}', ($ver + '.2026.' + (Get-Date -Format 'MMdd')))
    [IO.File]::WriteAllText($readme, $txt, [Text.UTF8Encoding]::new($false))
    Write-Host ('README updated to ' + $ver)
}

# VersionInfo.ini 写回（Rebuild 可能清掉 bin\Release）
[IO.File]::WriteAllText((Join-Path $out 'VersionInfo.ini'), $ver, [Text.ASCIIEncoding]::new())

# 校验 exe 版本（v6.0.0 起 exe 更名为 InkClass.exe）
$exe = Join-Path $out 'InkClass.exe'
$fvi = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
$asmV = [Reflection.AssemblyName]::GetAssemblyName($exe).Version
Write-Host ('AssemblyVersion:  ' + $asmV)
Write-Host ('FileVersion:      ' + $fvi.FileVersion)
Write-Host ('ProductVersion:   ' + $fvi.ProductVersion)
Write-Host ('VersionInfo.ini:  ' + [IO.File]::ReadAllText((Join-Path $out 'VersionInfo.ini')).Trim())

Write-Host ''
Write-Host 'Release 编译完成。后续：运行 build-zips.ps1 生成 Portable.zip 与 Inno Setup 安装包'
