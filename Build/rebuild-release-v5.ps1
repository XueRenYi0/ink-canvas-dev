param(
    [string]$msb = 'C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe',
    [string]$csproj = '',
    [string]$csc = 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe',
    [string]$src = '',
    [string]$releases = ''
)
# 未显式传参时，路径相对本脚本（Build/）与仓库根解析，任意检出位置可用
$repo = Split-Path -Parent $PSScriptRoot
if (-not $csproj)   { $csproj   = Join-Path $repo 'Ink Canvas\Ink Canvas.csproj' }
if (-not $src)      { $src      = Join-Path $PSScriptRoot 'InkCanvasSetup.cs' }
if (-not $releases) { $releases = Join-Path $repo 'Releases' }
$ErrorActionPreference = 'Continue'
$out = Split-Path $csproj -Parent ; $out = Join-Path $out 'bin\Release'
Write-Host ('--- MSBuild Rebuild Release (AnyCPU 32位首选，与 VS Debug 配置一致) ---')
& $msb $csproj /t:Rebuild /p:Configuration=Release /p:Platform=AnyCPU /v:minimal
Write-Host ('MSBuild exit: ' + $LASTEXITCODE)

# 版本 README 刷新
$readme = Join-Path $out '使用说明 README.txt'
if (Test-Path -LiteralPath $readme) {
    $txt = [IO.File]::ReadAllText($readme)
    $txt = $txt.Replace('2.1.0', '5.1.0').Replace('2.1.2026.0829', '5.1.2026.0829').Replace('5.0.0', '5.1.0').Replace('5.0.2026.0829', '5.1.2026.0829')
    [IO.File]::WriteAllText($readme, $txt, [Text.UTF8Encoding]::new($false))
    Write-Host 'README updated to 5.1.0'
}

# VersionInfo.ini 写回 5.1.0（Rebuild 可能清掉 bin\Release）
[IO.File]::WriteAllText((Join-Path $out 'VersionInfo.ini'), '5.1.0', [Text.ASCIIEncoding]::new())

# 校验 exe 版本
$exe = Join-Path $out 'Ink Canvas.exe'
$fvi = [Diagnostics.FileVersionInfo]::GetVersionInfo($exe)
$asmV = [Reflection.AssemblyName]::GetAssemblyName($exe).Version
Write-Host ('AssemblyVersion:  ' + $asmV)
Write-Host ('FileVersion:      ' + $fvi.FileVersion)
Write-Host ('ProductVersion:   ' + $fvi.ProductVersion)
Write-Host ('VersionInfo.ini:  ' + [IO.File]::ReadAllText((Join-Path $out 'VersionInfo.ini')).Trim())

# 重编 Setup.exe（v5 常量）
$setupOut = Join-Path $releases 'SetupSource.exe'
New-Item -ItemType Directory -Path $releases -Force | Out-Null
if (Test-Path -LiteralPath $setupOut) { Remove-Item -LiteralPath $setupOut -Force }
& $csc /nologo /target:winexe /out:$setupOut /r:System.dll /r:System.Core.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.IO.Compression.dll /r:System.IO.Compression.FileSystem.dll $src
Write-Host ('SetupSource exists: ' + (Test-Path -LiteralPath $setupOut))
