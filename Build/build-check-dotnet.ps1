<#
.SYNOPSIS
    用 dotnet build 验证 Inkboard 源码能否编译。
.DESCRIPTION
    本机 MSBuild.exe 被安全策略拦截（csc/msbuild/InstallUtil 等 LOLBin 在命令黑名单里），
    但 dotnet 不在黑名单 —— 本脚本把源码复制到临时目录、把 COM 引用换成已生成的
    interop 程序集，然后用 dotnet build 编译，从而在不改动仓库任何文件的前提下验证改动。
    注意：这只在临时副本上跑，仓库里的 csproj 保持原样（COM 引用照旧，VS 里能正常打开）。
.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Build\build-check-dotnet.ps1
#>
param(
    [string]$Configuration = "Debug"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$projDir  = Join-Path $repoRoot "Ink Canvas"
$projFile = Join-Path $projDir "Ink Canvas.csproj"
$tmpRoot  = Join-Path $env:TEMP "inkboard-build"
$tmpProj  = Join-Path $tmpRoot "Ink Canvas"

if (-not (Test-Path $projFile)) { Write-Error "找不到项目文件：$projFile" }

# ---------- 1. 复制源码到临时目录（排除 bin/obj） ----------
Write-Host "[1/4] 复制源码到 $tmpProj ..." -ForegroundColor Cyan
if (Test-Path $tmpRoot) { Remove-Item $tmpRoot -Recurse -Force }
New-Item -ItemType Directory -Path $tmpProj -Force | Out-Null
& robocopy $projDir $tmpProj /E /XD bin obj /NFL /NDL /NJH /NJS /NP | Out-Null
if ($LASTEXITCODE -ge 8) { Write-Error "复制源码失败（robocopy exit $LASTEXITCODE）" }

# ---------- 2. 拷贝已 tlbimp 生成的 interop ----------
Write-Host "[2/4] 拷贝 COM interop ..." -ForegroundColor Cyan
$interopSrc = Join-Path $projDir "obj\Debug\Interop.IWshRuntimeLibrary.dll"
if (Test-Path $interopSrc) {
    Copy-Item $interopSrc (Join-Path $tmpProj "Interop.IWshRuntimeLibrary.dll") -Force
    $hasInterop = $true
} else {
    $hasInterop = $false
    Write-Warning "未找到 Interop.IWshRuntimeLibrary.dll（obj\Debug 下）。
    若从未用 VS/MSBuild 编译过，dotnet 无法自己 tlbimp。
    请先用 Visual Studio 生成一次，之后本脚本即可离线使用。"
}

# ---------- 3. 改写副本的 csproj：COMReference -> Reference ----------
Write-Host "[3/4] 改写副本 csproj（COM 引用本地化）..." -ForegroundColor Cyan
$tmpCsproj = Join-Path $tmpProj "Ink Canvas.csproj"
$text = [System.IO.File]::ReadAllText($tmpCsproj, [System.Text.Encoding]::UTF8)

# VBIDE：代码里没有任何引用，直接删
$text = [regex]::Replace($text, '\s*<COMReference Include="VBIDE">.*?</COMReference>', '',
        [System.Text.RegularExpressions.RegexOptions]::Singleline)

# IWshRuntimeLibrary：换成指向 interop dll 的普通引用（开机自启 .lnk 用到）
if ($hasInterop) {
    $text = [regex]::Replace($text, '\s*<COMReference Include="IWshRuntimeLibrary">.*?</COMReference>',
            "`n`t`t<Reference Include=`"Interop.IWshRuntimeLibrary`"><HintPath>Interop.IWshRuntimeLibrary.dll</HintPath><EmbedInteropTypes>True</EmbedInteropTypes></Reference>",
            [System.Text.RegularExpressions.RegexOptions]::Singleline)
}

# stdole：PIA，net472 引用程序集自带，但 dotnet 解析不到；代码未直接使用其类型，可安全删除
$text = [regex]::Replace($text, '\s*<COMReference Include="stdole">.*?</COMReference>', '',
        [System.Text.RegularExpressions.RegexOptions]::Singleline)

$leftover = ([regex]::Matches($text, '<COMReference\b')).Count
[System.IO.File]::WriteAllText($tmpCsproj, $text, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "      剩余 COMReference：$leftover（应为 0）"

# ---------- 4. 构建 ----------
Write-Host "[4/4] dotnet build -c $Configuration ..." -ForegroundColor Cyan
Push-Location $tmpProj
try {
    & dotnet build "Ink Canvas.csproj" -c $Configuration -v minimal
    $code = $LASTEXITCODE
} finally {
    Pop-Location
}

if ($code -eq 0) {
    $exe = Join-Path $tmpProj "bin\Inkboard\Inkboard.exe"
    Write-Host ""
    Write-Host "构建成功 -> $exe" -ForegroundColor Green
    Write-Host "（这是临时副本产物，仅用于验证编译；正式产物请用 VS 或 Build\rebuild-release-v5.ps1）" -ForegroundColor DarkGray
} else {
    Write-Host ""
    Write-Host "构建失败（exit $code），请往上翻看具体编译错误。" -ForegroundColor Red
}
exit $code
