<#
    Inkboard 一键快照 / 回退工具
    ------------------------------------------------------------
    存档（每做完一小步就存一次，改坏了随时能退回来）：
        powershell -File Build\snapshot.ps1 -Msg "做完了穿透状态指示"

    列出所有分步快照：
        powershell -File Build\snapshot.ps1 -List

    回退到某个快照（只动被 git 跟踪的文件，未跟踪文件不受影响）：
        powershell -File Build\snapshot.ps1 -Back snap-0909-1930
#>
param(
    [string]$Msg = "",
    [switch]$List,
    [string]$Back = ""
)

$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

# ---- 列出快照 ----
if ($List) {
    Write-Host ""
    Write-Host "=== 分步快照（snap-*）===" -ForegroundColor Cyan
    $tags = git tag -l "snap-*" --sort=-creatordate
    if (-not $tags) { Write-Host "  （还没有分步快照）" }
    else {
        $tags | ForEach-Object {
            $info = git log -1 --pretty="%ad  %s" --date=format:"%m-%d %H:%M" $_
            Write-Host ("  {0}   {1}" -f $_, $info)
        }
    }
    Write-Host ""
    Write-Host ""
    Write-Host "=== 固定存档点 ===" -ForegroundColor Cyan
    Write-Host "  backup-20260909-ui   悬浮栏改造的完整状态（回退之前）"
    git log -1 --pretty="  %h   %s" 8e2017e
    exit 0
}

# ---- 回退 ----
if ($Back -ne "") {
    Write-Host ""
    Write-Host "即将把源码回退到快照：$Back" -ForegroundColor Yellow
    Write-Host "说明：只影响 git 跟踪过的文件；未跟踪的文件不会被删。" -ForegroundColor Yellow
    $c = Read-Host "确认请输入 Y（其他任意键取消）"
    if ($c -ne "Y") { Write-Host "已取消，什么都没改。"; exit 0 }
    git checkout $Back -- .
    Write-Host ""
    Write-Host "工作区已回退到 $Back（尚未提交，你可以先跑起来看效果）。" -ForegroundColor Green
    Write-Host "确认没问题后再执行：  git commit -m ""回退到 $Back"""
    exit 0
}

# ---- 存档 ----
if ($Msg -eq "") { $Msg = Read-Host "给这步改动写一句说明（例如：加好了穿透状态指示）" }
if ($Msg -eq "") { $Msg = "快照" }

git add -A -- "Ink Canvas" Build *.md
$staged = git diff --cached --name-only
if (-not $staged) { Write-Host "没有需要存档的改动，当前已经是最新状态。"; exit 0 }

$tag = "snap-" + (Get-Date -Format "MMdd-HHmm")
git commit -q -m $Msg
git tag $tag
Write-Host ""
Write-Host "已存档：" -ForegroundColor Green
git log -1 --pretty="  %h   %s"
Write-Host "  标签：$tag"
Write-Host ("  想退回这一步：powershell -File Build\snapshot.ps1 -Back " + $tag)
