# Ink-Canvas-Dev 开发说明

本项目基于 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas) 的个人优化分支，以 GPL-3.0 协议开源。
原始软件是针对希沃白板和 PowerPoint 优化的轻量课堂画板（WPF, .NET Framework 4.7.2）。
v6.0.0 起产品更名为 **Inkboard**。

> 本文档面向接手者：**怎么编译、怎么打包、发版要同步哪些文件** 全部在"构建方法"与"发版流程"两节，
> 照着做即可，不依赖任何口口相传的知识。

## 项目现状

- **框架**：.NET Framework 4.7.2 WPF（保持不动，Win10/Win11 均可运行）
- **版本**：6.2.3（版本号分散在 5 处，发版必须同步，见"版本号同步清单"）
- **主要依赖**：iNKORE UI WPF、Autoupdater.NET、Microsoft.Office.Interop.PowerPoint、
  内置墨迹识别库（IACore.dll / IALoader.dll / IAWinFX.dll / Microsoft.Ink.dll）
- **csproj 为 SDK 风格**：新增 .cs 文件自动包含进编译，无需手动登记
- **输出统一**：Debug/Release 都输出到 `Ink Canvas\bin\Inkboard\`，主程序名 `Inkboard.exe`
  （v6.0.0 前叫 `Ink Canvas.exe`、输出在 `bin\Release\`——老文档/老脚本如还这么写就是过时的）

## 版本演进摘要（相对上游）

1. **v5.x**：MainWindow 巨石拆分部类、笔记上下滚动、选择工具增强、自定义图形图库、
   动态快捷键、悬浮条重构、PPT 崩溃修复、单文件安装包
2. **v6.0.x**：截图三件套（遮罩框选/掀板截取/本地图片）、笔/粗细面板、输入稳定性
   （停顿拉直、矩形触摸橡皮、Ctrl+Alt+Shift+R 安全重启）、清屏优化、Fluent 图标
3. **v6.1.0**：图片操作与墨迹统一（插入即选中、操作条、8 向手柄缩放、翻转/复制模式/
   适配白板宽度、ISF 剪贴板）、操作条精简 16→10 键、双色色点、笔/橡皮/选择方式独立面板、
   自动更新接入国内镜像回退链
4. **v6.2.2**：笔设置收敛（设置页隐藏画笔粗细/光标/笔锋/橡皮模式/橡皮大小/墨迹识别六项，
   改由笔面板统一承载，新增「墨迹识别 / 显示光标」开关区）、悬浮栏双指手势锁迁入快捷设置弹层、
   识别图形粗细统一（去掉硬编码假笔锋）、墨迹识别改「最少笔优先」减少多笔误合并
5. **v6.2.3**：「更多」面板改版（图标加汉字标签、新增「系统」组、去重）、主栏精简
   （设置/退出/白板底纹/多人书写收进更多面板）、快捷键对齐 ClassIn（Ctrl+P/E/M + Ctrl+A 全选）、
   滚动修复（位置按页记忆、选区跟随、选区内滚轮转发）、滚动胶囊按需浮现、检查更新全程有反馈

## 已知待优化项（后续接手时先看这里）

- **COM 框选识别（WPF 笔迹 → IInkDisp → LoadInk）失败**：转换桥在 64 位进程
  创建 COM 本体时疑似静默失败（InkObjCore.msinkaut.InkObject）。将来重拾时
  可试 Prefer32Bit 或换 CLSID 路径。
- **AI API 识别函数**：设想 = 框选笔迹 → RenderTargetBitmap 转图 → 视觉大模型
  返回表达式 → 现成 MathGraph 解析出图。设置页加 API 配置（OpenAI 兼容格式 +
  自填地址通吃各家）。体积零增长（HttpClient 自带），未动工。
- **图片操作不进撤销栈**：拖动/缩放/翻转后的恢复靠"统一还原"按钮（选中时快照），
  误操作不能 Ctrl+Z。将来若接 TimeMachine 需为图片设计历史条目结构。
- **旋转手柄对图片无效**：旋转仅支持墨迹（矩阵变换）；图片 RenderTransform 旋转
  不参与布局会导致选框不贴合，暂不支持。

## 代码结构（MainWindow 已拆分为分部类）

分部类共享同一类的全部字段，这是**文件级拆分**（方便定位与修改），
类内耦合未降低。后续新功能的代码请放到对应职责的新分部文件
（如滚动功能在 `MW_Scroll.cs`、图形库在 `MW_CustomShapes.cs`、
快捷键在 `MW_Shortcuts.Dynamic.cs`），不要再往回堆。

主要分部文件（完整清单见仓库 `Ink Canvas/` 目录，MW_ 前缀）：
`MW_Init.cs` 初始化 / `MW_Scroll.cs` 笔记滚动 / `MW_TimeMachine.cs` 撤销重做 /
`MW_SelectionGestures.cs` 选区与手势 / `MW_ShapeDrawing.cs` 图形绘制 /
`MW_CustomShapes.cs` 自定义图形图库 / `MW_TouchEvents.cs` 触摸 /
`MW_PPT.cs` PowerPoint 交互 / `MW_Settings.cs` 设置 / `MW_FloatBar.cs` 悬浮条 /
`MW_ImageLayer.cs` 页面图片层 / `MW_GraphStrokes.cs` 墨迹转图形 /
`MW_SimulatePressure.cs` 压感模拟 / `MW_MiscFunctions.cs` 杂项 /
`MW_Shortcuts.Dynamic.cs` 动态快捷键 + 全局逃生热键

## 构建方法

### 日常开发（编译菜单）

1. 用 Visual Studio 2019/2022 打开 `Ink Canvas.sln`
   （需安装".NET 桌面开发"工作负载，含 .NET Framework 4.7.2 目标包）
2. 直接 F6 生成 / F5 运行，主项目为 `Ink Canvas`
3. 命令行构建（本机 BuildTools 路径）：
   ```powershell
   & "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe" `
       "Ink Canvas\Ink Canvas.csproj" -p:Configuration=Debug
   ```
   产物：`Ink Canvas\bin\Inkboard\Inkboard.exe`（Debug/Release 同目录，靠 csproj 顶层 OutputPath 统一）

### 打包发布（打包菜单）

脚本位于 `Build/`，全部使用**仓库相对路径**，在任何位置检出均可运行：

| 脚本 | 作用 |
|---|---|
| `rebuild-release-v5.ps1` | 重编 Release（脚本内 `$ver` 是版本号来源之一），刷新 bin 下使用说明版本、写 `VersionInfo.ini`、打印 exe 版本自检。参数 `-msb` 可覆盖 MSBuild 路径（默认指向 VS2022 Community，本机是 BuildTools 需传参） |
| `build-zips.ps1` | 从 `bin\Inkboard` 生成 `Releases\Inkboard-vX.Y.Z-Portable.zip`（自动排除用户数据：Settings.json / Log.txt / CustomShapes / History Versions 等），并调 Inno Setup 编译出 `Inkboard-vX.Y.Z-Setup.exe`。打包内"使用说明 README.txt"取自 `Build\使用说明 README.txt` 模板，版本号由脚本正则自动刷新 |
| `InkCanvas.iss` | Inno Setup 安装脚本：**用户级安装**（`%LocalAppData%\Programs`，免 UAC——软件把运行数据写在 exe 目录，装 Program Files 会导致普通权限无法保存设置），正规开始菜单/桌面快捷方式与卸载项 |
| `verify-v5.ps1` | 校验 `Releases/` 下最新产物的版本一致性（zip 文件名 / exe AssemblyVersion / VersionInfo.ini 三处一致），不一致 exit 1 |
| `read-logs.ps1` | 本地调试用：查看各输出目录下 Log.txt 尾部 |

> 前置依赖：安装包需 [Inno Setup 6](https://jrsoftware.org/isinfo.php)（`winget install JRSoftware.InnoSetup`）。
> 中文语言包在 `Build/InnoLang/ChineseSimplified.isl`。

本机典型命令（仓库根目录 PowerShell）：

```powershell
powershell -ExecutionPolicy Bypass -File Build\rebuild-release-v5.ps1 -msb "C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\MSBuild.exe"
powershell -ExecutionPolicy Bypass -File Build\build-zips.ps1
powershell -ExecutionPolicy Bypass -File Build\verify-v5.ps1
```

产物在 `Releases/` 目录（**该目录不入库**，仅本地暂存，附件上传到 GitHub Releases）。

### 版本号同步清单（发版必查，共 5 处）

升版本时以下文件里的版本号**必须全部一致改掉**，漏一处会造成"自动更新不触发"或"校验失败"：

| # | 文件 | 字段 |
|---|---|---|
| 1 | `Ink Canvas\Properties\AssemblyInfo.cs` | `AssemblyVersion` / `FileVersion` / `AssemblyInformationalVersion`（三行一起） |
| 2 | `Build\InkCanvas.iss` | `#define MyAppVersion "X.Y.Z"` |
| 3 | `Build\rebuild-release-v5.ps1` | `$ver = 'X.Y.Z'` |
| 4 | `Build\build-zips.ps1` | `$ver = 'X.Y.Z'` |
| 5 | `update.xml`（仓库根） | `<version>` + `<url>`（新 Setup.exe 的镜像地址）+ `<changelog>` |

（打包内"使用说明"的版本由脚本正则自动刷，无需手改模板。）

## 发版流程（完整 Runbook）

以发布 vX.Y.Z 为例，从头到尾：

```
① 同步版本号   → 上面 5 处清单
② 编译打包     → rebuild-release-v5.ps1（-msb 传本机路径）→ build-zips.ps1 → verify-v5.ps1
③ 提交推送     → git add 相关文件 → commit → git push origin master
④ 发 GitHub Release → 建标签 + 传附件（curl 方式，见下）
⑤ 更新 update.xml → 版本号 + changelog + 新 Setup.exe 镜像 URL → 推送 master
   （⑤必须在④之后：update.xml 指向的 Release 附件要先存在）
```

### GitHub Release 创建（curl 方式）

`release-artifacts/` 下有历次发版脚本（v6.1.0 为 `create-release-v610.ps1`，含正文模板）。
**注意用 curl.exe 而非 Invoke-RestMethod**——PS 5.1 发长中文 JSON 会 400（空响应体，难排查）。
要点：

```powershell
# token 从 git 凭据取（已配置过 git credential）
$cred = "protocol=https`nhost=github.com`n" | git credential fill
$token = ($cred | Select-String '^password=(.+)$').Matches.Groups[1].Value

# 建 Release（正文写进 json 文件，curl --data-binary 上传）
curl.exe -s -x http://127.0.0.1:7890 -X POST `
  "https://api.github.com/repos/XueRenYi0/ink-canvas-dev/releases" `
  -H "Authorization: token $token" -H "Content-Type: application/json" `
  --data-binary "@release-body.json"

# 传附件（zip 用 application/zip，exe 用 application/octet-stream）
curl.exe -s -x http://127.0.0.1:7890 -X POST `
  "https://uploads.github.com/repos/XueRenYi0/ink-canvas-dev/releases/<id>/assets?name=Inkboard-vX.Y.Z-Setup.exe" `
  -H "Authorization: token $token" -H "Content-Type: application/octet-stream" `
  --data-binary "@Releases\Inkboard-vX.Y.Z-Setup.exe"
```

### 发版环境坑（都踩过，别再踩）

- **仓库默认分支是 `dev-scroll`，但实际开发/发版分支是 `master`**：GitHub 仓库设置里
  默认分支还停在旧的 dev-scroll。创建 Release 必须显式 `target_commitish=master`，
  否则 tag 打不到你推的提交上。（建议：直接去 GitHub 设置把默认分支改成 master，一劳永逸）
- **网络走代理 `http://127.0.0.1:7890`**：curl 加 `-x`，Invoke-RestMethod 加 `-Proxy`。
- **.ps1 脚本必须带 UTF-8 BOM**：无 BOM 时 PowerShell 5.1 把中文当 GBK 解析直接报错。
  新建脚本后用 `release-artifacts` 里现成脚本的方式补 BOM。
- **PowerShell 5.1 不支持 heredoc 和三元运算符**：git commit 多行信息先写进 txt 文件再
  `git commit -F file.txt`；`? :` 改 if/else。
- **镜像缓存延迟**：gh-proxy 对 raw 文件有短暂负缓存，推送后立即测 404 是正常的，等 1-2 分钟。

## 自动更新机制（v6.1.0 起）

- **配置文件**：`update.xml`（仓库根，随 master 推送生效，无需任何服务器）
- **客户端链路**：`App.xaml.cs` `StartUpdateWatcher()`——启动延迟 5 秒静默检查
  （设置页"启动时检查更新"开关可关），右键悬浮条笑脸 → "检查更新"手动触发
- **国内加速**：update.xml 和安装包 URL 都走 `gh-proxy.com` 镜像；客户端拉 XML 有三级
  回退链 `gh-proxy.com → gh-proxy.at9.net → GitHub 直连`，镜像全挂也能慢速更新
- **发版动作**：就是上面 Runbook 的第 ⑤ 步，改 3 个字段推上去即可

## 分支策略

| 分支 | 用途 |
|---|---|
| `master` | **实际开发与发版分支**（所有功能提交、Release tag 都在这） |
| `dev-scroll` | 历史遗留：GitHub 仓库的默认分支但已停更（见"发版环境坑"第一条） |

约定：每个功能/重构一步一个提交，提交前跑一遍人工验证清单。

### 远程仓库配置

```bash
# 原上游远程改名保留，方便日后比对/同步上游
git remote rename origin upstream
# 本仓库作为 origin
git remote add origin https://github.com/XueRenYi0/ink-canvas-dev.git
```

## 人工验证清单（发版前过一遍）

- [ ] 画笔书写正常（含压感、荧光笔/激光笔切换）
- [ ] 撤销/重做正常；清屏后可继续书写
- [ ] 画板模式与 PPT 模式切换正常
- [ ] 笔记滚动：胶囊条/滚轮正常，PPT 放映中激活画布可滚
- [ ] 图片：插入即选中、拖动（含中央）、8 向手柄缩放、翻转、适配宽度、
      复制模式拖副本、删除后操作条消失、统一还原
- [ ] 混合选中（墨迹+图片）：拖动/缩放同步，操作条按钮全生效
- [ ] 选中状态下：截图 PNG 干净无选框；点图形/笔/橡皮自动取消选中
- [ ] 图形库：存入、插入、右键删除正常
- [ ] 快捷键：默认可用，自定义即时生效，重启保留
- [ ] 悬浮条：笑脸展开/收起/拖动/右键菜单（含检查更新）正常
- [ ] 自动更新：update.xml 版本 > 本地时启动 5 秒内弹窗；设置开关可关闭
- [ ] 打包三连（rebuild → zips → verify）通过，版本三处一致
