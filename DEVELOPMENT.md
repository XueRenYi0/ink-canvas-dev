# Ink-Canvas-Dev 开发说明

本项目基于 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas) 的个人优化分支，以 GPL-3.0 协议开源。
原始软件是针对希沃白板和 PowerPoint 优化的轻量课堂画板（WPF, .NET Framework 4.7.2）。

## 项目现状

- **框架**：.NET Framework 4.7.2 WPF（保持不动，Win10/Win11 均可运行）
- **版本**：5.0.0（版本号统一升到 5.0.0，避免客户端自动更新误报，见 `Properties/AssemblyInfo.cs`）
- **主要依赖**：iNKORE UI WPF、Autoupdater.NET、Microsoft.Office.Interop.PowerPoint、
  内置墨迹识别库（IACore.dll / IALoader.dll / IAWinFX.dll / Microsoft.Ink.dll）
- **csproj 为 SDK 风格**：新增 .cs 文件自动包含进编译，无需手动登记

## 已完成的主要改动（相对上游）

按提交顺序大致分为：

1. **重构**：MainWindow.xaml.cs（7662 行巨石文件）按 `#region` 边界纯机械拆分为
   20+ 个分部类文件（逻辑零改动，git 可追溯）
2. **笔记上下滚动**：右缘胶囊滚动条 + 位置指示 + 鼠标滚轮 + PPT 放映中可滚
3. **选择工具增强**：缩放组（放大/缩小/还原）、15° 细步进旋转、克隆重做、鼠标选区路径
4. **自定义图形图库**：选区墨迹存为 ISF 图形、图形面板"我的图形"行、右键删除、toast 反馈
5. **动态快捷键系统**：WPF KeyBinding（窗口焦点时生效），10 个动作可自定义，存 Settings.json
6. **悬浮条优化**：启动收起为笑脸把手、右键菜单、防拖飞看门狗与卡死自愈
7. **稳定性修复**：PPT 翻页 COM 后台线程调用补 try-catch（原版静默崩溃）、
   退出后单实例 mutex 立即释放、右键菜单拖拽卡死三层防御
8. **图形面板重构**：根层级定位、标题栏拖动、图钉/关闭按钮、图库换行
9. **新应用图标**：白底黑色马克笔斜置极简风（W4 方案，源图与生成脚本在 `Resources/`）

## 已知待优化项（后续接手时先看这里）

- ~~**双曲线渐近线"拆段式虚线"难整条擦除**~~（已解决，2026-08-30）：
  原定方案"单条笔迹 + `OutlineStrokes()` + `DashArray`"经技术推演**不可行**——
  WPF 墨迹模型中一条 Stroke 是连续笔迹（笔尖扫过必留墨），无法渲染出断开的虚线间隙。
  实际采用的方案：**整组擦除**（`InkCanvas_StrokeErasing`，MW_GraphStrokes.cs）。
  橡皮（按笔迹擦除模式）碰到图形组内任意一条笔迹（含渐近线的任意一小段虚线）时，
  拦截单条擦除，改为按 `GraphGroupGuid` 组标签一次性移除整组（渐近线 + 双曲线本体 + 焦点）。
  虚线外观不变，撤销（Ctrl+Z）整组恢复；点擦除模式不干预，保留精细雕刻语义；
  普通手写笔迹无标签，行为完全不变。
- **COM 框选识别（WPF 笔迹 → IInkDisp → LoadInk）失败**：转换桥在 64 位进程
  创建 COM 本体时疑似静默失败（InkObjCore.msinkaut.InkObject）。将来重拾时
  可试 Prefer32Bit 或换 CLSID 路径。
- **AI API 识别函数**：设想 = 框选笔迹 → RenderTargetBitmap 转图 → 视觉大模型
  返回表达式 → 现成 MathGraph 解析出图。设置页加 API 配置（OpenAI 兼容格式 +
  自填地址通吃各家）。体积零增长（HttpClient 自带），未动工。

## 代码结构（MainWindow 已拆分为分部类）

分部类共享同一类的全部字段，这是**文件级拆分**（方便定位与修改），
类内耦合未降低。后续新功能的代码请放到对应职责的新分部文件
（如滚动功能在 `MW_Scroll.cs`、图形库在 `MW_CustomShapes.cs`、
快捷键在 `MW_Shortcuts.Dynamic.cs`），不要再往回堆。

| 文件 | 行数 | 职责 |
|---|---|---|
| MainWindow.xaml.cs | 45 | 类声明骨架（唯一声明 `: Window` 基类） |
| MW_Init.cs | 144 | 窗口初始化、定时器（PPT 检测/进程清理） |
| MW_InkCanvas.cs | 125 | 墨迹画布基础功能 |
| MW_Hotkeys.cs | 116 | 快捷键 |
| MW_Scroll.cs | 146 | 笔记上下滚动（右缘滚动条） |
| MW_TimeMachine.cs | 342 | 撤销重做 |
| MW_DefinitionsLoading.cs | 623 | 字段定义与初始化加载 |
| MW_RightPanel.cs | 643 | 右侧工具面板与颜色按钮 |
| MW_TouchEvents.cs | 401 | 触摸事件（含多指） |
| MW_PPT.cs | 646 | PowerPoint 放映交互 |
| MW_Settings.cs | 528 | 设置（行为/外观/自动化/手势等） |
| MW_LeftPanel.cs | 111 | 左侧面板与其他控件 |
| MW_SelectionGestures.cs | 603 | 墨迹选区与手势（含鼠标选区路径） |
| MW_ShapeDrawing.cs | 1502 | 图形绘制（形状按钮） |
| MW_CustomShapes.cs | 305 | 自定义图形图库 |
| MW_Shortcuts.Dynamic.cs | 224 | 动态可自定义快捷键系统 |
| MW_WhiteboardControls.cs | 239 | 白板页面控制 |
| MW_SimulatePressure.cs | 597 | 压感模拟与墨迹转图形 |
| MW_MiscFunctions.cs | 347 | 杂项（自启/主题/截图/通知/工具） |
| MW_FloatBar.cs | 501 | 悬浮工具栏（含拖动/看门狗） |
| MW_SaveOpen.cs | 152 | 墨迹保存与打开 |
| PenPlugins.cs | 243 | 自定义 StylusPlugin 渲染器（命名空间级类，非 MainWindow 成员） |

## 构建方法

### 日常开发

1. 用 Visual Studio 2019/2022 打开 `Ink Canvas.sln`
   （需安装“.NET 桌面开发”工作负载，含 .NET Framework 4.7.2 目标包）
2. 直接 F6 生成 / F5 运行，主项目为 `Ink Canvas`
3. 命令行构建：
   ```
   msbuild "Ink Canvas.sln" /p:Configuration=Release
   ```
   产物：`Ink Canvas\bin\Release\Ink Canvas.exe`

### 打包发布（安装包 + 绿色版）

脚本位于 `Build/`，全部使用**仓库相对路径**，在任何位置检出均可运行
（MSBuild / csc 的默认路径为 VS2022 社区版标准位置，可用参数覆盖）：

| 脚本 | 作用 |
|---|---|
| `rebuild-release-v5.ps1` | 重新编译 Release（AnyCPU 32 位首选，与上游运行时行为一致），刷新版本号文件 |
| `build-zips.ps1` | 从 `bin\Release` 生成 `InkCanvas-v5.1.0-Portable.zip`（自动排除用户数据：Settings.json/Log.txt 等），并调用 Inno Setup 编译出单文件安装包 `InkCanvas-v5.1.0-Setup.exe` |
| `InkCanvas.iss` | Inno Setup 安装脚本：**用户级安装**（`%LocalAppData%\Programs`，免 UAC——软件把运行数据写在 exe 目录，装 Program Files 会导致普通权限无法保存设置），正规开始菜单/桌面快捷方式与卸载项 |
| `verify-v5.ps1` | 解包两个 zip 校验版本号一致性（AssemblyVersion / VersionInfo.ini / README） |
| `read-logs.ps1` | 本地调试用：查看 Release/Debug/Releases 下的 Log.txt 尾部 |

> 前置依赖：构建安装包需安装 [Inno Setup 6](https://jrsoftware.org/isinfo.php)（`winget install JRSoftware.InnoSetup`）。中文语言包在 `Build/InnoLang/ChineseSimplified.isl`（来自 kira-96 的 Inno-Setup-Chinese-Simplified-Translation，MIT）。旧 csc 自制安装器（`InkCanvasSetup.cs`、`package-v5.ps1`、`prepare-setup.ps1`）已删除。

典型发布流程（在仓库根目录的 PowerShell 中执行）：

```powershell
powershell -ExecutionPolicy Bypass -File Build\rebuild-release-v5.ps1
powershell -ExecutionPolicy Bypass -File Build\build-zips.ps1
powershell -ExecutionPolicy Bypass -File Build\verify-v5.ps1
```

产物在 `Releases/` 目录（该目录不入库，zip 上传到 GitHub Releases）。

## 分支策略

| 分支 | 用途 |
|---|---|
| `dev-scroll` | 开发与发布主分支（公开仓库的默认分支），CI 也触发于此 |
| `master` | 与上游基线一致，只读参照，不直接改动 |

约定：每个功能/重构一步一个提交，提交前跑一遍人工验证清单。
出问题时用 `git checkout master` 可随时回到纯净基线。

### 远程仓库配置

```bash
# 原上游远程改名保留，方便日后比对/同步上游
git remote rename origin upstream
# 本仓库作为 origin
git remote add origin https://github.com/XueRenYi0/ink-canvas-dev.git
```

发布：`git push -u origin dev-scroll`，并在 GitHub 仓库设置中将默认分支设为 `dev-scroll`。

## 笔记滚动功能设计（已实现）

**目标**：写满一屏后不清屏，向下滚动继续写，类似无限笔记。

**方案**：固定物理画布 + 坐标偏移（虚拟无限画布），不使用 ScrollViewer 包 InkCanvas
（ScrollViewer 会劫持笔/触摸事件，且墨迹坐标系会脱节）。

核心机制：
1. InkCanvas 物理尺寸始终等于屏幕，不动
2. 维护全局 `OffsetY`，滚动时对已有笔迹整体应用平移变换（矩阵变换，不重绘）
3. 新墨迹写入时换算为全局虚拟坐标存储
4. 滚动交互：屏幕右缘胶囊滚动条（上下按钮 + 位置指示条），支持鼠标滚轮
5. 橡皮/手势命中检测坐标统一经过换算入口
6. 滚动本身不进入 TimeMachine 撤销栈；仅墨迹操作可撤销
7. 画板模式与屏幕注释模式启用；PPT 模式的"页"对应幻灯片，不适用本机制
   （PPT 放映中激活画布时仍可滚动画布墨迹）

## 人工验证清单（每步提交前过一遍）

- [ ] 画笔书写正常（含压感）
- [ ] 笔尾自动识别为橡皮
- [ ] 手势擦除正常
- [ ] 撤销/重做正常
- [ ] 清屏后可继续书写
- [ ] 画板模式与 PPT 模式切换正常
- [ ] 倒计时、随机点名窗口正常弹出
- [ ] 墨迹自动保存/恢复正常
- [ ] 笔记滚动：按钮/滚轮/位置指示正常，PPT 放映中激活画布可滚
- [ ] 图形库：存入、插入、右键删除、toast 提示正常
- [ ] 快捷键：默认可用，自定义修改即时生效，重启后保留
- [ ] 悬浮条：笑脸展开/收起/拖动/右键菜单正常
- [ ] 打包脚本三连（rebuild → zips → verify）通过
