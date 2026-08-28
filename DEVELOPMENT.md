# Ink-Canvas-Dev 开发说明

本项目基于 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas)（上游提交 eafaf84）复制的个人优化分支。
原始软件是针对希沃白板和 PowerPoint 优化的轻量课堂画板（WPF, .NET Framework 4.7.2）。

## 项目现状

- **框架**：.NET Framework 4.7.2 WPF（保持不动，Win10/Win11 均可运行）
- **主要依赖**：iNKORE UI WPF、Autoupdater.NET、Microsoft.Office.Interop.PowerPoint、
  内置墨迹识别库（IACore.dll / IALoader.dll / IAWinFX.dll / Microsoft.Ink.dll）
- **csproj 为 SDK 风格**：新增 .cs 文件自动包含进编译，无需手动登记

## 代码结构（MainWindow 已拆分为分部类）

原 MainWindow.xaml.cs 共 7662 行（巨石文件），已按原有 `#region` 边界
**纯机械移动**拆分为以下分部类文件（逻辑零改动，git 可追溯）：

| 文件 | 行数 | 职责 |
|---|---|---|
| MainWindow.xaml.cs | 47 | 类声明骨架（唯一声明 `: Window` 基类） |
| MW_Init.cs | 155 | 窗口初始化、定时器（PPT 检测/进程清理） |
| MW_InkCanvas.cs | 122 | 墨迹画布基础功能 |
| MW_Hotkeys.cs | 145 | 快捷键 |
| MW_TimeMachine.cs | 356 | 撤销重做 |
| MW_DefinitionsLoading.cs | 667 | 字段定义与初始化加载 |
| MW_RightPanel.cs | 679 | 右侧工具面板与颜色按钮 |
| MW_TouchEvents.cs | 439 | 触摸事件（含多指） |
| MW_PPT.cs | 703 | PowerPoint 放映交互 |
| MW_Settings.cs | 648 | 设置（行为/外观/自动化/手势等） |
| MW_LeftPanel.cs | 120 | 左侧面板与其他控件 |
| MW_SelectionGestures.cs | 509 | 墨迹选区与手势 |
| MW_ShapeDrawing.cs | 1565 | 图形绘制（形状按钮） |
| MW_WhiteboardControls.cs | 275 | 白板页面控制 |
| MW_SimulatePressure.cs | 641 | 压感模拟与墨迹转图形 |
| MW_MiscFunctions.cs | 402 | 杂项（自启/主题/截图/通知/工具） |
| MW_FloatBar.cs | 515 | 浮动工具栏（含拖动） |
| MW_SaveOpen.cs | 158 | 墨迹保存与打开 |
| PenPlugins.cs | 289 | 自定义 StylusPlugin 渲染器（命名空间级类，非 MainWindow 成员） |

注意：分部类共享同一类的全部字段，这只是**文件级拆分**（方便定位与修改），
类内耦合未降低。后续功能开发（如笔记滚动）的代码请放到对应职责的新分部文件
（如 `MW_Scroll.cs`），不要再往回堆。

## 构建方法

1. 用 Visual Studio 2019/2022 打开 `Ink Canvas.sln`
   （需安装“.NET 桌面开发”工作负载，含 .NET Framework 4.7.2 目标包）
2. 直接 F6 生成 / F5 运行，主项目为 `Ink Canvas`
3. 命令行构建：
   ```
   msbuild "Ink Canvas.sln" /p:Configuration=Release
   ```
   产物：`Ink Canvas\bin\Release\Ink Canvas.exe`

## 分支策略

| 分支 | 用途 |
|---|---|
| `master` | 与上游 eafaf84 完全一致，只读基线，不直接改动 |
| `dev-scroll` | 当前开发分支：笔记上下滚动功能 |

约定：每个功能/重构一步一个提交，提交前跑一遍人工验证清单。
出问题时用 `git checkout master` 可随时回到纯净基线。

## 计划功能：笔记上下滚动（画板模式）

**目标**：写满一屏后不清屏，向下滚动继续写，类似无限笔记。

**方案**：固定物理画布 + 坐标偏移（虚拟无限画布），不使用 ScrollViewer 包 InkCanvas
（ScrollViewer 会劫持笔/触摸事件，且墨迹坐标系会脱节）。

核心机制：
1. InkCanvas 物理尺寸始终等于屏幕，不动
2. 维护全局 `OffsetY`，滚动时对已有笔迹整体应用平移变换（矩阵变换，不重绘）
3. 新墨迹写入时换算为全局虚拟坐标存储
4. 滚动交互：屏幕右缘细滚动条 + 上/下翻页按钮（鼠标键盘触屏均可操作，不占用现有手势）
5. 橡皮/手势命中检测坐标统一经过换算入口
6. 滚动本身不进入 TimeMachine 撤销栈；仅墨迹操作可撤销
7. 仅画板模式启用；PPT 模式的"页"对应幻灯片，不适用本机制

## 人工验证清单（每步提交前过一遍）

- [ ] 画笔书写正常（含压感）
- [ ] 笔尾自动识别为橡皮
- [ ] 手势擦除正常
- [ ] 撤销/重做正常
- [ ] 清屏后可继续书写
- [ ] 画板模式与 PPT 模式切换正常
- [ ] 倒计时、随机点名窗口正常弹出
- [ ] 墨迹自动保存/恢复正常
