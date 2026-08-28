# Ink-Canvas-Dev 开发说明

本项目基于 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas)（上游提交 eafaf84）复制的个人优化分支。
原始软件是针对希沃白板和 PowerPoint 优化的轻量课堂画板（WPF, .NET Framework 4.7.2）。

## 项目现状

- **框架**：.NET Framework 4.7.2 WPF（保持不动，Win10/Win11 均可运行）
- **代码规模**：约 9000 行 C#，其中 MainWindow.xaml.cs 约 6866 行（巨石文件，后续逐步拆分）
- **主要依赖**：iNKORE UI WPF、Autoupdater.NET、Microsoft.Office.Interop.PowerPoint、
  内置墨迹识别库（IACore.dll / IALoader.dll / IAWinFX.dll / Microsoft.Ink.dll）

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
