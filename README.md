<div align="center">

[![LOGO](Ink%20Canvas/Resources/InkCanvas.png?raw=true "LOGO")](# "LOGO")

# Ink Canvas · 板书白板

[直接下载](https://github.com/XueRenYi0/ink-canvas-dev/releases/latest "Latest Releases") | [使用指南](Manual.md "说明和指南") | [开发说明](DEVELOPMENT.md "开发文档") | [常见问题](#-faq "FAQ")

[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](LICENSE) ![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg) ![Framework](https://img.shields.io/badge/.NET%20Framework-4.7.2%2B-purple.svg)

</div>

A fantastic Ink Canvas in WPF/C#, with fantastic support for Seewo Boards.

学校从传统投影仪换成了希沃白板，由于自带的"希沃白板"软件太难用，也没有同类好用的画板软件，所以有了该画板。

本仓库是 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas) 的优化分支（基于上游 master 分支），面向课堂板书场景持续增强，同样以 GPL-3.0 协议开源发布。当前版本 **v5.0.0**。

## ✨ 本分支主要增强

- **笔记上下滚动**：写满一屏不用清屏，屏幕右缘胶囊滚动条（上/下按钮 + 位置指示）继续向下书写；支持鼠标滚轮，PPT 放映中激活画布时同样可滚动
- **自定义图形图库**：选中任意墨迹一键"存入图库"，常驻图形面板"我的图形"行，随时原样落笔复用（ISF 存储，可右键删除）
- **选择工具增强**：选中墨迹支持放大/缩小/还原、15° 细步进旋转、一键克隆偏移；补齐了鼠标拖动选区路径（原版仅触摸实现）
- **可自定义快捷键**：画笔/橡皮/选择/撤销/重做/清空/图形/悬浮条/设置/退出共 10 个动作，右键菜单内查看与修改，即时生效
- **悬浮工具条优化**：启动默认收起为笑脸把手，单击展开、拖动移动，右键菜单快捷操作；内置防拖飞与卡死自愈看门狗
- **图形面板重构**：标题栏可拖动、图钉置顶、独立关闭按钮；图库缩略图自动换行对齐
- **多项稳定性修复**：修复 PPT 翻页时进程静默崩溃（原版遗留 bug）、退出后进程残留导致"误报已有实例"、悬浮条拖飞等
- **全新应用图标**：白底黑色马克笔极简风格，多尺寸自适应
- 代码结构重构：7600+ 行巨石 MainWindow 拆分为 20+ 个分部类文件，便于维护

## 🔧 基础特性（继承自上游）

对 Microsoft PowerPoint 有优化支持（强烈不推荐使用 WPS，会导致 WPS 自己把自己卡住，并且 WPS 对触摸屏的支持实在是差，PPT 翻页点击就行，而不是滑动，也不能放大缩小）
**笔细的一头写字，反过来粗的一头是橡皮擦。（希沃白板自己并不支持此功能）**
当然，用手直接擦也是可以的（跟希沃白板一样）
支持 Active Pen (支持压感)
对于其他红外线屏也可以提供相似功能，欢迎大家测试！

完整功能列表与用法见[使用指南](Manual.md)。

## 📦 下载与安装

前往 [Releases](https://github.com/XueRenYi0/ink-canvas-dev/releases/latest) 下载最新版本：

- `InkCanvas-v5.0.0-Setup.zip` — 安装版（推荐）
- `InkCanvas-v5.0.0-Portable.zip` — 免安装绿色版，解压即用

**系统要求**：Windows 10 及以上 · .NET Framework 4.7.2 或更高版本 · 需要使用 PPT 模式时请安装 Microsoft Office

## 🛠️ 从源码构建

```bash
msbuild "Ink Canvas.sln" /p:Configuration=Release /p:Platform=AnyCPU
```

或用 Visual Studio 2019/2022 直接打开解决方案生成。打包安装包/绿色版请使用 `Build/` 目录脚本，详见[开发说明](DEVELOPMENT.md)。

## ⚠️ 提示

- 提问前请先读 [FAQ](#-faq)
- 遇到问题请先尝试自行解决，若无法自行解决，请简单描述你的期望与现实的差异性。如果有必要，请附上复现此问题的操作步骤或错误日志¹ （可适当配图），等待回复。
- 对新功能的有效意见和合理建议，开发者会适时回复并进行开发。Ink Canvas 并非商业性质的软件，请勿催促开发者，耐心才能让功能更少 BUG、更加稳定。

> 等待是人类的一种智慧

 [1] ：对于长文本，可以使用在线剪贴板 （如 https://pastes.dev/ ），粘贴完毕点击 `SAVE` 后复制链接进行分享

## 📗 FAQ

### 在 Windows 10 以下版本系统中部分图标显示为 "□" 怎么办？
[点击下载](https://aka.ms/SegoeFonts "SegoeFonts") SegoeFonts 文件，安装压缩包中 `SegMDL2.ttf` 字体后重启即可解决

### 点击放映后一翻页就闪退？
考虑是由于`Microsoft Office`未激活导致的，请自行激活

### 放映后画板程序不会切换到PPT模式？
如果你曾经安装过`WPS`且在卸载后发现此问题则是由于暂时未确定的问题所导致，可以尝试重新安装WPS
> "您好，关于您反馈的情况我们已经反馈技术同学进一步分析哈，辛苦您可以留意后续WPS版本更新哈~" --回复自WPS客服

另外，处在保护（只读）模式的PPT不会被识别

### **安装后**程序无法正常启动？
请检查你的电脑上是否安装了 `.NET Framework 4.7.2` 或更高版本。若没有，请前往官网下载
如果仍无法运行，请检查你的电脑上是否安装了 `Microsoft Office`。若没有，请安装后重试

### 我该在何处提出功能需求和错误报告？

GitHub Issues：https://github.com/XueRenYi0/ink-canvas-dev/issues

### 大小屏设备交替使用/手指或笔头过大 导致被识别成橡皮怎么办？
点击画板的"设置"按钮并开启`特殊屏幕`选项即可

## 📁 文档

| 文档 | 内容 |
|---|---|
| [Manual.md](Manual.md) | 使用说明与技巧 |
| [DEVELOPMENT.md](DEVELOPMENT.md) | 代码结构、构建与打包、分支策略 |
| [privacy.txt](privacy.txt) | 隐私政策 |

## 感谢

- 本分支基于 [WXRIW/Ink-Canvas](https://github.com/WXRIW/Ink-Canvas) 开发，感谢上游作者 [WXRIW](https://github.com/WXRIW)
- 感谢 [yuwenhui2020](https://github.com/yuwenhui2020) 为 `Ink Canvas 使用说明` 做出的贡献！
- 感谢 [CN-Ironegg](https://github.com/CN-Ironegg)、[jiajiaxd](https://github.com/jiajiaxd)、[Kengwang](https://github.com/kengwang)、[Raspberry Kan](https://github.com/Raspberry-Monster) 为上游项目贡献代码！

## License

本项目基于 GPL-3.0 协议开源（继承自上游项目），详见 [LICENSE](LICENSE)。
