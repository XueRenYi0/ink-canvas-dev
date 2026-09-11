using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：笔设置面板（PenSettingsPanel）的弹出/定位/关闭，
    /// 以及笔种类（画笔/荧光笔/激光笔）的实现。
    ///
    /// 交互模型（用户需求）：
    /// - 点笔图标 → 弹出面板（笔种类/粗细/颜色/笔锋四区）
    /// - 面板内点选不关面板（可连续调整）
    /// - 点面板外任意处 → 面板消失（判定见 MW_PopupLayers.cs，本面板已登记在表中）
    /// </summary>
    public partial class MainWindow
    {
        #region 笔种类状态

        /// <summary>笔种类：0=画笔 1=荧光笔 2=激光笔</summary>
        private enum PenType { Pen = 0, Highlighter = 1, Laser = 2 }

        /// <summary>当前笔种类（持久化于 Settings.Canvas.PenType）</summary>
        private PenType currentPenType = PenType.Pen;

        /// <summary>
        /// 画笔逻辑笔宽（与荧光笔各自独立记忆；设置页滑条/面板粗细档共用）。
        /// 激光笔也用此宽度（激光笔迹本质是"会淡出的临时画笔"）。
        /// </summary>
        private double _penWidth = 2.5;

        /// <summary>荧光笔笔宽（独立记忆：荧光笔是"高亮宽笔带"，量级本来就比画笔大）</summary>
        private double _highlighterWidth = 14;

        /// <summary>画笔粗细五档（面板粗细区从左到右）</summary>
        private static readonly double[] PenWidthPresets = { 1.5, 2.5, 4, 6, 9 };

        /// <summary>荧光笔粗细五档（荧光带需要足够宽度盖住文字）</summary>
        private static readonly double[] HighlighterWidthPresets = { 6, 10, 14, 20, 26 };

        /// <summary>
        /// 扩展三色（面板第 7~9 个色点）：橙/品红/青——纯色色相网格的补充色。
        /// 七色体系 = 纯色 hue 网格（红0°/橙30°/黄60°/绿120°/青180°/蓝~220°/品红~292°），
        /// 品红替代传统紫：hue 300° 与蓝、红各约隔 60°，色轮上完美对称不挤。
        /// 前 6 色（黑红绿蓝黄白）沿用右侧面板色块（支持 Colors\*.ini 自定义）。
        /// </summary>
        private static readonly Color[] ExtraColors =
        {
            Color.FromRgb(255, 127, 0),   // 橙（#FF7F00 纯色 hue 30°：插在红黄之间的补充色，亮度居中与两者拉开）
            Color.FromRgb(224, 64, 251),  // 品红（#E040FB 紫红：纯色 #FF00FF 的可读版——纯品红荧光感过强，略降饱和）
            Color.FromRgb(0, 213, 213)    // 青（#00D5D5 hue 180°：纯青 #00FFFF 白底看不清，降亮度到 L≈42% 兼顾黑白板）
        };

        /// <summary>
        /// 激光笔笔迹标记（DrawingAttributes 扩展属性）。
        /// 写入 DefaultDrawingAttributes 后，收集到的笔迹会带上此标记：
        /// - StrokesOnStrokesChanged 据此跳过 TimeMachine 提交（临时笔迹不进撤销栈）
        /// - inkCanvas_StrokeCollected 据此启动淡出计时（不走墨迹转图形/笔锋）
        /// </summary>
        private static readonly Guid LaserStrokeGuid = new Guid("7C1E9A34-5D2B-4E8F-A6C0-9B7D3E5F1024");

        #endregion

        #region 面板初始化与事件接线

        /// <summary>初始化笔设置面板：订阅面板事件、注入颜色（MainWindow 构造后调用一次）</summary>
        private void InitPenSettingsPanel()
        {
            // 面板事件 → 主窗口动作（面板自身不碰 inkCanvas/Settings，保持解耦）
            PenSettingsPanel.PenTypeSelected += t => SetPenType((PenType)t);
            PenSettingsPanel.ThicknessSelected += i => SetPenThicknessByIndex(i);
            PenSettingsPanel.ColorSelected += i => SelectPenColorByIndex(i);
            PenSettingsPanel.TaperSelected += s => SetPenTaperStyle(s);
            // 自定义格长按/右键 = 重新打开取色板换色（左键是"使用颜色"，见 SelectCustomColorSmart）
            PenSettingsPanel.CustomColorReselectRequested += SelectCustomColor;
            // 开关区（原设置面板「墨迹识别」组 + 「画板」组的两个开关，设置里已隐藏，改由笔面板承载）
            PenSettingsPanel.InkToShapeToggled += ToggleInkToShape;
            PenSettingsPanel.CursorToggled += ToggleShowCursor;

            // 注入 9 色：前 5 色来自右侧面板色块（支持 Colors\*.ini 自定义配色），
            // 白为固定白，后 3 色（橙/紫/青）为固定扩展色
            PenSettingsPanel.SetColorBrushes(new Brush[]
            {
                BtnColorBlack.Background, BtnColorRed.Background, BtnColorGreen.Background,
                BtnColorBlue.Background, BtnColorYellow.Background,
                new SolidColorBrush(StringToColor("#FFFEFEFE")),
                new SolidColorBrush(ExtraColors[0]),
                new SolidColorBrush(ExtraColors[1]),
                new SolidColorBrush(ExtraColors[2])
            });

            // 自定义色格（未设置过则显示虚线"＋"）
            var custom = GetSavedCustomColor();
            PenSettingsPanel.SetCustomColorBrush(custom.HasValue ? new SolidColorBrush(custom.Value) : null);
        }

        /// <summary>设置加载完成后恢复笔种类/笔宽（MW_DefinitionsLoading 的 Canvas 分支末尾调用）</summary>
        private void ApplyLoadedPenSettings()
        {
            // 笔种类：越界值（旧配置/手改 json）兜底为画笔
            int t = Settings.Canvas.PenType;
            currentPenType = (t >= 0 && t <= 2) ? (PenType)t : PenType.Pen;

            // 画笔/荧光笔各自的笔宽来自配置（滑条值的一半，最小 0.5）
            _penWidth = Math.Max(0.5, Settings.Canvas.InkWidth);
            _highlighterWidth = Math.Max(2, Settings.Canvas.HighlighterWidth);

            // 自定义色格在此恢复（InitPenSettingsPanel 执行时配置可能还没加载完）
            var custom = GetSavedCustomColor();
            PenSettingsPanel.SetCustomColorBrush(custom.HasValue ? new SolidColorBrush(custom.Value) : null);

            // 按笔种类套用笔宽 + 激光标记
            ApplyPenTypeToDrawingAttributes();

            // 颜色独立记忆：重启停在荧光笔时恢复其颜色记忆（默认黄）。
            // loadPenCanvas 默认把笔设为红色（画笔语义），不同步的话荧光笔会带着红色启动
            if (currentPenType == PenType.Highlighter)
            {
                int remembered = highlighterInkColor;
                if (remembered == 9 && !GetSavedCustomColor().HasValue) remembered = 1;
                inkColor = remembered;
                ApplyInkColorOnly(remembered);
            }
        }

        #endregion

        #region 面板弹出 / 定位 / 点击外部收起

        /// <summary>切换笔设置面板显隐（笔图标单击入口，PenIcon_MouseUp 调用）</summary>
        private void TogglePenSettingsPanel()
        {
            if (PenSettingsPanel.Visibility == Visibility.Visible)
            {
                PenSettingsPanel.Visibility = Visibility.Collapsed;
                return;
            }

            RefreshPenSettingsPanelState();
            PenSettingsPanel.Visibility = Visibility.Visible;
            PositionPenSettingsPanel();
        }

        /// <summary>关闭笔设置面板（HideSubPanels / 截图隐藏等场景调用）</summary>
        private void ClosePenSettingsPanel()
        {
            PenSettingsPanel.Visibility = Visibility.Collapsed;
        }

        /// <summary>弹出前刷新面板高亮状态（粗细/颜色/笔锋可能被设置页改过）</summary>
        private void RefreshPenSettingsPanelState()
        {
            PenSettingsPanel.UpdatePenTypeHighlight((int)currentPenType);
            // 粗细区换肤：荧光笔显横条（高亮带宽档），画笔/激光笔显圆点——
            // 两套档位独立记忆，换肤让"当前是哪套档位"一眼可辨
            PenSettingsPanel.SetThicknessStyle(currentPenType == PenType.Highlighter);
            PenSettingsPanel.UpdateThicknessHighlight(GetThicknessHighlightIndex());
            PenSettingsPanel.UpdateColorHighlight(inkColor);

            // 笔锋只对画笔有效（荧光笔矩形笔头/激光笔临时笔迹都不做压感模拟）：
            // 非画笔时灰显禁用笔锋区（不适用功能灰显规范），并清掉选中高亮
            bool taperRelevant = currentPenType == PenType.Pen;
            PenSettingsPanel.SetTaperSectionEnabled(taperRelevant);
            if (taperRelevant)
                PenSettingsPanel.UpdateTaperHighlight(Settings.Canvas.InkStyle);

            // 开关区：从母开关取当前值回显（母开关才是唯一状态源，面板不自己记账）
            PenSettingsPanel.UpdateInkToShapeState(ToggleSwitchEnableInkToShape.IsOn);
            PenSettingsPanel.UpdateCursorState(ToggleSwitchShowCursor.IsOn);
        }

        /// <summary>当前笔种类的粗细档位索引（0~4）：当前宽度与哪档预设最近就点亮哪档</summary>
        private int GetThicknessHighlightIndex()
        {
            double[] presets = currentPenType == PenType.Highlighter ? HighlighterWidthPresets : PenWidthPresets;
            double width = currentPenType == PenType.Highlighter ? _highlighterWidth : _penWidth;
            int best = 0;
            double bestDiff = double.MaxValue;
            for (int i = 0; i < presets.Length; i++)
            {
                double d = Math.Abs(presets[i] - width);
                if (d < bestDiff) { bestDiff = d; best = i; }
            }
            return best;
        }

        /// <summary>
        /// 面板定位：以笔图标为锚，图标上方空间够就弹在上方（顶不够时弹下方），
        /// 水平居中对齐笔图标并夹在屏幕内。
        /// 面板放在 Main_Grid 顶层（不在 Viewbox 内），不受浮动条缩放影响；
        /// 位置用变换实时换算（浮动条可拖动、可缩放，TransformToAncestor 全部涵盖）。
        /// </summary>
        private void PositionPenSettingsPanel()
        {
            // 先量出面板实际尺寸（刚设为 Visible，强制布局立即生效）
            Main_Grid.UpdateLayout();
            double panelW = PenSettingsPanel.ActualWidth;
            double panelH = PenSettingsPanel.ActualHeight;
            if (panelW < 1 || panelH < 1) return; // 布局未就绪，放弃定位（保持上次位置）

            // 笔图标在 Main_Grid 坐标系中的位置与渲染尺寸（含 Viewbox 缩放）
            var transform = PenIconHost.TransformToAncestor(Main_Grid);
            Point iconTopLeft = transform.Transform(new Point(0, 0));
            Point iconBottomRight = transform.Transform(new Point(PenIconHost.ActualWidth, PenIconHost.ActualHeight));
            double iconCenterX = (iconTopLeft.X + iconBottomRight.X) / 2;
            double iconBottomY = iconBottomRight.Y;
            double iconTopY = iconTopLeft.Y;

            const double Gap = 8; // 面板与浮动条间距
            double gridW = Main_Grid.ActualWidth;
            double gridH = Main_Grid.ActualHeight;

            // 垂直：图标上方放得下就放上方（贴近手、不挡下方板书），否则放图标下方
            double y = iconTopY - panelH - Gap;
            if (y < Gap) y = iconBottomY + Gap;
            if (y + panelH > gridH - Gap) y = Math.Max(Gap, gridH - panelH - Gap);

            // 水平：中心对齐笔图标，夹在屏幕内
            double x = iconCenterX - panelW / 2;
            if (x < Gap) x = Gap;
            if (x + panelW > gridW - Gap) x = Math.Max(Gap, gridW - panelW - Gap);

            PenSettingsPanel.Margin = new Thickness(x, y, 0, 0);
        }

        // 「点击面板外收起面板」的统一判定已移到 MW_PopupLayers.cs（Window_PreviewMouseDown，
        // XAML 仍绑定同名方法，故无需改 MainWindow.xaml）。
        // 那里有一张登记表（面板 + 开关图标 + 关闭动作），新增弹层只需加一行；
        // 请不要再在本文件恢复手写 if 判断——2026-09-11 前正是这种写法导致各面板行为不一致。

        #endregion

        #region 笔种类切换实现

        /// <summary>切换笔种类（面板事件入口）</summary>
        private void SetPenType(PenType type)
        {
            // 切笔种类时若在其他工具（橡皮/选择/图形），先切回画笔模式
            // （此调用发生在 currentPenType 赋值之前 → ColorSwitchCheck 里保存的是旧种类的颜色记忆）
            if (inkCanvas.EditingMode != InkCanvasEditingMode.Ink || drawingShapeMode != 0)
                ColorSwitchCheck();

            // 切笔种类（画笔/荧光笔/激光笔）= 书写新意图：旧选中清场（选中是临时上下文原则）。
            // 放在 ColorSwitchCheck 之后：那时已切回 Ink 模式，CancelActiveSelection 的
            // 模式保存/恢复逻辑保存的就是 Ink，不会有模式错乱
            CancelActiveSelection();

            currentPenType = type;

            // 持久化（记住上次用的笔，重启恢复）
            Settings.Canvas.PenType = (int)type;
            SaveSettingsToFile();

            ApplyPenTypeToDrawingAttributes();

            // 颜色独立记忆：切到哪个种类，就恢复哪个种类上次用的颜色
            // （荧光笔默认黄：半透明叠加，亮色才是荧光笔的正确语义）
            int remembered = type == PenType.Highlighter ? highlighterInkColor : penInkColor;
            if (remembered == 9 && !GetSavedCustomColor().HasValue)
                remembered = 1; // 自定义色已失效（清过设置）→ 回退红，不留无效索引
            inkColor = remembered;
            ApplyInkColorOnly(remembered);

            // 刷新全部颜色相关 UI（右面板高亮动画/笔图标颜色/快捷色条选中环），
            // 同时幂等地把恢复值写回记忆（见 ColorSwitchCheck 尾部）
            ColorSwitchCheck();

            RefreshPenSettingsPanelState();
        }

        /// <summary>
        /// 按索引直接把颜色写到 DefaultDrawingAttributes（纯设色，不触发换色流程）。
        /// 用途：笔种类切换时恢复记忆色——不能走 BtnColorX_Click 那套完整流程
        /// （会重复提交历史/改选中笔迹），只需要"上色"这一个动作。
        /// 索引语义与面板色格 Tag 一致：0黑 1红 2绿 3蓝 4黄 5白 6-8扩展色 9自定义。
        /// </summary>
        private void ApplyInkColorOnly(int index)
        {
            switch (index)
            {
                case 0: inkCanvas.DefaultDrawingAttributes.Color = Colors.Black; break;
                case 1: inkCanvas.DefaultDrawingAttributes.Color = ((SolidColorBrush)BtnColorRed.Background).Color; break;
                case 2: inkCanvas.DefaultDrawingAttributes.Color = ((SolidColorBrush)BtnColorGreen.Background).Color; break;
                case 3: inkCanvas.DefaultDrawingAttributes.Color = ((SolidColorBrush)BtnColorBlue.Background).Color; break;
                case 4: inkCanvas.DefaultDrawingAttributes.Color = ((SolidColorBrush)BtnColorYellow.Background).Color; break;
                case 5: inkCanvas.DefaultDrawingAttributes.Color = StringToColor("#FFFEFEFE"); break;
                case 6:
                case 7:
                case 8: inkCanvas.DefaultDrawingAttributes.Color = ExtraColors[index - 6]; break;
                case 9:
                    var custom = GetSavedCustomColor();
                    inkCanvas.DefaultDrawingAttributes.Color = custom ?? Colors.Red; // 失效回退红（调用方已保证 index=9 时必有色，此处纯兜底）
                    break;
            }
        }

        /// <summary>
        /// 图形模式笔迹属性规范化：进入图形绘制（UpdateShapeIconHighlight 搭车调用）前，
        /// 剥离笔种类专属属性（激光标记/荧光笔标志/荧光笔宽笔迹）。
        ///
        /// 为什么必须做：图形笔迹用 DefaultDrawingAttributes.Clone() 创建——
        /// 若激光标记残留，图形会"画完即渐隐消失"且不进撤销栈（被当成临时激光笔迹）；
        /// 若荧光笔标志残留，图形会变成 50% 半透明的宽笔迹。
        /// 切回笔/橡皮/选择（drawingShapeMode == 0）时由 ApplyPenTypeToDrawingAttributes 恢复。
        /// </summary>
        private void NormalizeAttributesForShapeMode()
        {
            // 注意：RemovePropertyData 对不存在的属性会抛 ArgumentException（WPF 文档明确），
            // 必须先 ContainsPropertyData 探测再删——否则从画笔/荧光笔直接切图形工具就会炸
            if (drawingAttributes.ContainsPropertyData(LaserStrokeGuid))
                drawingAttributes.RemovePropertyData(LaserStrokeGuid);
            drawingAttributes.IsHighlighter = false;
            // 图形笔宽用画笔宽度（图形是"墨迹类"笔迹，与荧光宽笔带无关）
            drawingAttributes.Width = _penWidth;
            drawingAttributes.Height = _penWidth;
        }

        /// <summary>把当前笔种类套用到 DefaultDrawingAttributes（笔宽/荧光标记/激光标记）</summary>
        private void ApplyPenTypeToDrawingAttributes()
        {
            // 各笔种类用各自记忆的笔宽（切换互不影响）
            double w = currentPenType == PenType.Highlighter ? _highlighterWidth : _penWidth;
            drawingAttributes.Width = w;
            drawingAttributes.Height = w;

            // 荧光笔：WPF 原生支持——矩形笔头 + 自动 50% 半透明，正好是"高亮笔"效果
            drawingAttributes.IsHighlighter = currentPenType == PenType.Highlighter;

            // 激光笔：打标记（收笔时据此启动淡出、跳过撤销栈）。
            // 修复切换 bug：RemovePropertyData 对不存在的属性抛 ArgumentException（WPF 行为），
            // 画笔→荧光笔切换时激光标记本来就不存在 → 异常直接打断切换流程（面板高亮不刷新的根因）
            if (currentPenType == PenType.Laser)
                drawingAttributes.AddPropertyData(LaserStrokeGuid, true);
            else if (drawingAttributes.ContainsPropertyData(LaserStrokeGuid))
                drawingAttributes.RemovePropertyData(LaserStrokeGuid);

            UpdatePenIconHighlight();
        }

        /// <summary>
        /// 面板粗细档点击（index 0~4）：写入当前笔种类的宽度记忆并立即生效。
        /// 荧光笔与画笔各自独立：在哪个种类下点的档位，就只改哪个种类的宽度。
        /// </summary>
        private void SetPenThicknessByIndex(int index)
        {
            if (index < 0 || index > 4) return;

            if (currentPenType == PenType.Highlighter)
            {
                _highlighterWidth = HighlighterWidthPresets[index];
                Settings.Canvas.HighlighterWidth = _highlighterWidth;
                drawingAttributes.Width = _highlighterWidth;
                drawingAttributes.Height = _highlighterWidth;
                SaveSettingsToFile();
            }
            else
            {
                _penWidth = PenWidthPresets[index];
                Settings.Canvas.InkWidth = _penWidth;
                drawingAttributes.Width = _penWidth;
                drawingAttributes.Height = _penWidth;
                SaveSettingsToFile();
                // 同步设置页滑条（滑条范围 1~20，荧光笔档位会越界，只在画笔/激光笔下同步）
                InkWidthSlider.Value = _penWidth * 2;
            }

            // 面板保持打开，刷新粗细档高亮
            if (PenSettingsPanel.Visibility == Visibility.Visible)
                RefreshPenSettingsPanelState();
        }

        /// <summary>
        /// 设置页滑条入口：只改画笔宽度（滑条在设置页里，语义就是"画笔粗细"）。
        /// 荧光笔模式下拖滑条不改变当前荧光带宽度——切回画笔时才生效。
        /// </summary>
        private void ApplyBasePenWidth(double width)
        {
            _penWidth = width;
            Settings.Canvas.InkWidth = width;

            if (currentPenType != PenType.Highlighter)
            {
                drawingAttributes.Width = width;
                drawingAttributes.Height = width;
            }

            // 面板打开时刷新粗细档位高亮（连续调整场景）
            if (PenSettingsPanel.Visibility == Visibility.Visible)
                PenSettingsPanel.UpdateThicknessHighlight(GetThicknessHighlightIndex());
        }

        /// <summary>面板颜色点击：索引 → 复用右面板换色逻辑（含切回画笔模式）</summary>
        private void SelectPenColorByIndex(int index)
        {
            switch (index)
            {
                case 0: BtnColorBlack_Click(BtnColorBlack, null); break;
                case 1: BtnColorRed_Click(BtnColorRed, null); break;
                case 2: BtnColorGreen_Click(BtnColorGreen, null); break;
                case 3: BtnColorBlue_Click(BtnColorBlue, null); break;
                case 4: BtnColorYellow_Click(BtnColorYellow, null); break;
                // 白色：内联右面板白色逻辑但不调 HideSubPanels——
                // （旧路径调 BorderPenColorWhite_MouseUp 会把整个笔设置面板关掉，与其他颜色行为不一致）
                case 5:
                    inkCanvas.DefaultDrawingAttributes.Color = StringToColor("#FFFEFEFE");
                    inkColor = 5;
                    forceEraser = false;
                    ColorSwitchCheck();
                    break;
                // 扩展三色（橙/紫/青）：直接设色，右面板无对应色块
                case 6: SelectExtraColor(0); break;
                case 7: SelectExtraColor(1); break;
                case 8: SelectExtraColor(2); break;
                // 自定义色：已设置→单击直接应用（像预设色一步到位）；未设置→打开系统取色板
                // （想重新选色用长按 0.6s 或右键——触屏没有右键，长按是主入口）
                case 9: SelectCustomColorSmart(); break;
            }
            // 注意：这里不走 HideSubPanels——面板保持打开，方便连续调整
        }

        /// <summary>
        /// 选择扩展色（橙/品红/青，ExtraColors 索引）。
        /// 右面板只有 5 个自定义色块 + 白，这 3 色无对应按钮——直接改 DefaultDrawingAttributes。
        /// </summary>
        private void SelectExtraColor(int extraIndex)
        {
            inkCanvas.DefaultDrawingAttributes.Color = ExtraColors[extraIndex];
            inkColor = 6 + extraIndex;
            forceEraser = false;
            ColorSwitchCheck();
        }

        /// <summary>
        /// 自定义色格单击（智能分流，避免"每次都要重新选色"）：
        /// - 已设置过颜色 → 直接应用该色（和预设色点完全同款的一步到位体验）
        /// - 从未设置 → 打开系统取色板引导首次选色
        /// 重新选色走长按 0.6s（触屏主入口）或右键（CustomColorReselectRequested → SelectCustomColor）。
        /// </summary>
        private void SelectCustomColorSmart()
        {
            var saved = GetSavedCustomColor();
            if (saved.HasValue)
            {
                inkCanvas.DefaultDrawingAttributes.Color = saved.Value;
                inkColor = 9;
                forceEraser = false;
                ColorSwitchCheck();
            }
            else
            {
                SelectCustomColor(); // 未设置过：打开取色板
            }
        }

        /// <summary>打开系统取色板选自定义色（WinForms ColorDialog，项目已启用 WinForms 互操作）</summary>
        private void SelectCustomColor()
        {
            // FullOpen：直接展开完整调色板（含 16 个自定义色槽），不用先点"规定自定义颜色"
            var dlg = new System.Windows.Forms.ColorDialog { FullOpen = true };
            if (dlg.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                var c = dlg.Color; // System.Drawing.Color → WPF Color
                ApplyCustomColor(Color.FromArgb(c.A, c.R, c.G, c.B));
            }
        }

        /// <summary>应用自定义色：写设置 + 面板色格更新 + 立即换色</summary>
        private void ApplyCustomColor(Color c)
        {
            Settings.Canvas.CustomColor = $"#{c.A:X2}{c.R:X2}{c.G:X2}{c.B:X2}";
            SaveSettingsToFile();
            PenSettingsPanel.SetCustomColorBrush(new SolidColorBrush(c));
            inkCanvas.DefaultDrawingAttributes.Color = c;
            inkColor = 9;
            forceEraser = false;
            ColorSwitchCheck();
        }

        /// <summary>读取已保存的自定义色（未设置/格式错返回 null，容错手改的 json）</summary>
        private Color? GetSavedCustomColor()
        {
            var s = Settings.Canvas.CustomColor;
            if (string.IsNullOrWhiteSpace(s)) return null;
            try { return StringToColor(s); }
            catch { return null; }
        }

        /// <summary>面板笔锋点击：复用设置页画笔样式（0=尾部 1=速度 2=无）</summary>
        private void SetPenTaperStyle(int style)
        {
            // 赋 SelectedIndex 触发 ComboBoxPenStyle_SelectionChanged 完成保存；
            // 若索引相同事件不触发（选了当前项），这里直接兜底写入设置
            ComboBoxPenStyle.SelectedIndex = style;
            Settings.Canvas.InkStyle = style;
            SaveSettingsToFile();

            // bug 修复：原来点完不刷新面板高亮，看起来"点击没反应"——
            // 实际设置已变（下笔生效），只是面板显示纹丝不动
            PenSettingsPanel.UpdateTaperHighlight(style);
        }

        /// <summary>
        /// 面板开关区「墨迹识别」：取反并落盘。
        /// 借道设置页母开关 ToggleSwitchEnableInkToShape：给它 IsOn 赋值会触发其 Toggled 落盘；
        /// 这里再显式写一次设置 + 保存，防止将来 Toggled 实现变动导致静默丢失。
        /// </summary>
        private void ToggleInkToShape()
        {
            bool next = !ToggleSwitchEnableInkToShape.IsOn;
            ToggleSwitchEnableInkToShape.IsOn = next;
            Settings.InkToShape.IsInkToShapeEnabled = next;
            SaveSettingsToFile();
            PenSettingsPanel.UpdateInkToShapeState(next);
        }

        /// <summary>
        /// 面板开关区「显示光标」：取反并落盘，同 ToggleInkToShape。
        /// 必须调 inkCanvas_EditingModeChanged——光标显隐是在切换编辑模式时套用的（与设置页 handler 一致），
        /// 只改设置不改画布，会出现"设置里是开、屏幕上没光标"。
        /// </summary>
        private void ToggleShowCursor()
        {
            bool next = !ToggleSwitchShowCursor.IsOn;
            ToggleSwitchShowCursor.IsOn = next;
            Settings.Canvas.IsShowCursor = next;
            inkCanvas_EditingModeChanged(inkCanvas, null);
            SaveSettingsToFile();
            PenSettingsPanel.UpdateCursorState(next);
        }

        #endregion

        #region 激光笔淡出

        /// <summary>
        /// 激光笔淡出：笔迹停留约 0.6s 后约 0.3s 内渐隐删除。
        /// 删除时置 CodeInput 绕过 TimeMachine（临时笔迹不进撤销栈，
        /// 否则撤销会把已消失的激光笔迹复活）。
        /// </summary>
        private void StartLaserFade(Stroke stroke)
        {
            Color baseColor = stroke.DrawingAttributes.Color;
            var timer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(40) };
            const int HoldTicks = 15;  // 停留 600ms
            const int FadeTicks = 8;   // 渐隐 320ms
            int tick = 0;

            timer.Tick += (s, e) =>
            {
                tick++;
                if (tick <= HoldTicks) return;

                // 笔迹已被清屏/撤销/手动擦除：停表走人
                if (!inkCanvas.Strokes.Contains(stroke))
                {
                    timer.Stop();
                    return;
                }

                double progress = (tick - HoldTicks) / (double)FadeTicks;
                if (progress >= 1)
                {
                    timer.Stop();
                    // CodeInput：不产生"删除"历史（激光笔迹从没进过历史，两边都不进）
                    _currentCommitType = CommitReason.CodeInput;
                    inkCanvas.Strokes.Remove(stroke);
                    _currentCommitType = CommitReason.UserInput;
                    // 清掉淡出期间累积的属性变更记录，防止污染下一次颜色切换的历史提交
                    DrawingAttributesHistory.Remove(stroke);
                }
                else
                {
                    // 渐隐：按进度降透明度（DrawingAttributes 是该笔迹私有副本，不影响其他笔迹）
                    byte alpha = (byte)(255 * (1 - progress));
                    stroke.DrawingAttributes.Color = Color.FromArgb(alpha, baseColor.R, baseColor.G, baseColor.B);
                }
            };
            timer.Start();
        }

        #endregion
    }
}
