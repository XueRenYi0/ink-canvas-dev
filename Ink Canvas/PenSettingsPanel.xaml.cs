using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;

namespace Ink_Canvas
{
    /// <summary>
    /// 笔设置面板：笔种类 / 粗细 / 颜色 / 笔锋。
    /// 纯 UI 组件——所有实际动作通过事件抛给 MainWindow 订阅处理，
    /// 自身不碰 inkCanvas / Settings，保持与主窗口解耦（一面板一控件规范）。
    /// </summary>
    public partial class PenSettingsPanel : UserControl
    {
        public PenSettingsPanel()
        {
            InitializeComponent();
        }

        // ===== 对外事件（MainWindow 订阅） =====

        /// <summary>笔种类被点击：0=画笔 1=荧光笔 2=激光笔</summary>
        public event Action<int> PenTypeSelected;

        /// <summary>粗细被点击：档位索引 0~4（具体宽度由 MainWindow 按笔种类换算）</summary>
        public event Action<int> ThicknessSelected;

        /// <summary>颜色被点击：0黑 1红 2绿 3蓝 4黄 5白</summary>
        public event Action<int> ColorSelected;

        /// <summary>笔锋被点击：0=尾部 1=速度 2=无（对应 Settings.Canvas.InkStyle）</summary>
        public event Action<int> TaperSelected;

        /// <summary>自定义颜色格右键被点击：请求重新打开取色板换色（左键是"使用颜色"语义）</summary>
        public event Action CustomColorReselectRequested;

        // ===== 状态更新入口（MainWindow 调用，用于同步面板高亮） =====

        /// <summary>注入 9 个颜色（前 5 色支持 Colors\*.ini 自定义，后 4 色为白/橙/紫/青固定色）</summary>
        public void SetColorBrushes(Brush[] brushes)
        {
            if (brushes == null) return;
            var dots = new[] { ColorDot0, ColorDot1, ColorDot2, ColorDot3, ColorDot4,
                               ColorDot5, ColorDot6, ColorDot7, ColorDot8 };
            for (int i = 0; i < Math.Min(9, brushes.Length); i++)
            {
                if (brushes[i] != null) dots[i].Fill = brushes[i];
            }
        }

        /// <summary>刷新笔种类选中高亮</summary>
        public void UpdatePenTypeHighlight(int type)
        {
            SetGridHighlight(PenTypePenHighlight, type == 0);
            SetGridHighlight(PenTypeHighlighterHighlight, type == 1);
            SetGridHighlight(PenTypeLaserHighlight, type == 2);
        }

        /// <summary>刷新粗细选中高亮（index 0~4 对应五档，由 MainWindow 计算好传入）</summary>
        public void UpdateThicknessHighlight(int index)
        {
            var highlights = new[] { Thickness1Highlight, Thickness2Highlight, Thickness3Highlight,
                                     Thickness4Highlight, Thickness5Highlight };
            for (int i = 0; i < highlights.Length; i++)
                SetGridHighlight(highlights[i], i == index);
        }

        /// <summary>
        /// 粗细区换肤：画笔/激光笔显圆点，荧光笔显横条（高亮带宽）。
        /// 两套档位各自独立记忆，换肤让用户一眼看出"当前是哪套档位"
        /// （PowerPoint/OneNote 荧光笔粗细预览也是扁矩形，同款设计）。
        /// tooltip 同步换成当前笔种类的实际宽度值。
        /// </summary>
        public void SetThicknessStyle(bool highlighter)
        {
            var dots = new[] { ThicknessDot1, ThicknessDot2, ThicknessDot3, ThicknessDot4, ThicknessDot5 };
            var bars = new[] { ThicknessBar1, ThicknessBar2, ThicknessBar3, ThicknessBar4, ThicknessBar5 };
            // 画笔/激光笔档位与荧光笔档位（与 MW_PenSettings 的预设数组一致）
            double[] penWidths = { 1.5, 2.5, 4, 6, 9 };
            double[] highlighterWidths = { 6, 10, 14, 20, 26 };
            var grids = new[] { BorderThickness1, BorderThickness2, BorderThickness3,
                                BorderThickness4, BorderThickness5 };
            for (int i = 0; i < 5; i++)
            {
                dots[i].Visibility = highlighter ? Visibility.Collapsed : Visibility.Visible;
                bars[i].Visibility = highlighter ? Visibility.Visible : Visibility.Collapsed;
                double w = highlighter ? highlighterWidths[i] : penWidths[i];
                grids[i].ToolTip = highlighter
                    ? $"荧光笔高亮带宽度：{w}{(i == 0 ? "（最细）" : i == 4 ? "（最粗）" : "")}"
                    : $"笔粗细：{w}{(i == 0 ? "（最细）" : i == 4 ? "（最粗）" : "")}";
            }
        }

        /// <summary>刷新颜色选中环（当前色中心显示细蓝环，支持 9 预设 + 自定义共 10 格）</summary>
        public void UpdateColorHighlight(int colorIndex)
        {
            var rings = new[] { ColorRing0, ColorRing1, ColorRing2, ColorRing3, ColorRing4,
                               ColorRing5, ColorRing6, ColorRing7, ColorRing8, ColorRing9 };
            for (int i = 0; i < rings.Length; i++)
                rings[i].Visibility = i == colorIndex ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>
        /// 设置自定义颜色格（第 10 格）：
        /// brush=null → 未设置状态（虚线圆 + "＋"号入口）；
        /// 有值 → 显示实色圆点 + 右下角"＋"角标（提示这格不是死色，再点可重新选色）。
        /// </summary>
        public void SetCustomColorBrush(Brush brush)
        {
            if (brush != null)
            {
                ColorDot9.Fill = brush;
                ColorDot9.StrokeDashArray = null; // 去掉虚线描边，变实心色点
                ColorDot9.Opacity = 1;
                CustomColorPlus.Visibility = Visibility.Collapsed;
                CustomColorBadge.Visibility = Visibility.Visible; // "可再选"角标
            }
            else
            {
                ColorDot9.Fill = Brushes.Transparent;
                ColorDot9.StrokeDashArray = new DoubleCollection { 2, 2 };
                ColorDot9.Opacity = 0.6;
                CustomColorPlus.Visibility = Visibility.Visible;
                CustomColorBadge.Visibility = Visibility.Collapsed; // 中央大＋号已够明显，角标不叠
            }
        }

        /// <summary>
        /// 笔锋区联动灰显：荧光笔（矩形笔头）和激光笔（临时笔迹）都不做压感模拟，
        /// 非画笔时禁用整区并压暗（不适用功能灰显规范）；禁用时清掉选中高亮。
        /// </summary>
        public void SetTaperSectionEnabled(bool enabled)
        {
            TaperSection.IsEnabled = enabled;   // 禁用时子元素收不到鼠标事件
            TaperSection.Opacity = enabled ? 1.0 : 0.35; // 压暗提示"不可用"
            if (!enabled)
            {
                SetGridHighlight(TaperEndHighlight, false);
                SetGridHighlight(TaperSpeedHighlight, false);
                SetGridHighlight(TaperNoneHighlight, false);
            }
        }

        /// <summary>刷新笔锋选中高亮</summary>
        public void UpdateTaperHighlight(int style)
        {
            SetGridHighlight(TaperEndHighlight, style == 0);
            SetGridHighlight(TaperSpeedHighlight, style == 1);
            SetGridHighlight(TaperNoneHighlight, style == 2);
        }

        // ===== 内部实现 =====

        /// <summary>选中态视觉：淡蓝底 + 蓝边，未选中透明</summary>
        private void SetGridHighlight(Border border, bool selected)
        {
            border.Background = selected
                ? (Brush)TryFindResource("PenPanelSoftHighlight")
                : Brushes.Transparent;
            border.BorderBrush = selected
                ? (Brush)TryFindResource("HighlightBrush")
                : Brushes.Transparent;
        }

        // 笔种类三个按钮共用一个 handler，用 Tag 区分（XAML 里没写 Tag，这里按控件名识别）
        private void PenTypeButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderPenTypePen) PenTypeSelected?.Invoke(0);
            else if (sender == BorderPenTypeHighlighter) PenTypeSelected?.Invoke(1);
            else if (sender == BorderPenTypeLaser) PenTypeSelected?.Invoke(2);
        }

        // 粗细五个按钮共用：只报档位索引，具体宽度由 MainWindow 按笔种类换算
        // （画笔：1.5/2.5/4/6/9，荧光笔：6/10/14/20/26——两套档位各自独立记忆）
        private void ThicknessButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderThickness1) ThicknessSelected?.Invoke(0);
            else if (sender == BorderThickness2) ThicknessSelected?.Invoke(1);
            else if (sender == BorderThickness3) ThicknessSelected?.Invoke(2);
            else if (sender == BorderThickness4) ThicknessSelected?.Invoke(3);
            else if (sender == BorderThickness5) ThicknessSelected?.Invoke(4);
        }

        // 色点用 Tag 区分（0~9：0黑 1红 2绿 3蓝 4黄 5白 6橙 7紫 8青 9自定义）
        private void ColorDot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return; // 右键走 CustomColorDot_MouseRightButtonUp
            _customColorPressTimer?.Stop(); // 抬起即取消长按计时（未到 0.6s = 普通单击）
            if (sender is Grid grid && int.TryParse(grid.Tag?.ToString(), out int index))
            {
                // 长按刚弹过取色板：这次抬起不再触发"应用颜色"（否则选完色关窗瞬间又把旧色应用回去）
                if (index == 9 && _customColorLongPressed)
                {
                    _customColorLongPressed = false;
                    return;
                }
                ColorSelected?.Invoke(index);
            }
        }

        // ===== 自定义色格长按（触屏没有右键，长按 = 重新打开取色板换色）=====
        private DispatcherTimer _customColorPressTimer; // 长按计时器
        private bool _customColorLongPressed;           // 本次按下是否已触发过长按

        /// <summary>按下自定义格：启动 0.6s 计时（按住不动 = 长按换色）</summary>
        private void CustomColorDot_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            _customColorLongPressed = false; // 新一次按下，清掉上一次的残留标志
            _customColorPressTimer?.Stop();
            _customColorPressTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(600) };
            _customColorPressTimer.Tick += (s, _) =>
            {
                _customColorPressTimer.Stop();
                _customColorLongPressed = true;
                CustomColorReselectRequested?.Invoke(); // 长按成立 → 打开取色板
            };
            _customColorPressTimer.Start();
        }

        /// <summary>手指/鼠标移出自定义格：取消长按计时（误触滑动不触发）</summary>
        private void CustomColorDot_MouseLeave(object sender, MouseEventArgs e)
        {
            _customColorPressTimer?.Stop();
        }

        /// <summary>
        /// 自定义颜色格右键（鼠标方案）：请求重新打开取色板换色。
        /// 触屏走长按（CustomColorDot_MouseDown 计时 0.6s），两者殊途同归。
        /// （左键语义是"使用颜色"——已设置时直接应用，未设置时才打开取色板，见 MW_PenSettings.SelectCustomColorSmart）
        /// </summary>
        private void CustomColorDot_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            _customColorPressTimer?.Stop(); // 右键抬起时取消可能还在跑的长按计时
            CustomColorReselectRequested?.Invoke();
        }

        // 笔锋三个按钮共用：尾部0 / 速度1 / 无2（对应 InkStyle）
        private void TaperButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderTaperEnd) TaperSelected?.Invoke(0);
            else if (sender == BorderTaperSpeed) TaperSelected?.Invoke(1);
            else if (sender == BorderTaperNone) TaperSelected?.Invoke(2);
        }
    }
}
