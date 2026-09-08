using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;

namespace Ink_Canvas
{
    /// <summary>
    /// 橡皮设置面板：擦除方式（面积擦/笔画擦）/ 大小（五档）/ 滑动清屏。
    /// 纯 UI 组件——所有实际动作通过事件抛给 MainWindow 订阅处理，
    /// 自身不碰 inkCanvas / Settings，保持与主窗口解耦（一面板一控件规范）。
    /// </summary>
    public partial class EraserSettingsPanel : UserControl
    {
        public EraserSettingsPanel()
        {
            InitializeComponent();
        }

        // ===== 对外事件（MainWindow 订阅） =====

        /// <summary>擦除方式被点击：0=面积擦 1=笔画擦</summary>
        public event Action<int> EraserModeSelected;

        /// <summary>大小被点击：档位索引 0~4（30/45/60/80/100）</summary>
        public event Action<int> EraserSizeSelected;

        /// <summary>滑动清屏触发（滑块拖到最右）</summary>
        public event Action ClearRequested;

        // ===== 状态更新入口（MainWindow 调用，用于同步面板高亮） =====

        /// <summary>刷新擦除方式选中高亮（0=面积擦 1=笔画擦）</summary>
        public void UpdateModeHighlight(int mode)
        {
            SetGridHighlight(ModePointHighlight, mode == 0);
            SetGridHighlight(ModeStrokeHighlight, mode == 1);
        }

        /// <summary>刷新大小档位选中高亮（index 0~4，-1 表示无选中）</summary>
        public void UpdateSizeHighlight(int index)
        {
            var highlights = new[] { Size1Highlight, Size2Highlight, Size3Highlight, Size4Highlight, Size5Highlight };
            for (int i = 0; i < highlights.Length; i++)
                SetGridHighlight(highlights[i], i == index);
        }

        /// <summary>
        /// 大小区可用性：笔画擦无大小概念（碰到墨迹即删整笔），整区压暗禁用；
        /// 切回面积擦时恢复。与笔面板"笔锋区灰显"同款规范（不适用功能灰显）。
        /// </summary>
        public void SetSizeSectionEnabled(bool enabled)
        {
            var grids = new[] { BorderSize1, BorderSize2, BorderSize3, BorderSize4, BorderSize5 };
            foreach (var grid in grids)
            {
                grid.Opacity = enabled ? 1 : 0.35;
                grid.IsEnabled = enabled;
                // 恢复时清掉灰显期间残留的高亮边框
                if (!enabled)
                {
                    var highlight = FindName($"Size{Array.IndexOf(grids, grid) + 1}Highlight") as Border;
                    if (highlight != null)
                    {
                        highlight.Background = Brushes.Transparent;
                        highlight.BorderBrush = Brushes.Transparent;
                    }
                }
            }
        }

        /// <summary>重置滑动清屏滑块到起点（清屏触发后/面板重新弹出时调用）</summary>
        public void ResetSlideThumb()
        {
            // 先移除可能残留的回弹动画再归零——WPF 动画默认 HoldEnd，动画结束后
            // 仍"占着"X 属性，之后直接赋值会被动画压制，导致滑块再也拖不动
            SlideThumbTranslate.BeginAnimation(TranslateTransform.XProperty, null);
            SlideThumbTranslate.X = 0;
        }

        // ===== 内部：按钮点击分发 =====

        private void ModeButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderModePoint) EraserModeSelected?.Invoke(0);
            else if (sender == BorderModeStroke) EraserModeSelected?.Invoke(1);
        }

        private void SizeButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderSize1) EraserSizeSelected?.Invoke(0);
            else if (sender == BorderSize2) EraserSizeSelected?.Invoke(1);
            else if (sender == BorderSize3) EraserSizeSelected?.Invoke(2);
            else if (sender == BorderSize4) EraserSizeSelected?.Invoke(3);
            else if (sender == BorderSize5) EraserSizeSelected?.Invoke(4);
        }

        // ===== 内部：滑动清屏拖动 =====

        /// <summary>是否正在拖动滑块（拖动中屏蔽 MouseMove 的杂散触发）</summary>
        private bool _isSlideDragging = false;

        /// <summary>触发清屏的滑动阈值（滑块行程的百分比，滑过即执行）</summary>
        private const double SlideTriggerRatio = 0.85;

        private void SlideThumb_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isSlideDragging = true;
            // 捕获鼠标：拖出滑块/轨道范围也能持续收到 Move/Up，不丢事件
            SlideClearThumb.CaptureMouse();
            e.Handled = true;
        }

        private void SlideThumb_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isSlideDragging) return;

            // 计算滑块在轨道内的可行进范围（轨道宽 - 滑块宽 - 两侧留白）
            SlideClearTrack.UpdateLayout();
            double trackWidth = SlideClearTrack.ActualWidth;
            double maxTranslate = Math.Max(0, trackWidth - SlideClearThumb.Width - 4); // 4 = 左右各 2px 留白

            // 鼠标相对轨道的 X 坐标 → 滑块中心位置 → TranslateTransform 偏移
            Point pos = e.GetPosition(SlideClearTrack);
            double thumbCenter = pos.X;
            double translate = thumbCenter - SlideClearThumb.Width / 2;

            // 夹在 [0, maxTranslate] 范围内
            if (translate < 0) translate = 0;
            if (translate > maxTranslate) translate = maxTranslate;
            SlideThumbTranslate.X = translate;
        }

        private void SlideThumb_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isSlideDragging) return;
            _isSlideDragging = false;
            SlideClearThumb.ReleaseMouseCapture();

            double trackWidth = SlideClearTrack.ActualWidth;
            double maxTranslate = Math.Max(1, trackWidth - SlideClearThumb.Width - 4);

            if (SlideThumbTranslate.X >= maxTranslate * SlideTriggerRatio)
            {
                // 滑到底：先回位再触发（MainWindow 会顺带关面板）
                ResetSlideThumb();
                ClearRequested?.Invoke();
            }
            else
            {
                // 未滑到底：回弹动画（120ms，快速利落）
                double from = SlideThumbTranslate.X;
                var anim = new DoubleAnimation(from, 0, TimeSpan.FromMilliseconds(120))
                {
                    EasingFunction = new QuadraticEase { EasingMode = EasingMode.EaseOut }
                };
                // 动画完成后必须移除动画释放属性（否则 HoldEnd 占着 X，下次拖不动——
                // 即用户反馈的"回弹以后后面划不动"bug 的根因）
                anim.Completed += (s, _) =>
                {
                    SlideThumbTranslate.BeginAnimation(TranslateTransform.XProperty, null);
                    SlideThumbTranslate.X = 0;
                };
                SlideThumbTranslate.BeginAnimation(TranslateTransform.XProperty, anim);
            }
        }

        // ===== 内部：通用高亮 =====

        /// <summary>设置档位高亮框：选中=淡蓝底+蓝边，未选中=透明（与笔面板同款）</summary>
        private void SetGridHighlight(Border highlight, bool selected)
        {
            if (selected)
            {
                highlight.Background = FindResource("EraserPanelSoftHighlight") as Brush;
                highlight.BorderBrush = FindResource("HighlightBrush") as Brush;
            }
            else
            {
                highlight.Background = Brushes.Transparent;
                highlight.BorderBrush = Brushes.Transparent;
            }
        }
    }
}
