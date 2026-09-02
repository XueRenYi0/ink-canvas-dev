using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// 截图遮罩窗口（自绘框选截图的交互核心）。
    ///
    /// 【设计要点】
    /// 1. 状态完全自持：窗口开着 = 正在框选；窗口关了 = 流程结束。
    ///    "取消"是一个瞬时动作（关窗即取消），不存在等系统回调的悬空状态。
    /// 2. 本窗口只负责"收集选区矩形"，不负责抓屏——
    ///    抓屏必须在窗口关闭之后进行（否则会把遮罩自己截进去），
    ///    由调用方（MW_Capture.cs）在 ShowDialog 返回后执行。
    /// 3. 坐标体系：窗口内是 WPF 逻辑坐标；返回的选区也是逻辑坐标，
    ///    调用方用 DpiScaleX/Y（本窗口暴露）换算成物理像素再 CopyFromScreen。
    ///    高分屏（如 150% 缩放的教学一体机）不做这一步截图区域就会偏移。
    /// 4. 输入兼容鼠标/触摸/手写笔：笔事件优先（e.Handled 阻断合成鼠标事件），
    ///    鼠标事件带 _stylusActive 防重标志兜底。
    /// </summary>
    public partial class ScreenshotMaskWindow : Window
    {
        // ---------- 对外结果 ----------
        /// <summary>用户框选的选区（窗口逻辑坐标）；DialogResult=true 时有效</summary>
        public Rect? SelectedRect { get; private set; }

        /// <summary>选区确认事件：在窗口关闭前一刻触发。
        /// 调用方借此在"遮罩还盖着屏幕"时同步隐藏自己的 UI（悬浮条等），
        /// 这样遮罩关闭后屏幕上不会有软件 UI 的闪现帧。</summary>
        public event Action SelectionMade;

        // ---------- DPI 换算（供调用方把逻辑坐标转物理像素） ----------
        public double DpiScaleX { get; private set; } = 1.0;
        public double DpiScaleY { get; private set; } = 1.0;

        // ---------- 内部状态 ----------
        /// <summary>拖选起点（逻辑坐标）</summary>
        private Point _startPoint;

        /// <summary>是否正在拖选（按下未松开）</summary>
        private bool _isDragging = false;

        /// <summary>当前输入来自手写笔/触摸（true=鼠标事件是被合成的，需忽略防双触发）</summary>
        private bool _stylusActive = false;

        /// <summary>关窗防重入标志：true=已在关闭流程中。
        /// 【实测坑】设 DialogResult 后窗口开始关闭，关闭过程还会触发一次 Deactivated，
        /// 若不拦住会二次设 DialogResult 抛 InvalidOperationException（"只能在对话框显示之后设置"）。</summary>
        private bool _isClosing = false;

        /// <summary>选区最小尺寸（逻辑像素）：比这还小视为误点/取消</summary>
        private const double MinSelection = 10;

        public ScreenshotMaskWindow()
        {
            InitializeComponent();

            // 铺满主屏（教学场景按单屏设计；多屏扩展时只截主屏）
            Left = 0;
            Top = 0;
            Width = SystemParameters.PrimaryScreenWidth;
            Height = SystemParameters.PrimaryScreenHeight;

            // 初始暗罩 = 只有一个全屏矩形（整屏全暗）
            MaskPath.Data = new RectangleGeometry(new Rect(0, 0, Width, Height));

            Loaded += (s, e) =>
            {
                // 读取本窗口实际 DPI 缩放（必须在 Loaded 后视觉树才就绪）
                var dpi = VisualTreeHelper.GetDpi(this);
                DpiScaleX = dpi.PixelsPerDip;
                DpiScaleY = dpi.PixelsPerDip;
            };
        }

        // ==================== 开始 / 更新 / 结束拖选 ====================

        /// <summary>开始拖选：记录起点，显示选区框，隐藏顶部提示</summary>
        private void BeginSelection(Point p)
        {
            _startPoint = p;
            _isDragging = true;
            HintText.Visibility = Visibility.Collapsed; // 提示只出现到第一次落点
            SelectionBorder.Visibility = Visibility.Visible;
            UpdateSelectionVisual(p);
        }

        /// <summary>拖选中：以起点和当前点实时更新选区框、暗罩挖洞、尺寸角标</summary>
        private void UpdateSelectionVisual(Point p)
        {
            Rect rect = new Rect(_startPoint, p); // Rect 构造自动处理起点在右/下的情况（归一化）

            // 选区蓝框（全限定名：项目里有个 Ink_Canvas.Canvas 类会遮蔽 WPF 的 Canvas）
            System.Windows.Controls.Canvas.SetLeft(SelectionBorder, rect.X);
            System.Windows.Controls.Canvas.SetTop(SelectionBorder, rect.Y);
            SelectionBorder.Width = rect.Width;
            SelectionBorder.Height = rect.Height;

            // 暗罩挖洞：全屏矩形 + 选区矩形 的 EvenOdd 组合 → 选区处透亮
            var group = new GeometryGroup { FillRule = FillRule.EvenOdd };
            group.Children.Add(new RectangleGeometry(new Rect(0, 0, ActualWidth, ActualHeight)));
            group.Children.Add(new RectangleGeometry(rect));
            MaskPath.Data = group;

            // 尺寸角标：贴在选区右下角（显示的是物理像素，用户预期的"实际截出来多大"）
            SizeBadge.Visibility = Visibility.Visible;
            SizeText.Text = string.Format("{0:0} × {1:0}",
                Math.Round(rect.Width * DpiScaleX), Math.Round(rect.Height * DpiScaleY));
            // 角标放选区右下角内侧，防超出屏幕；选区太矮时放框下方
            double badgeX = rect.X + rect.Width - SizeBadge.ActualWidth;
            double badgeY = rect.Y + rect.Height + 4;
            if (badgeY + 24 > ActualHeight) badgeY = rect.Y + rect.Height - SizeBadge.ActualHeight - 4;
            System.Windows.Controls.Canvas.SetLeft(SizeBadge, Math.Max(0, badgeX));
            System.Windows.Controls.Canvas.SetTop(SizeBadge, Math.Max(0, badgeY));
        }

        /// <summary>结束拖选（松手）：选区足够大 → 确认；太小 → 视为取消</summary>
        private void EndSelection(Point p)
        {
            _isDragging = false;
            Rect rect = new Rect(_startPoint, p);

            // 单击没拖开（< 10px）= 误触或想取消 → 按取消处理
            if (rect.Width < MinSelection || rect.Height < MinSelection)
            {
                Cancel();
                return;
            }

            SelectedRect = rect;
            _isClosing = true; // 先立关窗标志：下面的 Close() 关闭中会触发 Deactivated → Cancel()，标志拦住它防止把结果改写为取消
            // 关窗前通知调用方"选区定了"——趁机把软件 UI 藏掉（此刻遮罩还盖着，用户看不到闪现）
            try { SelectionMade?.Invoke(); } catch { }
            DialogResult = true; // ShowDialog 返回 true = 有选区
            Close();
        }

        // ==================== 取消 ====================

        /// <summary>取消截图：直接关窗（DialogResult=false），无任何副作用</summary>
        private void Cancel()
        {
            if (_isClosing) return; // 已在关闭流程（如 Esc 后关闭中又触发 Deactivated），跳过防二次设 DialogResult
            _isClosing = true;
            DialogResult = false;
            Close();
        }

        // ==================== 输入事件：手写笔/触摸（优先） ====================

        private void Window_PreviewStylusDown(object sender, StylusDownEventArgs e)
        {
            e.Handled = true; // 阻断 WPF 把笔事件合成的鼠标事件（防双触发）
            _stylusActive = true;
            try { Stylus.Capture(this); } catch { } // 捕获笔：拖出窗口也能收到 Move/Up
            BeginSelection(e.GetPosition(this));
        }

        private void Window_PreviewStylusMove(object sender, StylusEventArgs e)
        {
            if (!_stylusActive) return;
            e.Handled = true;
            if (_isDragging) UpdateSelectionVisual(e.GetPosition(this));
        }

        private void Window_PreviewStylusUp(object sender, StylusEventArgs e)
        {
            if (!_stylusActive) return;
            e.Handled = true;
            _stylusActive = false;
            try { Stylus.Capture(null); } catch { }
            EndSelection(e.GetPosition(this));
        }

        // ==================== 输入事件：鼠标（手写笔合成事件被 Handled 挡住，这里只接真实鼠标） ====================

        private void Window_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (_stylusActive) return; // 笔输入进行中，忽略合成鼠标事件
            try { Mouse.Capture(this); } catch { }
            BeginSelection(e.GetPosition(this));
        }

        private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (_stylusActive) return;
            if (_isDragging && e.LeftButton == MouseButtonState.Pressed)
                UpdateSelectionVisual(e.GetPosition(this));
        }

        private void Window_PreviewMouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (_stylusActive) return;
            try { Mouse.Capture(null); } catch { }
            EndSelection(e.GetPosition(this));
        }

        // ==================== 取消路径：Esc / 右键 / 失焦 ====================

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                Cancel();
            }
        }

        private void Window_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            e.Handled = true;
            Cancel();
        }

        /// <summary>窗口失焦（Alt+Tab 切走等）= 用户放弃截图，自动取消收场</summary>
        private void Window_Deactivated(object sender, EventArgs e)
        {
            if (IsLoaded) Cancel();
        }
    }
}
