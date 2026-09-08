using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：选择方式（矩形框选 / 自由选择）+ 选择方式面板。
    ///
    /// 交互模型（与笔面板同款弹出模式）：
    /// - 点"选择墨迹"图标 → 进入选择模式 + 弹出选择方式面板
    /// - 面板内选一项 → 立即生效并关闭面板（选完即关，用户确认的交互）
    /// - 点面板外任意处 → 面板消失（选择模式保持，Window_PreviewMouseDown 判定）
    ///
    /// 矩形框选实现原理：
    /// - WPF InkCanvas 的 Select 模式原生只支持套索（Lasso）
    /// - 本文件在 Preview 隧道阶段拦截空白处的按下事件（e.Handled=true），
    ///   阻止 InkCanvas 启动套索，改为自绘蓝色虚线矩形预览
    /// - 松手后用 inkCanvas.Select(选中集) 完成选中——之后走现有
    ///   inkCanvas_SelectionChanged 流程（选区控制条/快照等自动出现）
    /// - 点在墨迹上不拦截：保留原生"点选/拖动墨迹"体验；
    ///   点在已有选区附近不拦截：选区覆盖层（BorderSelectionHitArea）自己接管缩放/旋转
    /// </summary>
    public partial class MainWindow
    {
        #region 选择方式状态

        /// <summary>选择方式：0=矩形框选（默认） 1=自由选择（套索）</summary>
        private enum SelMode { Rect = 0, Lasso = 1 }

        /// <summary>当前选择方式（持久化于 Settings.Canvas.SelectionMode）</summary>
        private SelMode currentSelMode = SelMode.Rect;

        // ===== 矩形拖选运行时状态 =====

        /// <summary>是否正在拖拽矩形框选</summary>
        private bool _rectSelectDragging = false;

        /// <summary>拖选起点（inkCanvas 坐标系）</summary>
        private Point _rectSelectStart;

        /// <summary>矩形预览 Adorner（空=未显示）</summary>
        private RectSelectAdorner _rectAdorner;

        #endregion

        #region 初始化与恢复（MW_DefinitionsLoading 调用）

        /// <summary>初始化：订阅面板事件 + 挂矩形拖选拦截事件</summary>
        private void InitSelectionMode()
        {
            // 面板选择 → 切方式（选完即关）
            SelectionModePanel.SelectionModeSelected += m => SetSelectionMode(m);
            // 面板全选 → 立即全选当前页墨迹（选完即关）
            SelectionModePanel.SelectAllRequested += SelectAllStrokes;

            // 触摸笔尖（Stylus）与鼠标成对订阅（MW_CustomShapes 的成熟模式）：
            // StylusDown 设 Handled 后 promoted MouseDown 不再触发，天然去重
            inkCanvas.PreviewStylusDown += InkCanvas_RectSelectStylusDown;
            inkCanvas.PreviewMouseDown += InkCanvas_RectSelectMouseDown;
            // Move/Up 不设 Handled，两个来源可能都触发——处理器做成幂等（先查拖拽标志），双触发无害
            inkCanvas.PreviewStylusMove += InkCanvas_RectSelectStylusMove;
            inkCanvas.PreviewMouseMove += InkCanvas_RectSelectMouseMove;
            inkCanvas.PreviewStylusUp += InkCanvas_RectSelectStylusUp;
            inkCanvas.PreviewMouseUp += InkCanvas_RectSelectMouseUp;

            // 防御：拖拽中鼠标捕获被系统/其他代码抢走时收尾，
            // 避免预览框残留在画布上、拖拽状态卡死（此时视为取消，不选中任何内容）
            inkCanvas.LostMouseCapture += (s, e) => CancelRectSelectIfDragging();
        }

        /// <summary>配置加载后恢复选择方式（InitPenSettingsPanel 同批次调用）</summary>
        private void ApplyLoadedSelectionMode()
        {
            currentSelMode = (SelMode)Math.Min(1, Math.Max(0, Settings.Canvas.SelectionMode));
        }

        #endregion

        #region 选择方式切换与面板

        /// <summary>
        /// 切换选择方式（面板点击入口）：保存 + 关面板（选完即关）。
        /// 切换只影响"框选的手势"，选择模式本身不退出——可立刻用新方式框选。
        /// </summary>
        private void SetSelectionMode(int mode)
        {
            currentSelMode = (SelMode)Math.Min(1, Math.Max(0, mode));
            Settings.Canvas.SelectionMode = (int)currentSelMode;
            SaveSettingsToFile();

            CloseSelectionModePanel();
        }

        /// <summary>选择图标高亮：进入选择模式亮淡蓝底+蓝边（与笔图标同规范），离开熄灭</summary>
        private void UpdateSelectIconHighlight()
        {
            bool active = inkCanvas.EditingMode == InkCanvasEditingMode.Select;
            BorderSelectIconHighlight.Background = active ? PenToolHighlightSoftBrush : Brushes.Transparent;
            BorderSelectIconHighlight.BorderBrush = active ? PenToolHighlightBrush : Brushes.Transparent;
        }

        /// <summary>
        /// 全选当前页全部墨迹（面板"全选"项，原"双击选择图标"的旧功能迁移至此）。
        /// 复用旧 BtnSelect_Click 的全选实现：跳过零宽/零高的退化笔迹——
        /// 这类笔迹选中后无法拖动/缩放/删除（WPF 选区逻辑限制），混入会卡操作条。
        /// 全选后关闭面板（动作项，选完即关）。
        /// </summary>
        private void SelectAllStrokes()
        {
            CloseSelectionModePanel();
            if (inkCanvas.EditingMode != InkCanvasEditingMode.Select)
                inkCanvas.EditingMode = InkCanvasEditingMode.Select;

            StrokeCollection selectedStrokes = new StrokeCollection();
            foreach (Stroke stroke in inkCanvas.Strokes)
            {
                try
                {
                    if (stroke.GetBounds().Width > 0 && stroke.GetBounds().Height > 0)
                        selectedStrokes.Add(stroke);
                }
                catch { }
            }

            // 图片一并全选（与矩形框选同标准：只选当前视图可见的页面图片，
            // 其他页/隐藏层的不可见图片不参选）——走 Select 双参重载进原生选区
            var selectedImages = new System.Collections.Generic.List<UIElement>();
            try
            {
                foreach (var img in _imagePageKey.Keys)
                {
                    if (img.Visibility != Visibility.Visible) continue;
                    selectedImages.Add(img);
                }
            }
            catch { }

            // 空集 = 清空选中（与框选行为一致）
            inkCanvas.Select(selectedStrokes, selectedImages);
        }

        /// <summary>切换选择方式面板显隐（选择图标单击入口）</summary>
        private void ToggleSelectionModePanel()
        {
            if (SelectionModePanel.Visibility == Visibility.Visible)
            {
                SelectionModePanel.Visibility = Visibility.Collapsed;
                return;
            }

            SelectionModePanel.UpdateModeHighlight((int)currentSelMode);
            SelectionModePanel.Visibility = Visibility.Visible;
            PositionSelectionModePanel();
        }

        /// <summary>关闭选择方式面板（HideSubPanels / 工具条收起等场景调用）</summary>
        private void CloseSelectionModePanel()
        {
            SelectionModePanel.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// 面板定位：以选择图标为锚（逻辑与笔面板 PositionPenSettingsPanel 相同），
        /// 上方空间够弹上方，否则下方；水平居中对齐图标并夹在屏幕内。
        /// </summary>
        private void PositionSelectionModePanel()
        {
            Main_Grid.UpdateLayout();
            double panelW = SelectionModePanel.ActualWidth;
            double panelH = SelectionModePanel.ActualHeight;
            if (panelW < 1 || panelH < 1) return; // 布局未就绪，放弃定位

            var transform = GridSelectTool.TransformToAncestor(Main_Grid);
            Point iconTopLeft = transform.Transform(new Point(0, 0));
            Point iconBottomRight = transform.Transform(new Point(GridSelectTool.ActualWidth, GridSelectTool.ActualHeight));
            double iconCenterX = (iconTopLeft.X + iconBottomRight.X) / 2;
            double iconBottomY = iconBottomRight.Y;
            double iconTopY = iconTopLeft.Y;

            const double Gap = 8;
            double gridW = Main_Grid.ActualWidth;
            double gridH = Main_Grid.ActualHeight;

            double y = iconTopY - panelH - Gap;
            if (y < Gap) y = iconBottomY + Gap;
            if (y + panelH > gridH - Gap) y = Math.Max(Gap, gridH - panelH - Gap);

            double x = iconCenterX - panelW / 2;
            if (x < Gap) x = Gap;
            if (x + panelW > gridW - Gap) x = Math.Max(Gap, gridW - panelW - Gap);

            SelectionModePanel.Margin = new Thickness(x, y, 0, 0);
        }

        #endregion

        #region 矩形框选拦截（Preview 阶段）

        // ---- 按下：决定是否接管 ----

        private void InkCanvas_RectSelectStylusDown(object sender, StylusDownEventArgs e)
        {
            if (TryStartRectSelect(e.GetPosition(inkCanvas)))
            {
                // 拦下：InkCanvas 的套索不再启动，promoted MouseDown 也不再触发
                e.Handled = true;
                CaptureForRectSelect();
            }
        }

        private void InkCanvas_RectSelectMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            if (TryStartRectSelect(e.GetPosition(inkCanvas)))
            {
                e.Handled = true;
                CaptureForRectSelect();
            }
        }

        /// <summary>
        /// 尝试开始矩形拖选。接管条件（全部满足）：
        /// 1. 当前方式=矩形 2. 处于选择编辑模式 3. 落点在空白处（不在任何墨迹/图片上）。
        /// 点在墨迹上返回 false 不拦截——保留原生"点选墨迹/拖动已选墨迹"体验；
        /// 点在图片上同样放行——原生支持点选图片（选择模式下图片可命中）。
        /// </summary>
        private bool TryStartRectSelect(Point pos)
        {
            if (_rectSelectDragging) return false; // 已在拖拽中（防重入）
            if (currentSelMode != SelMode.Rect) return false;
            if (inkCanvas.EditingMode != InkCanvasEditingMode.Select) return false;

            // 落点命中测试——【坑】StrokeCollection.HitTest(Point, double) 的第二参数是
            // "笔迹宽度的百分比"（有效范围 0~1），误传 6 会让调用抛异常：旧代码 try/catch
            // 把异常吞掉后继续接管 → 点墨迹也变成拖矩形（单击=清空选中），点选完全失效。
            // 正确用法 1.0 = 1 倍笔宽容差（细线好点、粗线宽容）。
            try
            {
                if (inkCanvas.Strokes.HitTest(pos, 1.0).Count > 0) return false;
            }
            catch { }

            // 图片命中：点在图片上放行（原生点选图片；图片在选择模式下 IsHitTestVisible=true）
            foreach (var img in _imagePageKey.Keys)
            {
                if (img.Visibility != Visibility.Visible) continue;
                var b = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                b.Inflate(2, 2); // 2px 容差
                if (b.Contains(pos)) return false;
            }

            // 接管：记录起点 + 显示预览
            _rectSelectDragging = true;
            _rectSelectStart = pos;
            ShowRectAdorner(pos, pos);
            return true;
        }

        /// <summary>捕获鼠标+触摸笔：拖拽期间手指/光标移出画布也能持续收到 Move/Up</summary>
        private void CaptureForRectSelect()
        {
            try { inkCanvas.CaptureMouse(); } catch { }
            try { Stylus.Capture(inkCanvas); } catch { }
        }

        // ---- 移动：更新预览矩形 ----

        private void InkCanvas_RectSelectStylusMove(object sender, StylusEventArgs e)
        {
            if (!_rectSelectDragging) return;
            UpdateRectPreview(e.GetPosition(inkCanvas));
        }

        private void InkCanvas_RectSelectMouseMove(object sender, MouseEventArgs e)
        {
            if (!_rectSelectDragging) return;
            UpdateRectPreview(e.GetPosition(inkCanvas));
        }

        private void UpdateRectPreview(Point pos)
        {
            ShowRectAdorner(_rectSelectStart, pos);
        }

        // ---- 松手：完成选中 ----

        private void InkCanvas_RectSelectStylusUp(object sender, StylusEventArgs e)
        {
            if (!_rectSelectDragging) return;
            e.Handled = true;
            EndRectSelect(e.GetPosition(inkCanvas));
        }

        private void InkCanvas_RectSelectMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_rectSelectDragging) return;
            e.Handled = true;
            EndRectSelect(e.GetPosition(inkCanvas));
        }

        /// <summary>
        /// 结束拖选：矩形与墨迹/图片包围盒相交即选中（PowerPoint/OneNote 框选同标准）。
        /// 图片走 Select 的双参重载（墨迹+元素一起选）——选中后同样进原生选区
        /// （可拖动/缩放/Delete 删除，ImageLayer_DeleteSelectedImages 兼容）。
        /// 拖动距离过小视为"单击空白"→ 清空选中（与原生套索行为一致：空圈取消选中）。
        /// </summary>
        private void EndRectSelect(Point pos)
        {
            _rectSelectDragging = false;
            try { inkCanvas.ReleaseMouseCapture(); } catch { }
            try { Stylus.Capture(null); } catch { }
            RemoveRectAdorner();

            Rect rect = new Rect(_rectSelectStart, pos);
            var hits = new StrokeCollection();
            var hitImages = new System.Collections.Generic.List<UIElement>();
            if (rect.Width > 3 || rect.Height > 3) // 忽略纯单击的零尺寸框
            {
                foreach (Stroke s in inkCanvas.Strokes)
                {
                    try
                    {
                        if (s.GetBounds().IntersectsWith(rect)) hits.Add(s);
                    }
                    catch { }
                }
                // 图片：只框当前视图层可见的（其他页/隐藏层不可选）
                foreach (var img in _imagePageKey.Keys)
                {
                    if (img.Visibility != Visibility.Visible) continue;
                    var b = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                    if (b.IntersectsWith(rect)) hitImages.Add(img);
                }
            }
            // 空集=清空选中（会触发 SelectionChanged → 选区控制条自动收起）
            inkCanvas.Select(hits, hitImages);
        }

        /// <summary>取消进行中的矩形拖选（捕获丢失等异常场景）：清状态+移除预览，不选中</summary>
        private void CancelRectSelectIfDragging()
        {
            if (!_rectSelectDragging) return;
            _rectSelectDragging = false;
            try { Stylus.Capture(null); } catch { }
            RemoveRectAdorner();
        }

        // ---- 预览 Adorner 管理 ----

        /// <summary>显示/更新矩形预览（Adorner 渲染在 inkCanvas 坐标系，与拖选坐标一致）</summary>
        private void ShowRectAdorner(Point start, Point current)
        {
            var layer = AdornerLayer.GetAdornerLayer(inkCanvas);
            if (layer == null) return;
            if (_rectAdorner == null)
            {
                _rectAdorner = new RectSelectAdorner(inkCanvas);
                layer.Add(_rectAdorner);
            }
            _rectAdorner.UpdateRect(new Rect(start, current));
        }

        /// <summary>移除矩形预览</summary>
        private void RemoveRectAdorner()
        {
            if (_rectAdorner == null) return;
            try
            {
                var layer = AdornerLayer.GetAdornerLayer(inkCanvas);
                layer?.Remove(_rectAdorner);
            }
            catch { }
            _rectAdorner = null;
        }

        #endregion
    }

    /// <summary>
    /// 矩形框选预览 Adorner：渲染在 inkCanvas 上层的蓝色虚线矩形 + 淡蓝填充。
    /// IsHitTestVisible=false——纯视觉，不参与命中测试（不挡下方墨迹操作）。
    /// 画刷/画笔冻结提升渲染性能（Adorner 每帧重绘）。
    /// </summary>
    internal class RectSelectAdorner : Adorner
    {
        private Rect _rect;

        // 淡蓝填充（约 10% 不透明度，与面板高亮同色系）
        private static readonly SolidColorBrush FillBrush;
        // 蓝色虚线边框（与选区控制条同款视觉语言）
        private static readonly Pen BorderPen;

        static RectSelectAdorner()
        {
            var fill = new SolidColorBrush(Color.FromArgb(26, 0, 136, 255));
            fill.Freeze();
            FillBrush = fill;

            var pen = new Pen(new SolidColorBrush(Color.FromRgb(0, 120, 215)), 1.2)
            {
                DashStyle = new DashStyle(new DoubleCollection { 3, 2 }, 0)
            };
            pen.Freeze();
            BorderPen = pen;
        }

        public RectSelectAdorner(UIElement adornedElement) : base(adornedElement)
        {
            IsHitTestVisible = false;
        }

        /// <summary>更新预览矩形并重绘</summary>
        public void UpdateRect(Rect r)
        {
            _rect = r;
            InvalidateVisual();
        }

        protected override void OnRender(DrawingContext dc)
        {
            if (_rect.Width > 0 || _rect.Height > 0)
                dc.DrawRectangle(FillBrush, BorderPen, _rect);
        }
    }
}
