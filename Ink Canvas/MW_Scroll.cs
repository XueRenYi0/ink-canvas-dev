using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Point = System.Windows.Point;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：笔记上下滚动（虚拟无限画布）。
    /// 设计见 DEVELOPMENT.md。核心机制：
    /// 采用"坐标物化"方案——滚动时把平移矩阵直接应用到 Stroke 点数据
    /// （inkCanvas.Strokes.Transform(matrix, true)），而不是用渲染变换。
    /// 这样橡皮命中检测、墨迹保存、撤销/重做（TimeMachine 快照与画布
    /// 共享同一批 Stroke 对象引用）都天然保持坐标一致。
    /// 新墨迹直接写在当前可视区域（旧墨迹已物化移走，屏幕自然空白）。
    /// 启用范围：画布可见（正在书写）时——白板/黑板与屏幕/PPT 注释模式；
    /// 画布收起时滚轮穿透给底层应用（PPT 翻页等），互不干扰。
    /// 后续统一快捷键系统可直接注册 ScrollNote(delta)。
    /// </summary>
    public partial class MainWindow
    {
        #region Note Scroll（笔记上下滚动）

        /// <summary>当前累计滚动量：&gt;0 表示已向下滚动（历史墨迹已向上平移）。仅作记账，不参与坐标计算。</summary>
        double _noteScrollOffsetY = 0;

        /// <summary>按钮单次点击的步长：工作区高度的 35%（小半屏）</summary>
        double NoteScrollButtonStep => SystemParameters.WorkArea.Height * 0.35;

        /// <summary>滚轮每格（Delta=120）的步长：工作区高度的 10%，细粒度顺滑</summary>
        double NoteScrollWheelStep => SystemParameters.WorkArea.Height * 0.10;

        /// <summary>滑块行程：轨道高 200 - 滑块高 14 - 上下边距 8 = 178px（与 XAML 胶囊结构对应）</summary>
        const double ScrollThumbTravel = 178;

        /// <summary>滑块拖动中标记（右侧 BorderScrollThumb / 左侧 BorderScrollThumbL 共用一套 handler）</summary>
        bool _isDraggingScrollThumb = false;

        /// <summary>左右胶囊各自是否被悬停（悬停时该侧显形为 1.0，离开回落到基础半透明）</summary>
        bool _scrollHoverRight = false, _scrollHoverLeft = false;

        /// <summary>
        /// 最近一次"滚动活动"时间（滚动 / 拖滑块 / 悬停胶囊都会刷新）。
        /// 胶囊静止超过 ScrollIdleFadeSeconds 就淡到透明（macOS overlay scrollbar 式按需浮现），
        /// 避免常驻在小屏幕边缘碍眼。
        /// </summary>
        DateTime _lastScrollActivity = DateTime.MinValue;

        /// <summary>静止多久后胶囊淡出（秒）。取 3 秒：滚动后留够时间看位置/继续拖，
        /// 又不至于常驻碍眼（1.5 秒实测"消失得太快"）。</summary>
        const double ScrollIdleFadeSeconds = 3.0;

        /// <summary>胶囊基础不透明度：无墨迹更淡（没东西可滚），有墨迹仍保持安静姿态不抢书写视觉</summary>
        double ScrollControlsBaseOpacity => inkCanvas.Strokes.Count == 0 ? 0.15 : 0.35;

        /// <summary>按钮可见性同步定时器（模式切换点分散，统一低频同步，避免逐点挂钩）</summary>
        DispatcherTimer _noteScrollVisibilityTimer;

        /// <summary>
        /// 当前是否允许笔记滚动：
        /// 画布可见（正在书写）即可滚——白板/黑板模式，或屏幕/PPT 注释模式且画布已激活。
        /// 画布收起时滚轮穿透给底层应用（如 PPT 放映翻页），互不干扰。
        /// 后续统一快捷键系统可直接注册 ScrollNote()。
        /// </summary>
        private bool IsNoteScrollActive
        {
            get
            {
                if (inkCanvas.Visibility != Visibility.Visible) return false; // 画布收起：滚轮还给系统
                return currentMode == 0 || currentMode == 1;                  // 屏幕注释 / 白板黑板
            }
        }

        /// <summary>
        /// 初始化滚动功能（由 MainWindow 构造函数调用）：
        /// 1. 墨迹清空（清屏/换页/模式切换中转）时滚动记账归零；
        /// 2. 挂滚轮事件；
        /// 3. 启动按钮可见性同步定时器。
        /// </summary>
        private void InitNoteScroll()
        {
            inkCanvas.Strokes.StrokesChanged += NoteScroll_StrokesChanged;
            inkCanvas.PreviewMouseWheel += InkCanvas_PreviewMouseWheel;

            _noteScrollVisibilityTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _noteScrollVisibilityTimer.Tick += (s, e) =>
            {
                //图片层显隐低频兜底：模式切换点分散（白板/注释/隐藏画板），
                //换页已有即时刷新，其余切换点由这里兜底同步（内部有缓存，无谓遍历零开销）
                ImageLayer_RefreshVisibility();

                // 胶囊只在"画布可见且有可滚内容"时出现；没内容可滚就彻底收起（不占位、不拦鼠标）
                bool scrollable = IsNoteScrollActive && GetMaxScroll() > 0;
                var visibility = scrollable ? Visibility.Visible : Visibility.Collapsed;
                GridNoteScrollControls.Visibility = visibility;
                GridNoteScrollControlsLeft.Visibility = visibility;

                // 静止自动淡出：滚动/悬停/拖滑块后 1.5 秒内保持可见，之后淡到全透明。
                // 用 Opacity=0（而不是 Collapsed）——透明元素仍参与命中测试，鼠标贴到屏幕边缘
                // 划一下就能唤起，不会"找不到入口"。
                bool idle = (DateTime.Now - _lastScrollActivity).TotalSeconds > ScrollIdleFadeSeconds;
                double targetOpacity = idle ? 0.0 : ScrollControlsBaseOpacity;
                if (!_scrollHoverRight) GridNoteScrollControls.Opacity = targetOpacity;
                if (!_scrollHoverLeft) GridNoteScrollControlsLeft.Opacity = targetOpacity;
                // 墨迹增删会改变内容总深度 → 滑块与禁用态随定时器兜底刷新
                UpdateScrollIndicators();
            };
            _noteScrollVisibilityTimer.Start();
        }

        /// <summary>右侧胶囊悬停：显形为完全不透明，离开回落到基础半透明</summary>
        private void GridNoteScrollControls_MouseEnter(object sender, MouseEventArgs e)
        {
            _scrollHoverRight = true;
            _lastScrollActivity = DateTime.Now; // 悬停也算"活动"，鼠标停在胶囊上时不会淡出
            GridNoteScrollControls.Opacity = 1.0;
        }

        private void GridNoteScrollControls_MouseLeave(object sender, MouseEventArgs e)
        {
            _scrollHoverRight = false;
            GridNoteScrollControls.Opacity = ScrollControlsBaseOpacity;
        }

        /// <summary>左侧胶囊悬停：同右侧（两侧独立显形，互不牵连）</summary>
        private void GridNoteScrollControlsLeft_MouseEnter(object sender, MouseEventArgs e)
        {
            _scrollHoverLeft = true;
            _lastScrollActivity = DateTime.Now;
            GridNoteScrollControlsLeft.Opacity = 1.0;
        }

        private void GridNoteScrollControlsLeft_MouseLeave(object sender, MouseEventArgs e)
        {
            _scrollHoverLeft = false;
            GridNoteScrollControlsLeft.Opacity = ScrollControlsBaseOpacity;
        }

        private void NoteScroll_StrokesChanged(object sender, StrokeCollectionChangedEventArgs e)
        {
            // 换页/清屏流程为 ClearStrokes -> RestoreStrokes：
            // 墨迹数归零的瞬间重置滚动记账，新页从顶部开始。
            // 用户擦除全部墨迹时同理（无内容则滚动位置无意义）。
            if (inkCanvas.Strokes.Count == 0)
            {
                _noteScrollOffsetY = 0;
                UpdateScrollIndicators(); // 归零瞬间滑块立即回顶、上箭头立即禁用（不等定时器兜底）
            }
        }

        private void InkCanvas_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!IsNoteScrollActive) return;

            //Ctrl+滚轮：缩放批注（有选区缩选区，无选区缩当前屏幕全部批注），不滚动笔记
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (Math.Abs(e.Delta) >= 1 && ScaleAllOrSelection(e.Delta > 0 ? 1.1 : 0.9))
                    e.Handled = true;
                return;
            }

            // 滚轮向上（正 Delta）= 回看上方历史；向下 = 露出下方空白
            var notches = e.Delta / 120.0;
            if (Math.Abs(notches) < 0.01) return;
            ScrollNote(-notches * NoteScrollWheelStep);
            e.Handled = true;
        }

        private void BtnScrollUp_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(-NoteScrollButtonStep);
        }

        private void BtnScrollDown_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(NoteScrollButtonStep);
        }

        /// <summary>
        /// 滚动画布。delta &gt; 0 向下滚动（历史墨迹上移，露出下方空白区域继续书写）。
        /// 顶部夹取为 0（不允许滚出第一屏上方）；向下不设上限。
        /// 滚动本身不进入撤销栈（仅墨迹操作可撤销）。
        /// </summary>
        private void ScrollNote(double delta)
        {
            // 防御：仅白板/黑板与屏幕注释模式（与按钮可见性逻辑双保险）
            if (!IsNoteScrollActive) return;

            // 记下本次滚动活动时间（胶囊"静止自动淡出"用：滚动后 1.5 秒内保持可见）
            _lastScrollActivity = DateTime.Now;

            var target = Math.Max(0, _noteScrollOffsetY + delta);
            var actual = target - _noteScrollOffsetY;
            if (Math.Abs(actual) < 1) return;

            // 坐标物化：矩阵直接应用到 Stroke 点数据（applyOnStrokes: true）
            // 向下滚动 actual>0，墨迹整体上移（Y 减小）
            var matrix = new Matrix(1, 0, 0, 1, 0, -actual);
            inkCanvas.Strokes.Transform(matrix, true);
            // 图片层同向同量平移（墨迹滚走图片钉原地=不同步 bug，见 MW_ImageLayer.cs）
            ImageLayer_OnScrolled(actual);
            _noteScrollOffsetY = target;

            // 有选区时滚动：笔迹坐标已平移，但 InkCanvas 的选中装饰层（选框/手柄）不会自动重绘，
            // 会停在旧位置 → 选区与墨迹分离。用"程序化重选"刷一次装饰层（清空再选回同一批笔迹，
            // 坐标已更新，装饰层随新坐标重画）。选中与滚动因此可以共存，不再强行取消选中。
            RefreshSelectionAdorner();

            UpdateScrollIndicators(); // 滚动即时刷新位置指示与禁用态
        }

        /// <summary>
        /// 滚动后刷新选区 UI 位置。
        ///
        /// 关键：本项目的选区框 / 操作条 / 旋转钮 / 命中区都是**自定义元素**
        /// （GridInkCanvasSelectionCover 那一套，不是 WPF 原生 SelectionAdorner），
        /// 它们不会因为笔迹坐标变化而自动跟随——必须手动调统一的刷新方法。
        /// （最初试的"清空再重选"只对原生装饰层有效，对自定义框没用，所以选框仍旧留在原地。）
        /// </summary>
        private void RefreshSelectionAdorner()
        {
            try
            {
                // 墨迹 **或图片** 任一被选中都要刷新：
                // 纯图片/纯图形选中时 GetSelectedStrokes() 是空的——原来只守墨迹的话，
                // 这类选中的框就会留在原地不跟随滚动（选框几何由 GetGestureSelectionBounds 统一算，
                // 它本身是同时包含墨迹与图片的）。
                if (inkCanvas.GetSelectedStrokes().Count == 0 && GetSelectedPageImages().Count == 0) return;
                updateBorderStrokeSelectionControlLocation();
            }
            catch { }
        }

        #region 按页记忆滚动位置

        /// <summary>
        /// 按页记忆的滚动偏移（索引与 MW_WhiteboardControls.TimeMachineHistories 同构：
        /// 1..99 = 白板页；0 保留给"非白板墨迹"，不使用）。0 值即"未滚过/在顶部"，无需额外存在性标记。
        ///
        /// 为什么必须按页记：滚动是"坐标物化"——滚动过的坐标会连同页面历史一起被保存
        /// （TimeMachineHistories[页]）。而 _noteScrollOffsetY 若仍是全局单值，换页时会被
        /// ClearStrokes 触发的归零清掉，切回原页就出现"内容位置对不上、且被滚上去的内容再也滚不回来"。
        /// 记住每页偏移后，切回来记账值与笔迹坐标天然吻合。
        /// </summary>
        double[] _pageScrollOffsets = new double[101];

        /// <summary>换页前记下当前页的滚动位置（由 MW_WhiteboardControls.SaveStrokes 调用）。
        /// 索引 0 = 非白板（批注层）墨迹，与白板页同样记忆——两者的笔迹坐标都是物化的。</summary>
        private void NoteScroll_SavePageOffset(int pageIndex)
        {
            if (pageIndex < 0 || pageIndex >= _pageScrollOffsets.Length) return;
            _pageScrollOffsets[pageIndex] = _noteScrollOffsetY;
        }

        /// <summary>
        /// 换页后恢复目标页的滚动位置（由 MW_WhiteboardControls.RestoreStrokes 调用）。
        /// 只需把记账值设回去——该页笔迹坐标本身就带着那次滚动的位移，两者正好对应，不用再 Transform。
        /// </summary>
        private void NoteScroll_RestorePageOffset(int pageIndex)
        {
            _noteScrollOffsetY = (pageIndex >= 0 && pageIndex < _pageScrollOffsets.Length)
                ? _pageScrollOffsets[pageIndex]
                : 0;
            UpdateScrollIndicators();
        }

        #endregion

        /// <summary>
        /// 内容最大可滚深度：墨迹最低点（还原为虚拟坐标）超出首屏的部分 + 书写余量。
        /// 墨迹为物化坐标（滚动时已平移），bounds.Bottom + offset 还原为虚拟画布坐标。
        /// 供滑块比例计算与拖动跳转（拖到哪滚到哪）共用。
        /// </summary>
        private double GetMaxScroll()
        {
            var bounds = inkCanvas.Strokes.GetBounds();
            // 图片底边也计入内容深度（纯图片页没写字也能滚，否则长图下半截永远看不到）
            double imgBottom = ImageLayer_GetContentBottom();
            double contentBottom;
            if (bounds.IsEmpty)
            {
                if (double.IsNaN(imgBottom)) return 0;
                contentBottom = imgBottom;
            }
            else
            {
                contentBottom = bounds.Bottom;
                if (!double.IsNaN(imgBottom) && imgBottom > contentBottom) contentBottom = imgBottom;
            }
            // 墨迹为物化坐标（滚动时已平移），还原为虚拟画布坐标
            contentBottom += _noteScrollOffsetY;
            return Math.Max(0, contentBottom - SystemParameters.WorkArea.Height + 160); // 留书写余量
        }

        /// <summary>
        /// 刷新滚动控件的指示状态（左右两侧同步）：
        /// 1. 滑块位置 = 当前偏移 / 最大可滚深度；
        /// 2. 边界禁用态：滚到顶时上箭头淡化（下方向无限延伸，永不禁用）。
        /// </summary>
        private void UpdateScrollIndicators()
        {
            // 边界禁用态（两侧同步）
            PathScrollUp.Opacity = _noteScrollOffsetY <= 0.5 ? 0.35 : 1.0;
            PathScrollUpL.Opacity = PathScrollUp.Opacity;

            // 滑块位置：ratio = 当前偏移 / 最大可滚深度
            double ratio = 0;
            double maxScroll = GetMaxScroll();
            if (maxScroll > 0 && _noteScrollOffsetY > 0)
                ratio = Math.Min(1.0, _noteScrollOffsetY / maxScroll);
            // 指示区高 120、滑块高 14、上下边距 4 → 行程 ScrollThumbTravel=98（两侧同步）
            var thumbMargin = new Thickness(0, 4 + ratio * ScrollThumbTravel, 0, 0);
            BorderScrollThumb.Margin = thumbMargin;
            BorderScrollThumbL.Margin = thumbMargin;
        }

        #region 滑块拖动（左右胶囊共用一套 handler）

        /// <summary>滑块按下：捕获鼠标并立即跳到按压位置（无内容可滚时忽略）</summary>
        private void BorderScrollThumb_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            if (GetMaxScroll() <= 0) return; // 没有可滚内容，拖动无意义
            _isDraggingScrollThumb = true;
            _lastScrollActivity = DateTime.Now; // 拖动开始也算活动（按住不动时也不淡出）
            var thumb = (Border)sender;
            try { thumb.CaptureMouse(); } catch { }
            ScrollThumbFollowPointer(thumb, e.GetPosition((FrameworkElement)thumb.Parent));
            e.Handled = true;
        }

        /// <summary>拖动中：滑块中心跟随鼠标，滚动实时映射（ScrollNote 内部即时刷新指示）</summary>
        private void BorderScrollThumb_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingScrollThumb) return;
            var thumb = (Border)sender;
            ScrollThumbFollowPointer(thumb, e.GetPosition((FrameworkElement)thumb.Parent));
        }

        /// <summary>松开：释放捕获，指示回到精确比例位置</summary>
        private void BorderScrollThumb_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingScrollThumb) return;
            _isDraggingScrollThumb = false;
            try { ((Border)sender).ReleaseMouseCapture(); } catch { }
            UpdateScrollIndicators();
        }

        /// <summary>
        /// 鼠标位置 → 滚动位置：滑块中心对齐鼠标 Y，换算为轨道比例，
        /// 再乘最大可滚深度得到目标偏移，经 ScrollNote 跳转（矩阵物化一次完成）。
        /// </summary>
        private void ScrollThumbFollowPointer(Border thumb, Point pos)
        {
            // 滑块高 14：让滑块中心对准鼠标 → 顶边距 = y - 7；有效范围 [4, 4+Travel]
            double desiredTop = pos.Y - 7;
            double ratio = Math.Max(0, Math.Min(1, (desiredTop - 4) / ScrollThumbTravel));
            // 拖到哪滚到哪：目标偏移 = 比例 × 最大深度（顶部夹取由 ScrollNote 内部完成）
            ScrollNote(ratio * GetMaxScroll() - _noteScrollOffsetY);
        }

        #endregion

        #endregion
    }
}
