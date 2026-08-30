using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern;
using iNKORE.UI.WPF.Modern.Helpers;
using IWshRuntimeLibrary;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic;
using Microsoft.Win32;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Timers;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Input.StylusPlugIns;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Application = System.Windows.Application;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using Timer = System.Timers.Timer;

namespace Ink_Canvas
{
    /// <summary>MainWindow 分部类：墨迹选区与手势（含浮动控件）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Selection Gestures

        #region Floating Control

        object lastBorderMouseDownObject;

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            lastBorderMouseDownObject = sender;
        }

        private void BorderStrokeSelectionClone_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            //点击即克隆：立即复制一份选中墨迹，偏移一段距离后自动选中副本，
            //用户可直接拖走（原交互为开关模式，拖动时才复制，易忘关、不直观）
            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count == 0) return;

            var cloned = strokes.Clone();
            var m = new Matrix();
            m.Translate(24, 24); // 副本向右下偏移，与原件错开可见
            cloned.Transform(m, false);

            isProgramChangeStrokeSelection = true;
            inkCanvas.Select(new StrokeCollection());
            isProgramChangeStrokeSelection = false;
            inkCanvas.Strokes.Add(cloned);
            inkCanvas.Select(cloned);
        }

        private void BorderStrokeSelectionCloneToNewBoard_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            var strokes = inkCanvas.GetSelectedStrokes();
            inkCanvas.Select(new StrokeCollection());
            strokes = strokes.Clone();
            BtnWhiteBoardAdd_Click(null, null);
            inkCanvas.Strokes.Add(strokes);
        }

        private void BorderStrokeSelectionDelete_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject == sender)
            {
                SymbolIconDelete_MouseUp(sender, e);
            }
        }

        private void GridPenWidthDecrease_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ChangeStrokeThickness(0.8);
        }

        private void GridPenWidthIncrease_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ChangeStrokeThickness(1.25);
        }

        private void ChangeStrokeThickness(double multipler)
        {
            foreach (Stroke stroke in inkCanvas.GetSelectedStrokes())
            {
                //stroke.DrawingAttributes.Width *= 1.25;
                //stroke.DrawingAttributes.Height *= 1.25;

                var newWidth = stroke.DrawingAttributes.Width * multipler;
                var newHeight = stroke.DrawingAttributes.Height * multipler;

                if (newWidth >= DrawingAttributes.MinWidth && newWidth <= DrawingAttributes.MaxWidth
                    && newHeight >= DrawingAttributes.MinHeight && newHeight <= DrawingAttributes.MaxHeight)
                {
                    stroke.DrawingAttributes.Width = newWidth;
                    stroke.DrawingAttributes.Height = newHeight;
                }
            }
            if (DrawingAttributesHistory.Count > 0)
            {

                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
        }

        private void GridPenWidthRestore_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            foreach (Stroke stroke in inkCanvas.GetSelectedStrokes())
            {
                stroke.DrawingAttributes.Width = inkCanvas.DefaultDrawingAttributes.Width;
                stroke.DrawingAttributes.Height = inkCanvas.DefaultDrawingAttributes.Height;
            }
        }

        private void ImageFlipHorizontal_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Width / 2,
                inkCanvas.GetSelectionBounds().Top + inkCanvas.GetSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.ScaleAt(-1, 1, center.X, center.Y);  // 缩放

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }
            if (DrawingAttributesHistory.Count > 0)
            {
                var collecion = new StrokeCollection();
                foreach (var item in DrawingAttributesHistory)
                {
                    collecion.Add(item.Key);
                }
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
            //updateBorderStrokeSelectionControlLocation();
        }

        private void ImageFlipVertical_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Width / 2,
                inkCanvas.GetSelectionBounds().Top + inkCanvas.GetSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.ScaleAt(1, -1, center.X, center.Y);  // 缩放

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }
            if (DrawingAttributesHistory.Count > 0)
            {
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
        }

        private void ImageRotate45_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Width / 2,
                inkCanvas.GetSelectionBounds().Top + inkCanvas.GetSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.RotateAt(15, center.X, center.Y);  // 旋转（原 45°，改为 15° 细步进，教学场景更实用）

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }
            if (DrawingAttributesHistory.Count > 0)
            {
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
        }

        private void ImageRotate90_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Width / 2,
                inkCanvas.GetSelectionBounds().Top + inkCanvas.GetSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.RotateAt(90, center.X, center.Y);  // 旋转

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }
            if (DrawingAttributesHistory.Count > 0)
            {
                var collecion = new StrokeCollection();
                foreach (var item in DrawingAttributesHistory)
                {
                    collecion.Add(item.Key);
                }
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
        }

        #endregion


        bool isGridInkCanvasSelectionCoverMouseDown = false;
        StrokeCollection StrokesSelectionClone = new StrokeCollection();

        //鼠标拖动选区状态（触摸走 Manipulation 事件，鼠标/数位笔走此路径；dec>0 表示触摸进行中，让位）
        bool isMouseSelectionDragging = false;
        bool hasMouseSelectionDragMoved = false; // 按下后是否真的移动过（区分"拖动"与"单击取消选区"）
        Point lastMousePointOnSelectionCover = new Point(0, 0);

        private void GridInkCanvasSelectionCover_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isGridInkCanvasSelectionCoverMouseDown = true;
            if (e.ChangedButton != MouseButton.Left || dec.Count != 0) return;

            //缩放手柄命中：走手柄拖动路径（InkCanvas 原生把手被本覆盖层遮挡，此处实现同语义的拖动缩放）
            var handlePos = e.GetPosition(inkCanvas);
            var handle = HitTestSelectionHandle(handlePos);
            if (handle != SelectionHandleKind.None)
            {
                StartHandleDrag(handle, handlePos);
                try { GridInkCanvasSelectionCover.CaptureMouse(); } catch { }
                return;
            }

            //克隆已改为按钮点击即生成（见 BorderStrokeSelectionClone_MouseUp），此处仅负责拖动

            //【框内才能拖】与 PowerPoint 等全行业一致：按下点在选中笔迹包围盒内（含 10px 容差，
            //方便抓细线）才进入拖动；框外按下不遥控选中物——想在别处落笔时不会误拽走整个图形
            var dragBounds = inkCanvas.GetSelectedStrokes().GetBounds();
            dragBounds.Inflate(10, 10);
            if (!dragBounds.Contains(handlePos))
            {
                //框外按下：结束本层拖动意图，放行事件让 MouseUp 走"单击取消选中"路径
                isMouseSelectionDragging = false;
                return;
            }

            //开始鼠标拖动（捕获鼠标保证移出窗口也能收到 Move/Up）
            lastMousePointOnSelectionCover = e.GetPosition(null);
            isMouseSelectionDragging = true;
            hasMouseSelectionDragMoved = false;
            try { GridInkCanvasSelectionCover.CaptureMouse(); } catch { }
        }

        private void GridInkCanvasSelectionCover_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isMouseSelectionDragging) return;
            if (e.LeftButton != MouseButtonState.Pressed) { FinishMouseSelectionDrag(); return; }

            //手柄拖动缩放路径（优先于整体平移）
            if (_activeHandleDragKind != SelectionHandleKind.None)
            {
                UpdateHandleDrag(e.GetPosition(inkCanvas));
                return;
            }

            var pos = e.GetPosition(null);
            var dx = pos.X - lastMousePointOnSelectionCover.X;
            var dy = pos.Y - lastMousePointOnSelectionCover.Y;
            lastMousePointOnSelectionCover = pos;
            if (Math.Abs(dx) < 0.1 && Math.Abs(dy) < 0.1) return;
            hasMouseSelectionDragMoved = true; // 超过阈值的移动才算拖动

            //与触摸路径一致：克隆时拖动副本，否则拖动选中墨迹
            StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
            if (StrokesSelectionClone.Count != 0) strokes = StrokesSelectionClone;

            var m = new Matrix();
            m.Translate(dx, dy);
            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false);
            }

            //克隆拖动时选区（原件）未动，控制条无需跟随
            if (StrokesSelectionClone.Count == 0) updateBorderStrokeSelectionControlLocation();
        }

        private void GridInkCanvasSelectionCover_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!isGridInkCanvasSelectionCoverMouseDown) return;
            isGridInkCanvasSelectionCoverMouseDown = false;

            bool wasHandleDrag = _activeHandleDragKind != SelectionHandleKind.None;

            if (isMouseSelectionDragging && hasMouseSelectionDragMoved)
            {
                //真实拖动结束（平移或手柄缩放）：保持选区可见，提交撤销历史
                FinishMouseSelectionDrag();
            }
            else
            {
                //结束拖动状态；手柄上原地点击未拖动 → 保持选区（空白处单击才取消选区）
                if (isMouseSelectionDragging) FinishMouseSelectionDrag();
                if (wasHandleDrag) return;
                //单击（未拖动）：取消选区（原行为）
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection());
                isProgramChangeStrokeSelection = false;
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                TryEndOneShotSelection();
            }
        }

        /// <summary>结束鼠标拖动：释放捕获、清空克隆引用、提交撤销历史（同触摸路径 ManipulationCompleted）</summary>
        private void FinishMouseSelectionDrag()
        {
            isMouseSelectionDragging = false;
            _activeHandleDragKind = SelectionHandleKind.None; // 结束手柄缩放会话
            try { if (GridInkCanvasSelectionCover.IsMouseCaptured) GridInkCanvasSelectionCover.ReleaseMouseCapture(); } catch { }
            StrokesSelectionClone = new StrokeCollection();

            if (StrokeManipulationHistory?.Count > 0)
            {
                timeMachine.CommitStrokeManipulationHistory(StrokeManipulationHistory);
                foreach (var item in StrokeManipulationHistory)
                {
                    StrokeInitialHistory[item.Key] = item.Value.Item2;
                }
                StrokeManipulationHistory = null;
            }

            //手柄缩放会同步改笔画粗细：结束时一并提交（平移路径不产生该历史，无副作用）
            CommitDrawingAttributesHistoryNow();
        }

        #region 选区手柄拖动缩放 + Ctrl+滚轮缩放

        /// <summary>选区缩放手柄类型：4 角 + 4 边中点（与 InkCanvas 原生把手位置一致）+ 顶部旋转钮</summary>
        private enum SelectionHandleKind
        {
            None,
            TopLeft, Top, TopRight, Right, BottomRight, Bottom, BottomLeft, Left,
            Rotate //顶部中央伸出的旋转手柄（PowerPoint/Figma 惯例位置）
        }

        SelectionHandleKind _activeHandleDragKind = SelectionHandleKind.None;
        Rect _handleDragStartBounds = Rect.Empty;
        double _rotateDragStartAngle = 0.0; //旋转拖动开始时鼠标相对选区中心的方位角（弧度）

        /// <summary>旋转钮在选中框顶边上方伸出的距离（DIP）</summary>
        const double RotateHandleOffset = 22.0;

        /// <summary>手柄命中半径（DIP）。InkCanvas 原生把手视觉直径约 8，放宽到 14 便于鼠标/数位笔抓取</summary>
        const double SelectionHandleHitRadius = 14.0;

        /// <summary>
        /// 命中检测：判断 pos（inkCanvas 坐标）是否落在选区缩放手柄（小圆点）上。
        /// InkCanvas 原生把手被本覆盖层遮挡无法拖动，这里在覆盖层实现同语义交互。
        /// </summary>
        private SelectionHandleKind HitTestSelectionHandle(Point pos)
        {
            Rect b = inkCanvas.GetSelectionBounds();
            if (b.IsEmpty || b.Width <= 0 || b.Height <= 0) return SelectionHandleKind.None;

            //旋转钮：顶部中央向上伸出（优先于缩放手柄判定，位置在框外不冲突）
            var rotateCenter = new Point(b.Left + b.Width / 2, b.Top - RotateHandleOffset);
            if ((pos - rotateCenter).Length <= SelectionHandleHitRadius) return SelectionHandleKind.Rotate;

            var candidates = new Tuple<SelectionHandleKind, Point>[]
            {
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.TopLeft,      new Point(b.Left, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Top,         new Point(b.Left + b.Width / 2, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.TopRight,    new Point(b.Right, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Right,       new Point(b.Right, b.Top + b.Height / 2)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.BottomRight, new Point(b.Right, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Bottom,      new Point(b.Left + b.Width / 2, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.BottomLeft,  new Point(b.Left, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Left,        new Point(b.Left, b.Top + b.Height / 2)),
            };

            SelectionHandleKind best = SelectionHandleKind.None;
            double bestDist = SelectionHandleHitRadius;
            foreach (var c in candidates)
            {
                double d = (pos - c.Item2).Length;
                if (d <= bestDist) { best = c.Item1; bestDist = d; }
            }
            return best;
        }

        /// <summary>
        /// 开始手柄拖动：记录起始选区 bounds。复用 isMouseSelectionDragging 状态，
        /// 使 Stroke_StylusPointsChanged 在拖动期间只累计撤销历史、松开时一次提交
        /// （否则一次拖动会碎片化为多个撤销步骤）。
        /// </summary>
        private void StartHandleDrag(SelectionHandleKind kind, Point pos)
        {
            _activeHandleDragKind = kind;
            _handleDragStartBounds = inkCanvas.GetSelectionBounds();
            isMouseSelectionDragging = true;
            hasMouseSelectionDragMoved = false;
            lastMousePointOnSelectionCover = pos;

            //旋转分支：记录初始方位角（鼠标 → 选区中心的方向），后续按角度增量旋转
            if (kind == SelectionHandleKind.Rotate)
            {
                var b = _handleDragStartBounds;
                var center = new Point(b.Left + b.Width / 2, b.Top + b.Height / 2);
                _rotateDragStartAngle = Math.Atan2(pos.Y - center.Y, pos.X - center.X);
            }
        }

        /// <summary>
        /// 手柄拖动中：以起始 bounds 的对角/对边为固定锚点，把选区缩放到鼠标当前位置。
        /// 角点手柄 = 等比缩放（保持宽高比，手写批注多为文字/图形，自由拉伸易变形），
        /// 边中点手柄 = 单轴缩放（需要"只拉宽/压扁"时用，与 PPT 语义一致）。
        /// 每次移动按"当前 bounds → 目标尺寸"计算增量比例，围绕同一锚点复合即得总变换。
        /// </summary>
        private void UpdateHandleDrag(Point pos)
        {
            if (_activeHandleDragKind == SelectionHandleKind.None) return;
            Rect start = _handleDragStartBounds;
            if (start.IsEmpty || start.Width < 0.5 || start.Height < 0.5) return;

            //旋转手柄分支：按"鼠标绕选区中心的方位角增量"整体旋转选中笔迹。
            //旋转中心固定取拖动起始 bounds 的中心（整次拖动不漂移）；
            //角度跨过 ±180° 射线时按最短方向修正，避免图形瞬间反向猛转。
            if (_activeHandleDragKind == SelectionHandleKind.Rotate)
            {
                StrokeCollection rotStrokes = inkCanvas.GetSelectedStrokes();
                if (rotStrokes.Count == 0) return;

                var center = new Point(start.Left + start.Width / 2, start.Top + start.Height / 2);
                double currentAngle = Math.Atan2(pos.Y - center.Y, pos.X - center.X); //鼠标当前方位角（弧度）
                double delta = currentAngle - _rotateDragStartAngle;                    //相对上次的角度增量
                _rotateDragStartAngle = currentAngle;                                  //滚动累计，供下一帧
                if (delta > Math.PI) delta -= 2 * Math.PI;   //跨线修正：+180° 方向跳变 → 取最短转向
                if (delta < -Math.PI) delta += 2 * Math.PI;
                if (Math.Abs(delta) < 0.001) return;          //忽略亚像素级微动，省一次矩阵变换

                var rm = new Matrix();
                rm.RotateAt(delta * 180.0 / Math.PI, center.X, center.Y); //弧度 → 角度，绕中心旋转
                foreach (Stroke stroke in rotStrokes)
                {
                    stroke.Transform(rm, false); //StylusPointsChanged 自动累计撤销历史（拖动期间不提交）
                }

                hasMouseSelectionDragMoved = true;
                updateBorderStrokeSelectionControlLocation(); //旋转钮/工具条跟随新选区
                return;
            }

            //固定锚点 = 被拖动手柄的对角（角点）或对边（边中点），整次拖动保持不动
            double anchorX = 0, anchorY = 0;
            bool isCorner = false, dragX = false, dragY = false;
            switch (_activeHandleDragKind)
            {
                case SelectionHandleKind.TopLeft:      anchorX = start.Right; anchorY = start.Bottom; isCorner = true; break;
                case SelectionHandleKind.TopRight:     anchorX = start.Left;  anchorY = start.Bottom; isCorner = true; break;
                case SelectionHandleKind.BottomRight:  anchorX = start.Left;  anchorY = start.Top;    isCorner = true; break;
                case SelectionHandleKind.BottomLeft:   anchorX = start.Right; anchorY = start.Top;    isCorner = true; break;
                case SelectionHandleKind.Top:          anchorY = start.Bottom; dragY = true; break;
                case SelectionHandleKind.Bottom:       anchorY = start.Top;    dragY = true; break;
                case SelectionHandleKind.Left:         anchorX = start.Right;  dragX = true; break;
                case SelectionHandleKind.Right:        anchorX = start.Left;   dragX = true; break;
                default: return;
            }

            //目标尺寸 = 鼠标到锚点距离（下限 2 DIP，防翻转/退化），按当前 bounds 换算增量比例
            const double MinSize = 2.0;
            Rect cur = inkCanvas.GetSelectionBounds();
            if (cur.IsEmpty || cur.Width < 0.5 || cur.Height < 0.5) return;

            double fx = 1.0, fy = 1.0;
            if (isCorner)
            {
                //等比：统一因子 = 鼠标到锚点距离 / 当前被拖角到锚点距离（沿对角线跟手）
                Point curCorner;
                switch (_activeHandleDragKind)
                {
                    case SelectionHandleKind.TopLeft:     curCorner = new Point(cur.Left, cur.Top); break;
                    case SelectionHandleKind.TopRight:    curCorner = new Point(cur.Right, cur.Top); break;
                    case SelectionHandleKind.BottomRight: curCorner = new Point(cur.Right, cur.Bottom); break;
                    default:                              curCorner = new Point(cur.Left, cur.Bottom); break;
                }
                double dCur = (curCorner - new Point(anchorX, anchorY)).Length;
                if (dCur < 0.5) return;
                double f = (pos - new Point(anchorX, anchorY)).Length / dCur;
                if (f < 0.02) f = 0.02; //防缩至消失
                fx = fy = f;
            }
            else
            {
                if (dragX) fx = Math.Max(MinSize, Math.Abs(pos.X - anchorX)) / cur.Width;
                if (dragY) fy = Math.Max(MinSize, Math.Abs(pos.Y - anchorY)) / cur.Height;
            }
            if (Math.Abs(fx - 1) < 0.001 && Math.Abs(fy - 1) < 0.001) return;

            StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count == 0) return;

            var m = new Matrix();
            m.ScaleAt(fx, fy, anchorX, anchorY);
            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false); //StylusPointsChanged 自动累计撤销历史（拖动期间不提交）
                try
                {
                    //笔画粗细随选区同步缩放（与触摸双指缩放一致），夹取到 WPF 允许范围
                    double w = Math.Max(DrawingAttributes.MinWidth, Math.Min(DrawingAttributes.MaxWidth, stroke.DrawingAttributes.Width * fx));
                    double h = Math.Max(DrawingAttributes.MinHeight, Math.Min(DrawingAttributes.MaxHeight, stroke.DrawingAttributes.Height * fy));
                    stroke.DrawingAttributes.Width = w;
                    stroke.DrawingAttributes.Height = h;
                }
                catch { }
            }

            hasMouseSelectionDragMoved = true;
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>悬停反馈：手柄上显示缩放箭头；框内显示移动光标（四向箭头）；框外恢复默认</summary>
        private void GridInkCanvasSelectionCover_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseSelectionDragging || dec.Count != 0) return; // 拖动/触摸进行中保持当前光标
            if (inkCanvas.GetSelectedStrokes().Count == 0) return;

            var pos = e.GetPosition(inkCanvas);

            //优先手柄（缩放光标）
            var handleCursor = SelectionHandleCursor(HitTestSelectionHandle(pos));
            if (handleCursor != null)
            {
                GridInkCanvasSelectionCover.Cursor = handleCursor;
                return;
            }

            //框内（与 MouseDown 的拖动判定同一套几何：包围盒 + 10px 容差）→ 四向移动光标
            //光标可供性：告诉用户"这里按住可以拖走"，和实际能拖的范围严格一致
            var bounds = inkCanvas.GetSelectedStrokes().GetBounds();
            bounds.Inflate(10, 10);
            GridInkCanvasSelectionCover.Cursor = bounds.Contains(pos) ? Cursors.SizeAll : null;
        }

        private Cursor SelectionHandleCursor(SelectionHandleKind kind)
        {
            switch (kind)
            {
                case SelectionHandleKind.TopLeft:
                case SelectionHandleKind.BottomRight:
                    return Cursors.SizeNWSE;
                case SelectionHandleKind.TopRight:
                case SelectionHandleKind.BottomLeft:
                    return Cursors.SizeNESW;
                case SelectionHandleKind.Top:
                case SelectionHandleKind.Bottom:
                    return Cursors.SizeNS;
                case SelectionHandleKind.Left:
                case SelectionHandleKind.Right:
                    return Cursors.SizeWE;
                case SelectionHandleKind.Rotate:
                    return Cursors.ScrollAll; //旋转手柄（圆圈箭头，最接近旋转语义的系统光标）
                default:
                    return null; // 非手柄区域恢复默认光标
            }
        }

        /// <summary>Ctrl+滚轮：缩放批注（有选区缩放选区，无选区缩放整屏；与 Ctrl+加减号同语义，可撤销）</summary>
        private void GridInkCanvasSelectionCover_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control) return;
            if (Math.Abs(e.Delta) < 1) return;

            if (ScaleAllOrSelection(e.Delta > 0 ? 1.1 : 0.9)) e.Handled = true;
        }

        #endregion 选区手柄拖动缩放 + Ctrl+滚轮缩放

        #region 选中缩放/还原

        //选中快照：SelectionChanged 捕获（还原 = 恢复到本次选中时的状态）
        Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>> SelectionSnapshot;

        private void GridSelectionScaleUp_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            ScaleSelection(1.1);
        }

        private void GridSelectionScaleDown_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            ScaleSelection(0.9);
        }

        /// <summary>以选区中心整体缩放选中墨迹（浮动工具条"放大/缩小"按钮用）</summary>
        private void ScaleSelection(double factor)
        {
            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count == 0) return;
            Rect bounds = inkCanvas.GetSelectionBounds();
            if (ScaleStrokes(strokes, new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2), factor))
                updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>
        /// 统一缩放入口（Ctrl+滚轮 / Ctrl+加减号）：
        /// 有选区 → 缩放选中批注（绕选区中心）；无选区 → 缩放当前屏幕全部批注（绕屏幕中心）。
        /// 返回是否执行了缩放（无目标时 false，调用方决定是否放行事件）。
        /// </summary>
        private bool ScaleAllOrSelection(double factor)
        {
            if (inkCanvas.Visibility != Visibility.Visible) return false;

            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count > 0)
            {
                Rect bounds = inkCanvas.GetSelectionBounds();
                if (ScaleStrokes(strokes, new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2), factor))
                    updateBorderStrokeSelectionControlLocation();
                return true;
            }

            //无选区：整屏批注绕屏幕中心缩放
            if (inkCanvas.Strokes.Count == 0) return false;
            return ScaleStrokes(inkCanvas.Strokes,
                new Point(inkCanvas.ActualWidth / 2, inkCanvas.ActualHeight / 2), factor);
        }

        /// <summary>缩放一批笔画：点坐标绕 center 缩放、笔画粗细同步（夹取 WPF 范围）、提交撤销历史</summary>
        private bool ScaleStrokes(StrokeCollection strokes, Point center, double factor)
        {
            if (strokes == null || strokes.Count == 0) return false;

            //防误触缩没：已经小到 1 DIP 以内不再继续缩小
            if (factor < 1)
            {
                var b = strokes.GetBounds();
                if (b.Width < 1 && b.Height < 1) return false;
            }

            var m = new Matrix();
            m.ScaleAt(factor, factor, center.X, center.Y);

            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false); //触发 StylusPointsChanged，自动进撤销历史
                try
                {
                    double w = stroke.DrawingAttributes.Width * factor;
                    double h = stroke.DrawingAttributes.Height * factor;
                    //超出 WPF 允许范围则夹取（点坐标照常缩放，仅笔宽受限）
                    w = Math.Max(DrawingAttributes.MinWidth, Math.Min(DrawingAttributes.MaxWidth, w));
                    h = Math.Max(DrawingAttributes.MinHeight, Math.Min(DrawingAttributes.MaxHeight, h));
                    stroke.DrawingAttributes.Width = w;
                    stroke.DrawingAttributes.Height = h;
                }
                catch { }
            }

            CommitDrawingAttributesHistoryNow();
            return true;
        }

        private void GridSelectionScaleRestore_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count == 0 || SelectionSnapshot == null) return;

            var history = new Dictionary<Stroke, Tuple<StylusPointCollection, StylusPointCollection>>();
            foreach (Stroke s in strokes)
            {
                if (!SelectionSnapshot.TryGetValue(s, out var snap)) continue;
                try
                {
                    var oldPts = s.StylusPoints.Clone();
                    s.StylusPoints = snap.Item1; //赋值触发 StylusPointsReplaced，不自动提交历史，下面手动提交
                    s.DrawingAttributes.Width = snap.Item2.Width;
                    s.DrawingAttributes.Height = snap.Item2.Height;
                    history[s] = new Tuple<StylusPointCollection, StylusPointCollection>(oldPts, s.StylusPoints.Clone());
                }
                catch { }
            }

            if (history.Count > 0)
            {
                timeMachine.CommitStrokeManipulationHistory(history);
                foreach (var item in history)
                {
                    StrokeInitialHistory[item.Key] = item.Value.Item2;
                }
            }
            CommitDrawingAttributesHistoryNow();
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>立即提交笔画粗细变更历史（避免滞留到下次拖动/缩放时才一并提交，导致撤销步骤混乱）</summary>
        private void CommitDrawingAttributesHistoryNow()
        {
            try
            {
                if (DrawingAttributesHistory.Count > 0)
                {
                    timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                    DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                    foreach (var item in DrawingAttributesHistoryFlag) item.Value.Clear();
                }
            }
            catch { }
        }

        #endregion 选中缩放/还原


        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            forceEraser = true;

            //用户主动切换到选择工具：清掉"一次性选中"标志，
            //之后取消选中时不再自动恢复笔模式（详见 MW_GraphStrokes.cs）
            _isOneShotGraphSelection = false;

            #region 选中快照（用于"还原"按钮：恢复到本次选中时的大小/位置/粗细）

            try
            {
                SelectionSnapshot = new Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>>();
                foreach (Stroke s in inkCanvas.GetSelectedStrokes())
                {
                    SelectionSnapshot[s] = new Tuple<StylusPointCollection, DrawingAttributes>(
                        s.StylusPoints.Clone(), s.DrawingAttributes.Clone());
                }
            }
            catch { }

            #endregion
            drawingShapeMode = 0;
            UpdateShapeIconHighlight(); //切到选择工具时熄灭图形图标高亮
            inkCanvas.IsManipulationEnabled = false;
            if (inkCanvas.EditingMode == InkCanvasEditingMode.Select)
            {
                if (inkCanvas.GetSelectedStrokes().Count == inkCanvas.Strokes.Count)
                {
                    inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                    inkCanvas.IsManipulationEnabled = true;
                }
                else
                {
                    //inkCanvas.Select(inkCanvas.Strokes);
                    // Fixed bug: 当通过如鼠标点击等某些方式创建没有高度或长度的笔画时，全选功能不能使用克隆、旋转、翻转、调整笔画粗细、删除功能
                    StrokeCollection selectedStrokes = new StrokeCollection();
                    foreach (Stroke stroke in inkCanvas.Strokes)
                    {
                        if (stroke.GetBounds().Width > 0 && stroke.GetBounds().Height > 0)
                        {
                            selectedStrokes.Add(stroke);
                        }
                    }
                    inkCanvas.Select(selectedStrokes);
                }
            }
            else
            {
                inkCanvas.EditingMode = InkCanvasEditingMode.Select;

            }
        }

        double BorderStrokeSelectionControlWidth = 490.0;
        double BorderStrokeSelectionControlHeight = 80.0;
        bool isProgramChangeStrokeSelection = false;

        private void inkCanvas_SelectionChanged(object sender, EventArgs e)
        {
            if (isProgramChangeStrokeSelection) return;
            if (inkCanvas.GetSelectedStrokes().Count == 0)
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                TryEndOneShotSelection();
            }
            else
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Visible;

                //捕获选中快照（还原按钮用：恢复到选中时的大小/位置/粗细）
                try
                {
                    SelectionSnapshot = new Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>>();
                    foreach (Stroke s in inkCanvas.GetSelectedStrokes())
                    {
                        SelectionSnapshot[s] = new Tuple<StylusPointCollection, DrawingAttributes>(
                            s.StylusPoints.Clone(), s.DrawingAttributes.Clone());
                    }
                }
                catch { }

                updateBorderStrokeSelectionControlLocation();
            }
        }

        /// <summary>
        /// 选区可命中区域几何更新：跟随选中框外扩——
        /// 四边按手柄命中半径（14）外扩，顶部为旋转钮多留（钮心在框顶上方 22 + 命中半径 14 = 36）。
        /// 之外的区域不可命中 → 输入直接落到 inkCanvas（选中状态下也能直接书写）。
        /// 由 updateBorderStrokeSelectionControlLocation 搭车调用（选区变化/拖动/旋转/缩放后都会走它）。
        /// </summary>
        private void UpdateSelectionHitArea()
        {
            try
            {
                var b = inkCanvas.GetSelectionBounds();
                if (b.IsEmpty || b.Width <= 0 || b.Height <= 0)
                {
                    BorderSelectionHitArea.Width = 0;
                    BorderSelectionHitArea.Height = 0;
                    return;
                }
                //inkCanvas 坐标系 → 覆盖层 Grid 坐标系（覆盖层无背景可命中，
                //Border 用 Margin+尺寸绝对定位，所以必须自己换算坐标）
                var t = inkCanvas.TransformToVisual(GridInkCanvasSelectionCover);
                var tl = t.Transform(b.TopLeft);

                const double side = SelectionHandleHitRadius;      //四边：手柄命中半径
                const double top = RotateHandleOffset + SelectionHandleHitRadius; //顶边：再加旋转钮伸出距离

                BorderSelectionHitArea.Margin = new Thickness(tl.X - side, tl.Y - top, 0, 0);
                BorderSelectionHitArea.Width = b.Width + side * 2;
                BorderSelectionHitArea.Height = b.Height + top + side;
            }
            catch { }
        }

        /// <summary>
        /// 框外落笔时取消选中（配合覆盖层只覆盖选区附近的新结构）：
        /// 选中不再拦截画布输入，笔/鼠标落到 inkCanvas 的瞬间调用本方法，
        /// 清掉选中让这一笔直接写出（含一次性选中残留的 Select 模式恢复为笔模式）。
        /// 用户主动用"选择工具"选中的场景（_isOneShotGraphSelection=false）保持 Select 模式——
        /// 用户接下来通常是继续框选新内容。
        /// </summary>
        private void DeselectStrokesForCanvasInput()
        {
            try
            {
                if (inkCanvas.GetSelectedStrokes().Count == 0) return;

                //先读一次性选中标志：true = 选中来自图形插入（模式被 Select() 偷偷切成了 Select），
                //取消后要恢复笔模式，当前这一笔才画得出来
                bool restoreInk = _isOneShotGraphSelection;
                bool eraserShapeBefore = forcePointEraser;

                //屏蔽 SelectionChanged 的快照副作用（取消不需要快照）
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection());
                isProgramChangeStrokeSelection = false;

                if (restoreInk)
                {
                    inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                    forcePointEraser = eraserShapeBefore;
                }

                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                TryEndOneShotSelection(); //一次性选中收尾（内部幂等，延迟恢复笔模式与本处不冲突）
            }
            catch { }
        }

        private void updateBorderStrokeSelectionControlLocation()
        {
            //按钮增减后实际宽度会变，优先用 ActualWidth（布局完成后 > 10），常量仅作初值兜底
            double controlWidth = BorderStrokeSelectionControl.ActualWidth > 10 ? BorderStrokeSelectionControl.ActualWidth : BorderStrokeSelectionControlWidth;
            double borderLeft = (inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Right - controlWidth) / 2;
            double borderTop = inkCanvas.GetSelectionBounds().Bottom + 15;
            if (borderLeft < 0) borderLeft = 0;
            if (borderTop < 0) borderTop = 0;
            if (Width - borderLeft < controlWidth || double.IsNaN(borderLeft)) borderLeft = Width - controlWidth;
            if (Height - borderTop < BorderStrokeSelectionControlHeight || double.IsNaN(borderTop)) borderTop = Height - BorderStrokeSelectionControlHeight;
            BorderStrokeSelectionControl.Margin = new Thickness(borderLeft, borderTop, 0, 0);

            //旋转钮跟随选中框顶部中央（与 HitTestSelectionHandle 的命中点同一位置：
            //框顶边中点向上伸出 RotateHandleOffset 像素处，再减去钮自身半径居中）
            try
            {
                var b = inkCanvas.GetSelectionBounds();
                GridRotateHandle.Margin = new Thickness(
                    b.Left + b.Width / 2 - GridRotateHandle.Width / 2,
                    Math.Max(0, b.Top - RotateHandleOffset - GridRotateHandle.Height / 2),
                    0, 0);
            }
            catch { }

            //选区可命中区域同步跟随（覆盖层新结构：只覆盖选中框附近，框外输入直达画布）
            UpdateSelectionHitArea();
        }

        private void GridInkCanvasSelectionCover_ManipulationStarting(object sender, ManipulationStartingEventArgs e)
        {
            e.Mode = ManipulationModes.All;
        }

        private void GridInkCanvasSelectionCover_ManipulationCompleted(object sender, ManipulationCompletedEventArgs e)
        {
            if (StrokeManipulationHistory?.Count > 0)
            {
                timeMachine.CommitStrokeManipulationHistory(StrokeManipulationHistory);
                foreach (var item in StrokeManipulationHistory)
                {
                    StrokeInitialHistory[item.Key] = item.Value.Item2;
                }
                StrokeManipulationHistory = null;
            }
            if (DrawingAttributesHistory.Count > 0)
            {
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
        }

        private void GridInkCanvasSelectionCover_ManipulationDelta(object sender, ManipulationDeltaEventArgs e)
        {
            try
            {
                if (dec.Count >= 1)
                {
                    ManipulationDelta md = e.DeltaManipulation;
                    Vector trans = md.Translation;  // 获得位移矢量
                    double rotate = md.Rotation;  // 获得旋转角度
                    Vector scale = md.Scale;  // 获得缩放倍数

                    Matrix m = new Matrix();

                    // Find center of element and then transform to get current location of center
                    FrameworkElement fe = e.Source as FrameworkElement;
                    Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
                    center = new Point(inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Width / 2,
                        inkCanvas.GetSelectionBounds().Top + inkCanvas.GetSelectionBounds().Height / 2);
                    center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

                    // Update matrix to reflect translation/rotation
                    m.Translate(trans.X, trans.Y);  // 移动
                    m.ScaleAt(scale.X, scale.Y, center.X, center.Y);  // 缩放

                    StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
                    if (StrokesSelectionClone.Count != 0)
                    {
                        strokes = StrokesSelectionClone;
                    }
                    else if (Settings.Gesture.IsEnableTwoFingerRotationOnSelection)
                    {
                        m.RotateAt(rotate, center.X, center.Y);  // 旋转
                    }
                    foreach (Stroke stroke in strokes)
                    {
                        stroke.Transform(m, false);

                        try
                        {
                            stroke.DrawingAttributes.Width *= md.Scale.X;
                            stroke.DrawingAttributes.Height *= md.Scale.Y;
                        }
                        catch { }
                    }
                    updateBorderStrokeSelectionControlLocation();
                }
            }
            catch { }
        }

        private void GridInkCanvasSelectionCover_TouchDown(object sender, TouchEventArgs e)
        {
        }

        private void GridInkCanvasSelectionCover_TouchUp(object sender, TouchEventArgs e)
        {
        }

        Point lastTouchPointOnGridInkCanvasCover = new Point(0, 0);
        private void GridInkCanvasSelectionCover_PreviewTouchDown(object sender, TouchEventArgs e)
        {
            dec.Add(e.TouchDevice.Id);
            //设备1个的时候，记录中心点
            if (dec.Count == 1)
            {
                TouchPoint touchPoint = e.GetTouchPoint(null);
                centerPoint = touchPoint.Position;
                lastTouchPointOnGridInkCanvasCover = touchPoint.Position;

                //克隆已改为按钮点击即生成（见 BorderStrokeSelectionClone_MouseUp），触摸路径不再消费开关
            }
        }

        private void GridInkCanvasSelectionCover_PreviewTouchUp(object sender, TouchEventArgs e)
        {
            dec.Remove(e.TouchDevice.Id);
            if (dec.Count >= 1) return;
            isProgramChangeStrokeSelection = false;
            if (lastTouchPointOnGridInkCanvasCover == e.GetTouchPoint(null).Position)
            {
                if (lastTouchPointOnGridInkCanvasCover.X < inkCanvas.GetSelectionBounds().Left ||
                    lastTouchPointOnGridInkCanvasCover.Y < inkCanvas.GetSelectionBounds().Top ||
                    lastTouchPointOnGridInkCanvasCover.X > inkCanvas.GetSelectionBounds().Right ||
                    lastTouchPointOnGridInkCanvasCover.Y > inkCanvas.GetSelectionBounds().Bottom)
                {
                    inkCanvas.Select(new StrokeCollection());
                    StrokesSelectionClone = new StrokeCollection();
                    //与鼠标路径（MouseUp 分支）保持一致：取消选中的同时收起选区遮罩，
                    //否则触摸点掉选中后拖动控制条仍残留悬空。
                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                    //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                    TryEndOneShotSelection();
                }
            }
            else if (inkCanvas.GetSelectedStrokes().Count == 0)
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                StrokesSelectionClone = new StrokeCollection();
            }
            else
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                StrokesSelectionClone = new StrokeCollection();
            }
        }

        #endregion Selection Gestures
    }
}
