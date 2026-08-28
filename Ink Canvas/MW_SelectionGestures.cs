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
            m.RotateAt(45, center.X, center.Y);  // 旋转

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

            //克隆已改为按钮点击即生成（见 BorderStrokeSelectionClone_MouseUp），此处仅负责拖动

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

            if (isMouseSelectionDragging && hasMouseSelectionDragMoved)
            {
                //真实拖动结束：保持选区可见，提交撤销历史
                FinishMouseSelectionDrag();
            }
            else
            {
                //单击（未拖动）：结束拖动状态并取消选区（恢复原行为）
                if (isMouseSelectionDragging) FinishMouseSelectionDrag();
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection());
                isProgramChangeStrokeSelection = false;
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>结束鼠标拖动：释放捕获、清空克隆引用、提交撤销历史（同触摸路径 ManipulationCompleted）</summary>
        private void FinishMouseSelectionDrag()
        {
            isMouseSelectionDragging = false;
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
        }

        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 0;
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
            }
            else
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                updateBorderStrokeSelectionControlLocation();
            }
        }

        private void updateBorderStrokeSelectionControlLocation()
        {
            double borderLeft = (inkCanvas.GetSelectionBounds().Left + inkCanvas.GetSelectionBounds().Right - BorderStrokeSelectionControlWidth) / 2;
            double borderTop = inkCanvas.GetSelectionBounds().Bottom + 15;
            if (borderLeft < 0) borderLeft = 0;
            if (borderTop < 0) borderTop = 0;
            if (Width - borderLeft < BorderStrokeSelectionControlWidth || double.IsNaN(borderLeft)) borderLeft = Width - BorderStrokeSelectionControlWidth;
            if (Height - borderTop < BorderStrokeSelectionControlHeight || double.IsNaN(borderTop)) borderTop = Height - BorderStrokeSelectionControlHeight;
            BorderStrokeSelectionControl.Margin = new Thickness(borderLeft, borderTop, 0, 0);
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
