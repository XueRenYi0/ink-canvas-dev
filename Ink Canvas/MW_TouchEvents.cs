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
    /// <summary>MainWindow 分部类：触摸事件（含多指）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Touch Events

        #region Multi-Touch

        bool isInMultiTouchMode = false;
        private void BorderMultiTouchMode_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (isInMultiTouchMode)
            {
                inkCanvas.StylusDown -= MainWindow_StylusDown;
                inkCanvas.StylusMove -= MainWindow_StylusMove;
                inkCanvas.StylusUp -= MainWindow_StylusUp;
                inkCanvas.TouchDown -= MainWindow_TouchDown;
                inkCanvas.TouchDown += Main_Grid_TouchDown;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                inkCanvas.Children.Clear();
                isInMultiTouchMode = false;
                //切回单人书写：显示单人像（自绘 Path，与工具栏描边风格统一）
                PathPersonSingle.Visibility = Visibility.Visible;
                PathPeopleMulti.Visibility = Visibility.Collapsed;
            }
            else
            {
                inkCanvas.StylusDown += MainWindow_StylusDown;
                inkCanvas.StylusMove += MainWindow_StylusMove;
                inkCanvas.StylusUp += MainWindow_StylusUp;
                inkCanvas.TouchDown -= Main_Grid_TouchDown;
                inkCanvas.TouchDown += MainWindow_TouchDown;
                //多人模式：保持 Ink 模式——鼠标/手写板仍走原生墨迹通道；
                //Stylus/Touch 自定义多笔路径通过 e.Handled 防止双写，切后鼠标无需再点颜色就能写。
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                inkCanvas.Children.Clear();
                isInMultiTouchMode = true;
                //开启多指书写：显示双人像（左前大人 + 右后小孩，错位布局）
                PathPersonSingle.Visibility = Visibility.Collapsed;
                PathPeopleMulti.Visibility = Visibility.Visible;
            }
        }

        private void MainWindow_TouchDown(object sender, TouchEventArgs e)
        {
            double boundWidth = e.GetTouchPoint(null).Bounds.Width;
            if (boundWidth > 20)
            {
                //粗触摸 → 矩形橡皮（与单人模式/工具栏同款黄金比例竖矩形，用户偏好矩形）
                inkCanvas.EraserShape = new RectangleStylusShape(boundWidth, boundWidth * 1.618);
                TouchDownPointsList[e.TouchDevice.Id] = InkCanvasEditingMode.EraseByPoint;
                inkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
            }
            else
            {
                TouchDownPointsList[e.TouchDevice.Id] = InkCanvasEditingMode.None;
                //注意：Touch 不走原生 Ink 通道（保留 EditingMode=Ink 给鼠标用），触摸多笔靠自定义 Stylus 通道合成
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
            e.Handled = true;
        }

        private void MainWindow_StylusDown(object sender, StylusDownEventArgs e)
        {
            TouchDownPointsList[e.StylusDevice.Id] = InkCanvasEditingMode.None;
            e.Handled = true; //阻止 Stylus 再进入 Ink 原生通道，避免与多笔自定义路径双写
        }

        private void MainWindow_StylusUp(object sender, StylusEventArgs e)
        {
            try
            {
                inkCanvas.Strokes.Add(GetStrokeVisual(e.StylusDevice.Id).Stroke);
                inkCanvas.Children.Remove(GetVisualCanvas(e.StylusDevice.Id));

                inkCanvas_StrokeCollected(inkCanvas, new InkCanvasStrokeCollectedEventArgs(GetStrokeVisual(e.StylusDevice.Id).Stroke));
            }
            catch (Exception ex)
            {
                Label.Content = ex.ToString();
            }
            try
            {
                StrokeVisualList.Remove(e.StylusDevice.Id);
                VisualCanvasList.Remove(e.StylusDevice.Id);
                TouchDownPointsList.Remove(e.StylusDevice.Id);
                if (StrokeVisualList.Count == 0 || VisualCanvasList.Count == 0 || TouchDownPointsList.Count == 0)
                {
                    inkCanvas.Children.Clear();
                    StrokeVisualList.Clear();
                    VisualCanvasList.Clear();
                    TouchDownPointsList.Clear();
                }
            }
            catch { }
            e.Handled = true;
        }

        private void MainWindow_StylusMove(object sender, StylusEventArgs e)
        {
            try
            {
                if (GetTouchDownPointsList(e.StylusDevice.Id) != InkCanvasEditingMode.None) return;
                try
                {
                    if (e.StylusDevice.StylusButtons[1].StylusButtonState == StylusButtonState.Down) return;
                }
                catch { }
                var strokeVisual = GetStrokeVisual(e.StylusDevice.Id);
                var stylusPointCollection = e.GetStylusPoints(this);
                foreach (var stylusPoint in stylusPointCollection)
                {
                    strokeVisual.Add(new StylusPoint(stylusPoint.X, stylusPoint.Y, stylusPoint.PressureFactor));
                }

                strokeVisual.Redraw();
            }
            catch { }
            e.Handled = true;
        }

        private StrokeVisual GetStrokeVisual(int id)
        {
            if (StrokeVisualList.TryGetValue(id, out var visual))
            {
                return visual;
            }

            var strokeVisual = new StrokeVisual(inkCanvas.DefaultDrawingAttributes.Clone());
            StrokeVisualList[id] = strokeVisual;
            StrokeVisualList[id] = strokeVisual;
            var visualCanvas = new VisualCanvas(strokeVisual);
            VisualCanvasList[id] = visualCanvas;
            inkCanvas.Children.Add(visualCanvas);

            return strokeVisual;
        }

        private VisualCanvas GetVisualCanvas(int id)
        {
            if (VisualCanvasList.TryGetValue(id, out var visualCanvas))
            {
                return visualCanvas;
            }
            return null;
        }

        private InkCanvasEditingMode GetTouchDownPointsList(int id)
        {
            if (TouchDownPointsList.TryGetValue(id, out var inkCanvasEditingMode))
            {
                return inkCanvasEditingMode;
            }
            return inkCanvas.EditingMode;
        }

        private Dictionary<int, InkCanvasEditingMode> TouchDownPointsList { get; } = new Dictionary<int, InkCanvasEditingMode>();
        private Dictionary<int, StrokeVisual> StrokeVisualList { get; } = new Dictionary<int, StrokeVisual>();
        private Dictionary<int, VisualCanvas> VisualCanvasList { get; } = new Dictionary<int, VisualCanvas>();

        #endregion

        int lastTouchDownTime = 0, lastTouchUpTime = 0;

        Point iniP = new Point(0, 0);
        bool isLastTouchEraser = false;
        private bool forcePointEraser = true;
        private bool _lockSmith = false; //临时停用双指手势

        // ===== 任务4：触摸分级防误触（大面积接触时长过滤）=====
        // 大面积触摸（手掌/手背）按下时刻（TickCount 毫秒），用于判定"极快抬起 = 衣袖扫过类误触"
        private int _largeTouchDownTickCount = 0;
        // 当前触摸序列中是否出现过大面积接触（抬起时才做时长过滤判定，避免普通手指书写受影响）
        private bool _largeTouchActive = false;
        // 大面积接触误触判定阈值：接触后 120ms 内抬起视为无意扫过（有意手掌擦除通常持续更久）
        private const int LargeTouchAccidentalMs = 120;

        // ===== 任务5：双指手势轴锁定（垂直滚动 / 水平平移分流）=====
        // 轴锁定状态：0=未锁定（意图不明，累计观察中） 1=垂直（滚动笔记） 2=其他（平移/缩放/旋转走原逻辑）
        private int _twoFingerAxisLock = 0;
        // 双指手势累计位移（用于判定主方向）
        private Vector _twoFingerTotalTranslation = new Vector();
        // 双指手势累计缩放（乘积；两指不平行滑动会产生缩放噪声，需要门限过滤）
        private double _twoFingerTotalScale = 1.0;
        // 双指手势累计旋转角度（度；用于区分旋转手势与垂直滚动）
        private double _twoFingerTotalRotation = 0;

        /// <summary>
        /// 重置双指手势轴锁定状态（手势完全结束时调用：
        /// inkCanvas_PreviewTouchUp 全部手指抬起 / ManipulationCompleted 双保险）
        /// </summary>
        private void ResetTwoFingerAxisLock()
        {
            _twoFingerAxisLock = 0;
            _twoFingerTotalTranslation = new Vector();
            _twoFingerTotalScale = 1.0;
            _twoFingerTotalRotation = 0;
        }

        private void Main_Grid_TouchDown(object sender, TouchEventArgs e)
        {
            BorderClearInDelete.Visibility = Visibility.Collapsed;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }

            if (NeedUpdateIniP())
            {
                iniP = e.GetTouchPoint(inkCanvas).Position;
            }
            if (drawingShapeMode == 9 && isFirstTouchCuboid == false)
            {
                MouseTouchMove(iniP);
            }
            inkCanvas.Opacity = 1;
            double boundsWidth = GetTouchBoundWidth(e);
            var eraserMultiplier = 1d;
            if (!Settings.Advanced.EraserBindTouchMultiplier && Settings.Advanced.IsSpecialScreen) eraserMultiplier = 1 / Settings.Advanced.TouchMultiplier;
            if (boundsWidth > BoundsWidth)
            {
                isLastTouchEraser = true;
                if (drawingShapeMode == 0 && forceEraser) return;
                if (boundsWidth > BoundsWidth * 2.5)
                {
                    //任务4：记录大面积接触按下时刻——若 120ms 内抬起且墨迹被擦，
                    //判定为衣袖/手掌快速扫过类误触，抬起时撤销误擦（恢复墨迹快照）
                    _largeTouchDownTickCount = Environment.TickCount;
                    _largeTouchActive = true;

                    double k = 1;
                    switch (Settings.Canvas.EraserSize)
                    {
                        case 0:
                            k = 0.5;
                            break;
                        case 1:
                            k = 0.8;
                            break;
                        case 3:
                            k = 1.25;
                            break;
                        case 4:
                            k = 1.8;
                            break;
                    }
                    //大面积（手掌/手背）→ 大号矩形橡皮：与工具栏橡皮同款黄金比例竖矩形（用户偏好矩形）。
                    //宽度随接触面积自适应（手掌大→橡皮大），并应用橡皮档位系数 k 与触摸倍率
                    double w = boundsWidth * 1.5 * k * eraserMultiplier;
                    inkCanvas.EraserShape = new RectangleStylusShape(w, w * 1.618);
                    inkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
                }
                else
                {
                    if (StackPanelPPTControls.Visibility == Visibility.Visible && inkCanvas.Strokes.Count == 0 && Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl)
                    {
                        isLastTouchEraser = false;
                        inkCanvas.EditingMode = InkCanvasEditingMode.GestureOnly;
                        inkCanvas.Opacity = 0.1;
                    }
                    else
                    {
                        //inkCanvas.EraserShape = new RectangleStylusShape(8, 8); //old old
                        //inkCanvas.EraserShape = forcePointEraser ? new EllipseStylusShape(50, 50) : new EllipseStylusShape(5, 5); //last
                        //inkCanvas.EraserShape = new EllipseStylusShape(boundsWidth * 1.5, boundsWidth * 1.5); //old old
                        //inkCanvas.EditingMode = forcePointEraser ? InkCanvasEditingMode.EraseByPoint : InkCanvasEditingMode.EraseByStroke; //last
                        //指背等中等接触 → 按笔画擦：形状统一走 CreateEraserShape（笔画擦无轮廓，形状无视觉影响）
                        inkCanvas.EraserShape = CreateEraserShape(false);
                        inkCanvas.EditingMode = InkCanvasEditingMode.EraseByStroke;
                    }
                }
            }
            else
            {
                isLastTouchEraser = false;
                // 触摸事件橡皮形状：黄金比例竖矩形（与按钮逻辑共用同一方法）
                inkCanvas.EraserShape = CreateEraserShape(forcePointEraser);
                if (forceEraser) return;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
        }

        public double GetTouchBoundWidth(TouchEventArgs e)
        {
            var args = e.GetTouchPoint(null).Bounds;
            double value;
            if (!Settings.Advanced.IsQuadIR) value = args.Width;
            else value = Math.Sqrt(args.Width * args.Height); //四边红外
            if (Settings.Advanced.IsSpecialScreen) value *= Settings.Advanced.TouchMultiplier;
            return value;
        }

        //记录触摸设备ID
        private List<int> dec = new List<int>();
        //中心点
        System.Windows.Point centerPoint;
        InkCanvasEditingMode lastInkCanvasEditingMode = InkCanvasEditingMode.Ink;
        bool isSingleFingerDragMode = false;

        //防止衣服误触造成的墨迹消失

        private void inkCanvas_PreviewTouchDown(object sender, TouchEventArgs e)
        {
            dec.Add(e.TouchDevice.Id);
            //设备1个的时候，记录中心点
            if (dec.Count == 1)
            {
                TouchPoint touchPoint = e.GetTouchPoint(inkCanvas);
                centerPoint = touchPoint.Position;

                //记录第一根手指点击时的 StrokeCollection
                lastTouchDownStrokeCollection = inkCanvas.Strokes.Clone();
            }
            //设备两个及两个以上，将画笔功能关闭
            if (dec.Count > 1 || isSingleFingerDragMode || !Settings.Gesture.IsEnableTwoFingerGesture)
            {
                if (isInMultiTouchMode || !Settings.Gesture.IsEnableTwoFingerGesture) return;
                if (inkCanvas.EditingMode != InkCanvasEditingMode.None && inkCanvas.EditingMode != InkCanvasEditingMode.Select)
                {
                    lastInkCanvasEditingMode = inkCanvas.EditingMode;
                    inkCanvas.EditingMode = InkCanvasEditingMode.None;
                }
            }
        }

        private void inkCanvas_PreviewTouchUp(object sender, TouchEventArgs e)
        {
            //手势完成后切回之前的状态
            if (dec.Count > 1)
            {
                if (inkCanvas.EditingMode == InkCanvasEditingMode.None)
                {
                    inkCanvas.EditingMode = lastInkCanvasEditingMode;
                }
            }
            dec.Remove(e.TouchDevice.Id);
            inkCanvas.Opacity = 1;
            if (dec.Count == 0)
            {
                //任务5：全部手指抬起 = 双指手势结束，重置轴锁定状态（下次手势重新判定意图）
                ResetTwoFingerAxisLock();

                //任务4：大面积接触时长过滤——手掌/手背级接触在 120ms 内抬起且墨迹被擦，
                //判定为衣袖扫过类误触，恢复按下时的墨迹快照直接撤销误擦（显示与数据同步还原）。
                //有意的手掌擦除通常持续按压超过 120ms，不受影响。
                if (_largeTouchActive)
                {
                    _largeTouchActive = false;
                    if (Environment.TickCount - _largeTouchDownTickCount < LargeTouchAccidentalMs &&
                        lastTouchDownStrokeCollection.Count != inkCanvas.Strokes.Count)
                    {
                        //Clone 恢复：画布与备份各持一份，避免后续单方修改互相牵连
                        inkCanvas.Strokes = lastTouchDownStrokeCollection.Clone();
                    }
                }

                if (lastTouchDownStrokeCollection.Count() != inkCanvas.Strokes.Count() &&
                    !(drawingShapeMode == 9 && !isFirstTouchCuboid))
                {
                    int whiteboardIndex = CurrentWhiteboardIndex;
                    if (currentMode == 0)
                    {
                        whiteboardIndex = 0;
                    }
                    strokeCollections[whiteboardIndex] = lastTouchDownStrokeCollection;
                }
            }
        }
        private void inkCanvas_ManipulationStarting(object sender, ManipulationStartingEventArgs e)
        {
            e.Mode = ManipulationModes.All;
        }

        private void inkCanvas_ManipulationInertiaStarting(object sender, ManipulationInertiaStartingEventArgs e)
        {

        }

        private void Main_Grid_ManipulationCompleted(object sender, ManipulationCompletedEventArgs e)
        {
            if (e.Manipulators.Count() == 0)
            {
                //任务5：手势完全结束（所有触点抬起），重置轴锁定（与 PreviewTouchUp 双保险，
                //覆盖个别设备 TouchUp 事件时序异常导致未重置的情况）
                ResetTwoFingerAxisLock();
                if (forceEraser) return;
                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
        }

        private void Main_Grid_ManipulationDelta(object sender, ManipulationDeltaEventArgs e)
        {
            if (isInMultiTouchMode || !Settings.Gesture.IsEnableTwoFingerGesture || _lockSmith) return;
            if ((dec.Count >= 2 && (Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode || StackPanelPPTControls.Visibility != Visibility.Visible || StackPanelPPTButtons.Visibility == Visibility.Collapsed)) || isSingleFingerDragMode)
            {
                ManipulationDelta md = e.DeltaManipulation;
                Vector trans = md.Translation;  // 获得位移矢量
                double rotate = md.Rotation;  // 获得旋转角度
                Vector scale = md.Scale;  // 获得缩放倍数

                // ===== 任务5：双指手势轴锁定与分流 =====
                // 目标：双指垂直滑动 → 笔记滚动（与滚轮/滚动胶囊一致，符合触摸屏"内容跟手"惯例）；
                //       水平滑动/捏合缩放/旋转 → 保持原有平移缩放逻辑。
                // 做法：手势初期先累计位移/缩放/旋转观察意图，主方向明确（超过锁定阈值）后
                //       锁定该轴直到手势结束，防止斜向拖动时"既滚又移"互相打架。
                if (!isSingleFingerDragMode && IsNoteScrollActive && Settings.Gesture.IsEnableTwoFingerTranslate)
                {
                    _twoFingerTotalTranslation += trans;   // 累计位移（含判定期内的所有增量）
                    _twoFingerTotalScale *= scale.X;       // 累计缩放（乘积）
                    _twoFingerTotalRotation += rotate;    // 累计旋转角

                    if (_twoFingerAxisLock == 0)
                    {
                        // 缩放噪声门限：两指不严格平行滑动时会持续上报约 1.0x 附近的微小缩放，
                        // 累计偏离超过 8% 才认定为有意的捏合手势（走原缩放逻辑）
                        bool isPinch = Math.Abs(_twoFingerTotalScale - 1.0) > 0.08;
                        // 旋转判定：累计转角超过 5° 视为有意旋转（走原旋转逻辑）
                        bool isRotate = Math.Abs(_twoFingerTotalRotation) > 5.0;

                        if (isPinch || isRotate)
                        {
                            _twoFingerAxisLock = 2; // 捏合/旋转 → 锁定为原手势通道
                        }
                        else
                        {
                            // 轴判定：主轴位移超过 15px 且明显大于副轴（1.5 倍）才锁定方向，
                            // 判定期内的微小位移不应用（避免手势开始时产生意外平移）
                            double ax = Math.Abs(_twoFingerTotalTranslation.X);
                            double ay = Math.Abs(_twoFingerTotalTranslation.Y);
                            if (ay > 15 && ay > ax * 1.5)
                                _twoFingerAxisLock = 1; // 垂直 → 滚动笔记
                            else if (ax > 15 && ax > ay * 1.5)
                                _twoFingerAxisLock = 2; // 水平 → 原平移逻辑
                            else
                                return; // 意图尚不明确：本次增量不应用，继续观察
                        }
                    }

                    if (_twoFingerAxisLock == 1)
                    {
                        // 垂直滑动 → ScrollNote 滚动（内容跟随手指）：
                        // 手指上滑（trans.Y < 0）→ delta > 0 → 向下滚动，历史墨迹上移，露出下方空白；
                        // 手指下滑（trans.Y > 0）→ delta < 0 → 回看上方历史（顶部夹取为 0 由 ScrollNote 内部处理）。
                        // 滚动即时同步滚动胶囊指示；锁定后本 delta 不再叠加平移/缩放。
                        ScrollNote(-trans.Y);
                        return;
                    }
                    // 锁定为 2（水平平移/捏合/旋转）或单指拖动模式：落入下方原有手势逻辑
                }
                // ===== 任务5 结束 =====

                Matrix m = new Matrix();

                // Find center of element and then transform to get current location of center
                FrameworkElement fe = e.Source as FrameworkElement;
                Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
                center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

                // Update matrix to reflect translation/rotation
                if (Settings.Gesture.IsEnableTwoFingerTranslate)
                    m.Translate(trans.X, trans.Y);  // 移动
                if (Settings.Gesture.IsEnableTwoFingerRotation)
                    m.RotateAt(rotate, center.X, center.Y);  // 旋转
                if (Settings.Gesture.IsEnableTwoFingerZoom)
                    m.ScaleAt(scale.X, scale.Y, center.X, center.Y);  // 缩放

                StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
                if (strokes.Count != 0)
                {
                    foreach (Stroke stroke in strokes)
                    {
                        stroke.Transform(m, false);

                        foreach (Circle circle in circles)
                        {
                            if (stroke == circle.Stroke)
                            {
                                circle.R = GetDistance(circle.Stroke.StylusPoints[0].ToPoint(), circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].ToPoint()) / 2;
                                circle.Centroid = new Point((circle.Stroke.StylusPoints[0].X + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].X) / 2,
                                                            (circle.Stroke.StylusPoints[0].Y + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].Y) / 2);
                                break;
                            }
                        }

                        if (Settings.Gesture.IsEnableTwoFingerZoom)
                        {
                            try
                            {
                                stroke.DrawingAttributes.Width *= md.Scale.X;
                                stroke.DrawingAttributes.Height *= md.Scale.Y;
                            }
                            catch { }
                        }
                    }
                }
                else
                {
                    foreach (Stroke stroke in inkCanvas.Strokes)
                    {
                        stroke.Transform(m, false);

                        if (Settings.Gesture.IsEnableTwoFingerZoom)
                        {
                            try
                            {
                                stroke.DrawingAttributes.Width *= md.Scale.X;
                                stroke.DrawingAttributes.Height *= md.Scale.Y;
                            }
                            catch { }
                        }
                    }
                    foreach (Circle circle in circles)
                    {
                        circle.R = GetDistance(circle.Stroke.StylusPoints[0].ToPoint(), circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].ToPoint()) / 2;
                        circle.Centroid = new Point((circle.Stroke.StylusPoints[0].X + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].X) / 2,
                                                    (circle.Stroke.StylusPoints[0].Y + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].Y) / 2);
                    }
                }
            }
        }

        #endregion Touch Events
    }
}
