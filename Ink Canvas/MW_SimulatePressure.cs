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
    /// <summary>MainWindow 分部类：压感模拟与墨迹转图形（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Simulate Pen Pressure & Ink To Shape

        StrokeCollection newStrokes = new StrokeCollection();
        List<Circle> circles = new List<Circle>();

        //此函数中的所有代码版权所有 WXRIW，在其他项目中使用前必须提前联系（wxriw@outlook.com），谢谢！
        private void inkCanvas_StrokeCollected(object sender, InkCanvasStrokeCollectedEventArgs e)
        {
            try
            {
                inkCanvas.Opacity = 1;
                if (Settings.InkToShape.IsInkToShapeEnabled && !Environment.Is64BitProcess)
                {
                    void InkToShapeProcess()
                    {
                        try
                        {
                            newStrokes.Add(e.Stroke);
                            if (newStrokes.Count > 4) newStrokes.RemoveAt(0);
                            for (int i = 0; i < newStrokes.Count; i++)
                            {
                                if (!inkCanvas.Strokes.Contains(newStrokes[i])) newStrokes.RemoveAt(i--);
                            }
                            for (int i = 0; i < circles.Count; i++)
                            {
                                if (!inkCanvas.Strokes.Contains(circles[i].Stroke)) circles.RemoveAt(i);
                            }
                            var strokeReco = new StrokeCollection();
                            var result = InkRecognizeHelper.RecognizeShape(newStrokes);
                            for (int i = newStrokes.Count - 1; i >= 0; i--)
                            {
                                strokeReco.Add(newStrokes[i]);
                                var newResult = InkRecognizeHelper.RecognizeShape(strokeReco);
                                if (newResult.InkDrawingNode.GetShapeName() == "Circle" || newResult.InkDrawingNode.GetShapeName() == "Ellipse")
                                {
                                    result = newResult;
                                    break;
                                }
                                //Label.Visibility = Visibility.Visible;
                                Label.Content = circles.Count.ToString() + "\n" + newResult.InkDrawingNode.GetShapeName();
                            }
                            if (result.InkDrawingNode.GetShapeName() == "Circle")
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                if (shape.Width > 75)
                                {
                                    foreach (Circle circle in circles)
                                    {
                                        //判断是否画同心圆
                                        if (Math.Abs(result.Centroid.X - circle.Centroid.X) / shape.Width < 0.12 &&
                                            Math.Abs(result.Centroid.Y - circle.Centroid.Y) / shape.Width < 0.12)
                                        {
                                            result.Centroid = circle.Centroid;
                                            break;
                                        }
                                        else
                                        {
                                            double d = (result.Centroid.X - circle.Centroid.X) * (result.Centroid.X - circle.Centroid.X) +
                                               (result.Centroid.Y - circle.Centroid.Y) * (result.Centroid.Y - circle.Centroid.Y);
                                            d = Math.Sqrt(d);
                                            //判断是否画外切圆
                                            double x = shape.Width / 2.0 + circle.R - d;
                                            if (Math.Abs(x) / shape.Width < 0.1)
                                            {
                                                double sinTheta = (result.Centroid.Y - circle.Centroid.Y) / d;
                                                double cosTheta = (result.Centroid.X - circle.Centroid.X) / d;
                                                double newX = result.Centroid.X + x * cosTheta;
                                                double newY = result.Centroid.Y + x * sinTheta;
                                                result.Centroid = new Point(newX, newY);
                                            }
                                            //判断是否画外切圆
                                            x = Math.Abs(circle.R - shape.Width / 2.0) - d;
                                            if (Math.Abs(x) / shape.Width < 0.1)
                                            {
                                                double sinTheta = (result.Centroid.Y - circle.Centroid.Y) / d;
                                                double cosTheta = (result.Centroid.X - circle.Centroid.X) / d;
                                                double newX = result.Centroid.X + x * cosTheta;
                                                double newY = result.Centroid.Y + x * sinTheta;
                                                result.Centroid = new Point(newX, newY);
                                            }
                                        }
                                    }

                                    Point iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                    Point endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);
                                    var pointList = GenerateEllipseGeometry(iniP, endP);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(point)
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    circles.Add(new Circle(result.Centroid, shape.Width / 2.0, stroke));
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Ellipse"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                //var shape1 = result.InkDrawingNode.GetShape();
                                //shape1.Fill = Brushes.Gray;
                                //Canvas.Children.Add(shape1);
                                var p = result.InkDrawingNode.HotPoints;
                                double a = GetDistance(p[0], p[2]) / 2; //长半轴
                                double b = GetDistance(p[1], p[3]) / 2; //短半轴
                                if (a < b)
                                {
                                    double t = a;
                                    a = b;
                                    b = t;
                                }

                                result.Centroid = new Point((p[0].X + p[2].X) / 2, (p[0].Y + p[2].Y) / 2);
                                bool needRotation = true;

                                if (shape.Width > 75 || shape.Height > 75 && p.Count == 4)
                                {
                                    Point iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                    Point endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);

                                    foreach (Circle circle in circles)
                                    {
                                        //判断是否画同心椭圆
                                        if (Math.Abs(result.Centroid.X - circle.Centroid.X) / a < 0.2 &&
                                            Math.Abs(result.Centroid.Y - circle.Centroid.Y) / a < 0.2)
                                        {
                                            result.Centroid = circle.Centroid;
                                            iniP = new Point(result.Centroid.X - shape.Width / 2, result.Centroid.Y - shape.Height / 2);
                                            endP = new Point(result.Centroid.X + shape.Width / 2, result.Centroid.Y + shape.Height / 2);

                                            //再判断是否与圆相切
                                            if (Math.Abs(a - circle.R) / a < 0.2)
                                            {
                                                if (shape.Width >= shape.Height)
                                                {
                                                    iniP.X = result.Centroid.X - circle.R;
                                                    endP.X = result.Centroid.X + circle.R;
                                                    iniP.Y = result.Centroid.Y - b;
                                                    endP.Y = result.Centroid.Y + b;
                                                }
                                                else
                                                {
                                                    iniP.Y = result.Centroid.Y - circle.R;
                                                    endP.Y = result.Centroid.Y + circle.R;
                                                    iniP.X = result.Centroid.X - a;
                                                    endP.X = result.Centroid.X + a;
                                                }
                                            }
                                            break;
                                        }
                                        else if (Math.Abs(result.Centroid.X - circle.Centroid.X) / a < 0.2)
                                        {
                                            double sinTheta = Math.Abs(circle.Centroid.Y - result.Centroid.Y) / circle.R;
                                            double cosTheta = Math.Sqrt(1 - sinTheta * sinTheta);
                                            double newA = circle.R * cosTheta;
                                            if (circle.R * sinTheta / circle.R < 0.9 && a / b > 2 && Math.Abs(newA - a) / newA < 0.3)
                                            {
                                                iniP.X = circle.Centroid.X - newA;
                                                endP.X = circle.Centroid.X + newA;
                                                iniP.Y = result.Centroid.Y - newA / 5;
                                                endP.Y = result.Centroid.Y + newA / 5;

                                                double topB = endP.Y - iniP.Y;

                                                SetNewBackupOfStroke();
                                                _currentCommitType = CommitReason.ShapeRecognition;
                                                inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                                newStrokes = new StrokeCollection();

                                                var _pointList = GenerateEllipseGeometry(iniP, endP, false, true);
                                                var _point = new StylusPointCollection(_pointList);
                                                var _stroke = new Stroke(_point)
                                                {
                                                    DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                                };
                                                var _dashedLineStroke = GenerateDashedLineEllipseStrokeCollection(iniP, endP, true, false);
                                                StrokeCollection strokes = new StrokeCollection()
                                                {
                                                    _stroke,
                                                    _dashedLineStroke
                                                };
                                                inkCanvas.Strokes.Add(strokes);
                                                _currentCommitType = CommitReason.UserInput;
                                                return;
                                            }
                                        }
                                        else if (Math.Abs(result.Centroid.Y - circle.Centroid.Y) / a < 0.2)
                                        {
                                            double cosTheta = Math.Abs(circle.Centroid.X - result.Centroid.X) / circle.R;
                                            double sinTheta = Math.Sqrt(1 - cosTheta * cosTheta);
                                            double newA = circle.R * sinTheta;
                                            if (circle.R * sinTheta / circle.R < 0.9 && a / b > 2 && Math.Abs(newA - a) / newA < 0.3)
                                            {
                                                iniP.X = result.Centroid.X - newA / 5;
                                                endP.X = result.Centroid.X + newA / 5;
                                                iniP.Y = circle.Centroid.Y - newA;
                                                endP.Y = circle.Centroid.Y + newA;
                                                needRotation = false;
                                            }
                                        }
                                    }

                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[2]);
                                    p[0] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[3]);
                                    p[1] = newPoints[0];
                                    p[3] = newPoints[1];

                                    var pointList = GenerateEllipseGeometry(iniP, endP);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(point)
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };

                                    if (needRotation)
                                    {
                                        Matrix m = new Matrix();
                                        FrameworkElement fe = e.Source as FrameworkElement;
                                        double tanTheta = (p[2].Y - p[0].Y) / (p[2].X - p[0].X);
                                        double theta = Math.Atan(tanTheta);
                                        m.RotateAt(theta * 180.0 / Math.PI, result.Centroid.X, result.Centroid.Y);
                                        stroke.Transform(m, false);
                                    }

                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Triangle"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                var p = result.InkDrawingNode.HotPoints;
                                if ((Math.Max(Math.Max(p[0].X, p[1].X), p[2].X) - Math.Min(Math.Min(p[0].X, p[1].X), p[2].X) >= 100 ||
                                    Math.Max(Math.Max(p[0].Y, p[1].Y), p[2].Y) - Math.Min(Math.Min(p[0].Y, p[1].Y), p[2].Y) >= 100) && result.InkDrawingNode.HotPoints.Count == 3)
                                {
                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[1]);
                                    p[0] = newPoints[0];
                                    p[1] = newPoints[1];
                                    newPoints = FixPointsDirection(p[0], p[2]);
                                    p[0] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[2]);
                                    p[1] = newPoints[0];
                                    p[2] = newPoints[1];

                                    var pointList = p.ToList();
                                    //pointList.Add(p[0]);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(GenerateFakePressureTriangle(point))
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                            else if (result.InkDrawingNode.GetShapeName().Contains("Rectangle") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Diamond") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Parallelogram") ||
                                     result.InkDrawingNode.GetShapeName().Contains("Square"))
                            {
                                var shape = result.InkDrawingNode.GetShape();
                                var p = result.InkDrawingNode.HotPoints;
                                if ((Math.Max(Math.Max(Math.Max(p[0].X, p[1].X), p[2].X), p[3].X) - Math.Min(Math.Min(Math.Min(p[0].X, p[1].X), p[2].X), p[3].X) >= 100 ||
                                    Math.Max(Math.Max(Math.Max(p[0].Y, p[1].Y), p[2].Y), p[3].Y) - Math.Min(Math.Min(Math.Min(p[0].Y, p[1].Y), p[2].Y), p[3].Y) >= 100) && result.InkDrawingNode.HotPoints.Count == 4)
                                {
                                    //纠正垂直与水平关系
                                    var newPoints = FixPointsDirection(p[0], p[1]);
                                    p[0] = newPoints[0];
                                    p[1] = newPoints[1];
                                    newPoints = FixPointsDirection(p[1], p[2]);
                                    p[1] = newPoints[0];
                                    p[2] = newPoints[1];
                                    newPoints = FixPointsDirection(p[2], p[3]);
                                    p[2] = newPoints[0];
                                    p[3] = newPoints[1];
                                    newPoints = FixPointsDirection(p[3], p[0]);
                                    p[3] = newPoints[0];
                                    p[0] = newPoints[1];

                                    var pointList = p.ToList();
                                    pointList.Add(p[0]);
                                    var point = new StylusPointCollection(pointList);
                                    var stroke = new Stroke(GenerateFakePressureRectangle(point))
                                    {
                                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                                    };
                                    SetNewBackupOfStroke();
                                    _currentCommitType = CommitReason.ShapeRecognition;
                                    inkCanvas.Strokes.Remove(result.InkDrawingNode.Strokes);
                                    inkCanvas.Strokes.Add(stroke);
                                    _currentCommitType = CommitReason.UserInput;
                                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                                    newStrokes = new StrokeCollection();
                                }
                            }
                        }
                        catch { }
                    }
                    //前后对比法拦截识别结果：识别过程若把"手写笔迹"替换成了"标准图形笔迹"
                    //（圆/椭圆/三角/矩形分支都是先 Remove 原笔迹再 Add 新笔迹），
                    //新增的那条就是要选中的图形。四个分支不用逐个改，这里统一接住。
                    List<Stroke> strokesBefore = new List<Stroke>();
                    foreach (Stroke s in inkCanvas.Strokes) strokesBefore.Add(s);
                    InkToShapeProcess();
                    StrokeCollection recognizedGraph = null;
                    foreach (Stroke s in inkCanvas.Strokes)
                    {
                        if (!strokesBefore.Contains(s))
                        {
                            if (recognizedGraph == null) recognizedGraph = new StrokeCollection();
                            recognizedGraph.Add(s);
                        }
                    }
                    if (recognizedGraph != null)
                    {
                        //手写识别成图 → 打标签 + 自动选中（统一入口，见 MW_GraphStrokes.cs）
                        InsertGraphStrokes(recognizedGraph);
                    }
                }

                // ============ 直线识别（ClassIn 式"停顿拉直"） ============
                // 笔按住不动超过阈值 → 当前笔迹拉直成预览线 → 按住不放可拖着端点
                // 以起笔点为轴心旋转/伸缩 → 抬笔定型（详见 region 内注释）。
                // 不依赖 InkAnalyzer（其 Line 判定过宽，且整条链路被 !Is64BitProcess 限制），
                // 纯几何 + 停顿手势，32/64 位都生效。
                if (_lineAssistArmed)
                {
                    try
                    {
                        if (_lineAssistCommitted)
                        {
                            //LineAssistEnd 已定型提交：这条是切 None 前的残影笔迹（部分 WPF 版本
                            //会在抬笔时补交），直接删掉防"两条线"复活
                            SetNewBackupOfStroke();
                            _currentCommitType = CommitReason.ShapeRecognition;
                            inkCanvas.Strokes.Remove(e.Stroke);
                            _currentCommitType = CommitReason.UserInput;
                            //删除动作会把 ReplacedStroke 悬空挂上这条残影（StrokesOnStrokesChanged
                            //的 ShapeRecognition 分支只赋值不提交），清掉防止污染下一次形状识别的历史配对
                            ReplacedStroke = null;
                        }
                        else
                        {
                            //兜底路径：采集仍产出了笔迹（未切 None 的场景），正常定型
                            CommitLineAssist(e.Stroke);
                        }
                    }
                    catch { }
                    finally { ResetLineAssist(); }
                }

                // 检查是否是压感笔书写
                foreach (StylusPoint stylusPoint in e.Stroke.StylusPoints)
                {
                    if (stylusPoint.PressureFactor != 0.5 && stylusPoint.PressureFactor != 0)
                    {
                        return;
                    }
                }


                try
                {
                    if (e.Stroke.StylusPoints.Count > 3)
                    {
                        Random random = new Random();
                        double _speed = GetPointSpeed(e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint(), e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint(), e.Stroke.StylusPoints[random.Next(0, e.Stroke.StylusPoints.Count - 1)].ToPoint());

                        RandWindow.randSeed = (int)(_speed * 100000 * 1000);
                    }
                }
                catch { }

                switch (Settings.Canvas.InkStyle)
                {
                    case 1:
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            string s = "";

                            for (int i = 0; i <= n; i++)
                            {
                                double speed = GetPointSpeed(e.Stroke.StylusPoints[Math.Max(i - 1, 0)].ToPoint(), e.Stroke.StylusPoints[i].ToPoint(), e.Stroke.StylusPoints[Math.Min(i + 1, n)].ToPoint());
                                s += speed.ToString() + "\t";
                                StylusPoint point = new StylusPoint();
                                if (speed >= 0.25)
                                {
                                    point.PressureFactor = (float)(0.5 - 0.3 * (Math.Min(speed, 1.5) - 0.3) / 1.2);
                                }
                                else if (speed >= 0.05)
                                {
                                    point.PressureFactor = (float)0.5;
                                }
                                else
                                {
                                    point.PressureFactor = (float)(0.5 + 0.4 * (0.05 - speed) / 0.05);
                                }
                                point.X = e.Stroke.StylusPoints[i].X;
                                point.Y = e.Stroke.StylusPoints[i].Y;
                                stylusPoints.Add(point);
                            }
                            //Label.Visibility = Visibility.Visible;
                            //Label.Content = s;
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                    case 0:
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            double pressure = 0.1;
                            int x = 10;
                            if (n == 1) return;
                            if (n >= x)
                            {
                                for (int i = 0; i < n - x; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)0.5;
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                                for (int i = n - x; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)((0.5 - pressure) * (n - i) / x + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            else
                            {
                                for (int i = 0; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)(0.4 * (n - i) / n + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                    case 3: //根据 mode == 0 改写，目前暂未完成
                        try
                        {
                            StylusPointCollection stylusPoints = new StylusPointCollection();
                            int n = e.Stroke.StylusPoints.Count - 1;
                            double pressure = 0.1;
                            int x = 8;
                            if (lastTouchDownTime < lastTouchUpTime)
                            {
                                double k = (lastTouchUpTime - lastTouchDownTime) / (n + 1); // 每个点之间间隔 k 毫秒
                                x = (int)(1000 / k); // 取 1000 ms 内的点
                            }

                            if (n >= x)
                            {
                                for (int i = 0; i < n - x; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)0.5;
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                                for (int i = n - x; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)((0.5 - pressure) * (n - i) / x + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            else
                            {
                                for (int i = 0; i <= n; i++)
                                {
                                    StylusPoint point = new StylusPoint();

                                    point.PressureFactor = (float)(0.4 * (n - i) / n + pressure);
                                    point.X = e.Stroke.StylusPoints[i].X;
                                    point.Y = e.Stroke.StylusPoints[i].Y;
                                    stylusPoints.Add(point);
                                }
                            }
                            e.Stroke.StylusPoints = stylusPoints;
                        }
                        catch
                        {

                        }
                        break;
                }
            }
            catch { }
        }

        private void SetNewBackupOfStroke()
        {
            lastTouchDownStrokeCollection = inkCanvas.Strokes.Clone();
            int whiteboardIndex = CurrentWhiteboardIndex;
            if (currentMode == 0)
            {
                whiteboardIndex = 0;
            }
            strokeCollections[whiteboardIndex] = lastTouchDownStrokeCollection;
        }

        public double GetDistance(Point point1, Point point2)
        {
            return Math.Sqrt((point1.X - point2.X) * (point1.X - point2.X) + (point1.Y - point2.Y) * (point1.Y - point2.Y));
        }

        #region 直线识别（ClassIn 式"停顿拉直"：停顿→拉直→拖拽转向→抬笔定型）

        // ---------- 触发参数 ----------
        private const double LineAssistMinLen = 40.0;      // 最短长度：短笔画（标点/部首）不参与
        private const int LineAssistHoldMs = 600;         // 停顿时长：按住笔尖不动 600ms 才触发。
                                                          // 防"一、二、三"误伤依据：写字收笔顿笔后通常
                                                          // 100~200ms 内即抬笔，够不到 600ms 门槛；
                                                          // 刻意画线时会主动按住等待，两拨人天然分开。
        private const double LineAssistMaxDevRatio = 0.12; // 触发时直度要求（宽松）：写字中途停顿的
                                                          // 弯笔画（竖弯钩画一半发呆）不触发。
        private const double LineAssistSnapDeg = 4.0;     // 定型时角度吸附：接近水平/垂直吸正（画坐标轴刚需）
        private const double LineAssistJitterPx = 2.0;   // 位移死区（DIP）：手写板/触摸屏静止接触时
                                                          // 仍持续上报亚像素~1px 抖动（鼠标静止则不再
                                                          // 发事件）。小于该值的"移动"视为手在抖/停顿
                                                          // 中，不刷新停顿计时——否则停顿永远检测不到。

        // ---------- 运行状态（单笔生命周期） ----------
        private bool _lineAssistTracking = false;    // 笔已按下，正在跟踪（计时开始）
        private bool _lineAssistArmed = false;      // 停顿已触发：进入"拉直+拖拽转向"模式
        private Point _lineAssistStart;            // 锚点（起笔位置 = 旋转轴心）
        private Point _lineAssistLastPos;          // 最新笔尖位置
        private DateTime _lineAssistLastMoveTime;  // 最后一次移动时刻（停顿检测基准）
        private List<Point> _lineAssistPoints;     // 到目前为止的采样点（触发时直度检查用）
        private DispatcherTimer _lineAssistTimer;  // 停顿检测定时器
        private System.Windows.Shapes.Line _lineAssistPreviewLine; // 预览直线（覆盖层，不挡输入）
        private InkCanvasEditingMode _lineAssistSavedMode = InkCanvasEditingMode.Ink; // 触发时的原模式（恢复用）
        private bool _lineAssistCommitted = false;     // 本次拉直是否已定型提交（防双提交/双删除）

        // ===== 输入事件入口（XAML 挂在 inkCanvas 上，笔/触摸/鼠标三路全覆盖） =====

        // 数位笔/触摸走 Stylus 事件（触摸会提升为 Stylus）
        private void inkCanvas_LineAssistStylusDown(object sender, StylusDownEventArgs e)
        {
            //输入落在画布上（选区覆盖层只覆盖选中框附近）→ 取消选中让当前这一笔直接写出，
            //不用先点一下空白。必须在各 return 之前：否则笔迹跟踪不启动时选中也无人取消
            DeselectStrokesForCanvasInput();
            LineAssistBegin(e.GetPosition(inkCanvas));
        }

        private void inkCanvas_LineAssistStylusMove(object sender, StylusEventArgs e)
        {
            LineAssistMove(e.GetPosition(inkCanvas));
        }

        private void inkCanvas_LineAssistStylusUp(object sender, StylusEventArgs e)
        {
            LineAssistEnd();
        }

        // 纯鼠标走鼠标事件（e.StylusDevice != null 表示是数位笔提升的鼠标事件，跳过防重复）
        private void inkCanvas_LineAssistMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return;
            //框外鼠标落点：取消选中（同 stylus 路径，鼠标按住拖动即可直接画出这一笔）
            DeselectStrokesForCanvasInput();
            LineAssistBegin(e.GetPosition(inkCanvas));
        }

        private void inkCanvas_LineAssistMouseMove(object sender, MouseEventArgs e)
        {
            if (e.StylusDevice != null) return;
            LineAssistMove(e.GetPosition(inkCanvas));
        }

        private void inkCanvas_LineAssistMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.StylusDevice != null) return;
            LineAssistEnd();
        }

        // ===== 状态机 =====

        /// <summary>笔落下：初始化跟踪状态并启动停顿检测定时器</summary>
        private void LineAssistBegin(Point p)
        {
            try
            {
                // 仅画笔模式 + 墨迹识别开启 + 非多指手势时启用
                if (!Settings.InkToShape.IsInkToShapeEnabled) return;
                if (inkCanvas.EditingMode != InkCanvasEditingMode.Ink) return;
                //注意：单指触摸书写时 dec.Count == 1（PreviewTouchDown 已登记），
                //必须放行——只有双指及以上的手势（dec.Count > 1）才排除。
                //原写法 dec.Count != 0 会把所有触摸/手写板书写全部拒之门外（鼠标不触发 Touch 事件所以没事）
                if (dec.Count > 1) return;

                _lineAssistTracking = true;
                _lineAssistArmed = false;
                _lineAssistCommitted = false;
                _lineAssistStart = p;
                _lineAssistLastPos = p;
                _lineAssistLastMoveTime = DateTime.Now;
                _lineAssistPoints = new List<Point> { p };

                if (_lineAssistTimer == null)
                {
                    _lineAssistTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(100) };
                    _lineAssistTimer.Tick += LineAssistTimerTick;
                }
                _lineAssistTimer.Start();
            }
            catch { }
        }

        /// <summary>笔尖移动：刷新位置与时刻；已进入拉直模式时预览线终点实时跟随（旋转/伸缩）</summary>
        private void LineAssistMove(Point p)
        {
            if (!_lineAssistTracking) return;

            //已触发拉直：笔尖实时拖动预览线（此时笔必然在动，无需死区判断）
            if (_lineAssistArmed)
            {
                _lineAssistLastPos = p;
                UpdateLineAssistPreview();
                return;
            }

            //位移死区（防手写板/触摸屏微抖导致停顿检测失效）：
            //相对上次"有效位置"的位移小于阈值 → 视为静止（手在抖或停顿中），
            //不刷新停顿计时、不记采样点；累计位移一旦超过阈值即视为真实移动。
            //（缓慢书写的鼠标也安全：位移相对"上次有效点"累计，慢移每几个事件就会被记一次）
            if (GetDistance(_lineAssistLastPos, p) >= LineAssistJitterPx)
            {
                _lineAssistLastPos = p;
                _lineAssistLastMoveTime = DateTime.Now;
                _lineAssistPoints.Add(p);
            }
        }

        /// <summary>定时检测：笔按住不动超时 → 直度检查通过 → 触发拉直</summary>
        private void LineAssistTimerTick(object sender, EventArgs e)
        {
            try
            {
                if (!_lineAssistTracking || _lineAssistArmed) return;

                //停顿时长不足则继续等
                if ((DateTime.Now - _lineAssistLastMoveTime).TotalMilliseconds < LineAssistHoldMs) return;

                //长度门槛
                double len = GetDistance(_lineAssistStart, _lineAssistLastPos);
                if (len < LineAssistMinLen) return;

                //直度门槛：所有采样点到"首点-停顿点"连线的最大垂距（叉积法）
                double ux = (_lineAssistLastPos.X - _lineAssistStart.X) / len;
                double uy = (_lineAssistLastPos.Y - _lineAssistStart.Y) / len;
                double maxDev = 0;
                foreach (var pt in _lineAssistPoints)
                {
                    double d = Math.Abs((pt.X - _lineAssistStart.X) * uy - (pt.Y - _lineAssistStart.Y) * ux);
                    if (d > maxDev) maxDev = d;
                }
                if (maxDev / len > LineAssistMaxDevRatio) return; //弯笔画（写字中途停顿）不触发

                //触发拉直：停顿已确认，进入"拉直 + 拖拽转向"模式
                _lineAssistArmed = true;

                // 关键：临时停止墨迹采集，消灭"两条线"。
                // 进入 armed 后笔还没抬，InkCanvas 仍在实时采集笔尖位置——拖动预览线
                // 转向时笔尖后面会拖着一条歪歪扭扭的"尾巴墨迹"，和预览直线并存。
                // 切到 None 让采集立即停止（歪线定格不再生长），抬笔定型后再恢复 Ink。
                // 注：不能走 DynamicRenderer.Enabled（保护成员，外部无法访问），此为公开 API 等效方案。
                try
                {
                    _lineAssistSavedMode = inkCanvas.EditingMode;   // 记住原模式（正常是 Ink）
                    inkCanvas.EditingMode = InkCanvasEditingMode.None;
                }
                catch { }

                //若停止采集时这条歪线已被中途提交进 Strokes（笔按了很久的场景），直接删掉——
                //最终反正会被直线替换，留着只会双线并存。走备份恢复路径，Ctrl+Z 还能救回。
                try
                {
                    List<Stroke> toRemove = null;
                    foreach (Stroke s in inkCanvas.Strokes)
                    {
                        if (s.StylusPoints.Count > 0)
                        {
                            var first = s.StylusPoints[0].ToPoint();
                            //起点吻合且是本次跟踪期间新增的笔迹（用首个点位置近似匹配）
                            if (Math.Abs(first.X - _lineAssistStart.X) < 1.0 && Math.Abs(first.Y - _lineAssistStart.Y) < 1.0)
                            {
                                if (toRemove == null) toRemove = new List<Stroke>();
                                toRemove.Add(s);
                            }
                        }
                    }
                    if (toRemove != null)
                    {
                        SetNewBackupOfStroke();
                        _currentCommitType = CommitReason.ShapeRecognition;
                        foreach (Stroke s in toRemove) inkCanvas.Strokes.Remove(s);
                        _currentCommitType = CommitReason.UserInput;
                    }
                }
                catch { }

                ShowLineAssistPreview();
            }
            catch { }
        }

        /// <summary>
        /// 抬笔：停表 + 定型提交。
        /// 注意：armed 后 EditingMode 已临时切 None，InkCanvas 大概率不再产出 StrokeCollected，
        /// 所以提交必须在这里闭环（用实时跟踪点生成直线），不能依赖 StrokeCollected。
        /// </summary>
        private void LineAssistEnd()
        {
            if (!_lineAssistTracking) return;
            _lineAssistTracking = false;
            try { _lineAssistTimer?.Stop(); } catch { }

            //armed 且未提交 → 抬笔即定型。暂不 ResetLineAssist：万一 StrokeCollected
            //仍随后触发（残留笔迹被提交的场景），交给它清残影后再复位；
            //若没触发（常态），由下方 ContextIdle 兜底复位，避免预览线残留。
            if (_lineAssistArmed && !_lineAssistCommitted)
            {
                try { CommitLineAssist(null); }
                catch { }
            }

            //等事件队列排空（StrokeCollected 处理完残影后）再兜底清一次
            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, (Action)(() =>
            {
                if (!_lineAssistTracking && _lineAssistArmed) ResetLineAssist();
            }));
        }

        /// <summary>显示预览直线（颜色/粗细跟随当前画笔）</summary>
        private void ShowLineAssistPreview()
        {
            try
            {
                if (_lineAssistPreviewLine == null)
                {
                    _lineAssistPreviewLine = new System.Windows.Shapes.Line
                    {
                        IsHitTestVisible = false //纯显示，不挡任何输入
                    };
                    Panel.SetZIndex(_lineAssistPreviewLine, 9999); //盖在最上层
                }
                _lineAssistPreviewLine.StrokeThickness = inkCanvas.DefaultDrawingAttributes.Width;
                _lineAssistPreviewLine.Stroke = new SolidColorBrush(inkCanvas.DefaultDrawingAttributes.Color);

                if (!Main_Grid.Children.Contains(_lineAssistPreviewLine))
                    Main_Grid.Children.Add(_lineAssistPreviewLine);
                UpdateLineAssistPreview();
            }
            catch { }
        }

        /// <summary>预览线坐标换算：inkCanvas 坐标系 → 主 Grid 坐标系（二者可能有边距差）</summary>
        private void UpdateLineAssistPreview()
        {
            if (_lineAssistPreviewLine == null) return;
            try
            {
                var toGrid = inkCanvas.TransformToAncestor(Main_Grid);
                var s = toGrid.Transform(_lineAssistStart);
                var t = toGrid.Transform(_lineAssistLastPos);
                _lineAssistPreviewLine.X1 = s.X; _lineAssistPreviewLine.Y1 = s.Y;
                _lineAssistPreviewLine.X2 = t.X; _lineAssistPreviewLine.Y2 = t.Y;
            }
            catch { }
        }

        /// <summary>
        /// 抬笔定型：手绘笔迹 → 笔直的线。
        /// 端点 = 起笔点 + 抬笔点（拖拽转向的最终位置），接近水平/垂直时吸正；
        /// 走 ShapeRecognition 替换型历史（Ctrl+Z 撤回手绘原迹）；打组标签 + 自动选中。
        /// rawStroke 可为 null：armed 后墨迹采集已停（EditingMode=None），常态下没有原迹可传。
        /// </summary>
        private void CommitLineAssist(Stroke rawStroke)
        {
            //端点必须用实时跟踪值：rawStroke（若存在）的末点停在触发时刻，
            //不是拖拽转向后的最终位置——以 _lineAssistLastPos 为准，否则定型线与预览线不一致。
            var start = _lineAssistStart;
            var end = _lineAssistLastPos;

            //角度吸附（与水平/垂直夹角 < 4° 时终点对齐吸正，起点不动）
            double angle = Math.Atan2(Math.Abs(end.Y - start.Y), Math.Abs(end.X - start.X)) * 180.0 / Math.PI;
            if (angle < LineAssistSnapDeg) end = new Point(end.X, start.Y);        //吸水平
            else if (angle > 90 - LineAssistSnapDeg) end = new Point(start.X, end.Y); //吸垂直

            //笔迹属性：优先取手绘原迹（颜色粗细跟手），无原迹时跟当前画笔（与预览线一致）
            DrawingAttributes da = (rawStroke != null ? rawStroke.DrawingAttributes : inkCanvas.DefaultDrawingAttributes).Clone();

            //生成两点直线笔迹
            var straight = new Stroke(new StylusPointCollection(new List<Point> { start, end }))
            {
                DrawingAttributes = da
            };

            //替换（与圆/矩形识别同一套替换型历史，Ctrl+Z 可撤回手绘原迹）
            SetNewBackupOfStroke();
            _currentCommitType = CommitReason.ShapeRecognition;
            if (rawStroke != null) inkCanvas.Strokes.Remove(rawStroke); //残影笔迹若已提交则删掉
            inkCanvas.Strokes.Add(straight);
            _currentCommitType = CommitReason.UserInput;

            //拉直的线也是"图形"：打标签 + 自动选中，可整组拖动/撤销
            InsertGraphStrokes(new StrokeCollection(new[] { straight }));

            _lineAssistCommitted = true; //标记已提交，后续 StrokeCollected 兜底只清残影、不再重复提交
        }

        /// <summary>复位：停表、恢复采集模式（Ink）、移除预览线</summary>
        private void ResetLineAssist()
        {
            _lineAssistTracking = false;
            _lineAssistArmed = false;
            try { _lineAssistTimer?.Stop(); } catch { }
            //恢复触发前保存的采集模式（消除"两条线"时临时切的 None，异常路径也保证恢复）
            try { if (inkCanvas.EditingMode == InkCanvasEditingMode.None) inkCanvas.EditingMode = _lineAssistSavedMode; } catch { }
            try
            {
                if (_lineAssistPreviewLine != null && Main_Grid.Children.Contains(_lineAssistPreviewLine))
                    Main_Grid.Children.Remove(_lineAssistPreviewLine);
            }
            catch { }
        }

        #endregion

        public double GetPointSpeed(Point point1, Point point2, Point point3)
        {
            return (Math.Sqrt((point1.X - point2.X) * (point1.X - point2.X) + (point1.Y - point2.Y) * (point1.Y - point2.Y))
                + Math.Sqrt((point3.X - point2.X) * (point3.X - point2.X) + (point3.Y - point2.Y) * (point3.Y - point2.Y)))
                / 20;
        }

        public Point[] FixPointsDirection(Point p1, Point p2)
        {
            if (Math.Abs(p1.X - p2.X) / Math.Abs(p1.Y - p2.Y) > 8)
            {
                //水平
                double x = Math.Abs(p1.Y - p2.Y) / 2;
                if (p1.Y > p2.Y)
                {
                    p1.Y -= x;
                    p2.Y += x;
                }
                else
                {
                    p1.Y += x;
                    p2.Y -= x;
                }
            }
            else if (Math.Abs(p1.Y - p2.Y) / Math.Abs(p1.X - p2.X) > 8)
            {
                //垂直
                double x = Math.Abs(p1.X - p2.X) / 2;
                if (p1.X > p2.X)
                {
                    p1.X -= x;
                    p2.X += x;
                }
                else
                {
                    p1.X += x;
                    p2.X -= x;
                }
            }

            return new Point[2] { p1, p2 };
        }

        public StylusPointCollection GenerateFakePressureTriangle(StylusPointCollection points)
        {
            var newPoint = new StylusPointCollection();
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            var cPoint = GetCenterPoint(points[0], points[1]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            cPoint = GetCenterPoint(points[1], points[2]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            cPoint = GetCenterPoint(points[2], points[0]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            return newPoint;
        }

        public StylusPointCollection GenerateFakePressureRectangle(StylusPointCollection points)
        {
            var newPoint = new StylusPointCollection();
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            var cPoint = GetCenterPoint(points[0], points[1]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[1].X, points[1].Y, (float)0.4));
            cPoint = GetCenterPoint(points[1], points[2]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[2].X, points[2].Y, (float)0.4));
            cPoint = GetCenterPoint(points[2], points[3]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[3].X, points[3].Y, (float)0.4));
            newPoint.Add(new StylusPoint(points[3].X, points[3].Y, (float)0.4));
            cPoint = GetCenterPoint(points[3], points[0]);
            newPoint.Add(new StylusPoint(cPoint.X, cPoint.Y, (float)0.8));
            newPoint.Add(new StylusPoint(points[0].X, points[0].Y, (float)0.4));
            return newPoint;
        }

        public Point GetCenterPoint(Point point1, Point point2)
        {
            return new Point((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2);
        }

        public StylusPoint GetCenterPoint(StylusPoint point1, StylusPoint point2)
        {
            return new StylusPoint((point1.X + point2.X) / 2, (point1.Y + point2.Y) / 2);
        }

        #endregion
    }
}
