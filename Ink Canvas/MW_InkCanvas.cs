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
    /// <summary>MainWindow 分部类：墨迹画布基础功能（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Ink Canvas Functions

        DrawingAttributes drawingAttributes;
        private void loadPenCanvas()
        {
            SetColors();
            try
            {
                //drawingAttributes = new DrawingAttributes();
                drawingAttributes = inkCanvas.DefaultDrawingAttributes;
                drawingAttributes.Color = ((SolidColorBrush)BtnColorRed.Background).Color;

                drawingAttributes.Height = 2.5;
                drawingAttributes.Width = 2.5;
                // 笔迹平滑（2026-09-12 起改为恒关）：WPF 的贝塞尔曲线拟合（FitToCurve）会把
                // 识别出的图形和汉字的直角"一刀切"磨圆，所以画布默认画笔属性不再走它。
                // 手写笔迹的平滑改由"保角平滑"在抬笔后于数据层完成（MW_PreserveCornerSmoothing.cs），
                // 由设置项 Settings.Canvas.FitToCurve 开关控制（见 inkCanvas_StrokeCollected）。
                drawingAttributes.FitToCurve = false;

                inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                inkCanvas.Gesture += InkCanvas_Gesture;

                //初始化悬浮条画笔区：笔图标颜色/高亮 + 6 色点选中描边（默认红笔）
                UpdatePenIconColor();
                UpdatePenIconHighlight();
                UpdateFloatBarColorDots();
            }
            catch { }
        }
        //ApplicationGesture lastApplicationGesture = ApplicationGesture.AllGestures;
        DateTime lastGestureTime = DateTime.Now;
        private void InkCanvas_Gesture(object sender, InkCanvasGestureEventArgs e)
        {
            ReadOnlyCollection<GestureRecognitionResult> gestures = e.GetGestureRecognitionResults();
            try
            {
                foreach (GestureRecognitionResult gest in gestures)
                {
                    //Trace.WriteLine(string.Format("Gesture: {0}, Confidence: {1}", gest.ApplicationGesture, gest.RecognitionConfidence));
                    if (StackPanelPPTControls.Visibility == Visibility.Visible)
                    {
                        if (gest.ApplicationGesture == ApplicationGesture.Left)
                        {
                            BtnPPTSlidesDown_Click(BtnPPTSlidesDown, null);
                        }
                        if (gest.ApplicationGesture == ApplicationGesture.Right)
                        {
                            BtnPPTSlidesUp_Click(BtnPPTSlidesUp, null);
                        }
                    }
                }
            }
            catch { }
        }

        private void inkCanvas_EditingModeChanged(object sender, RoutedEventArgs e)
        {
            var inkCanvas1 = sender as InkCanvas;
            if (inkCanvas1 == null) return;
            if (Settings.Canvas.IsShowCursor)
            {
                if (inkCanvas1.EditingMode == InkCanvasEditingMode.Ink || drawingShapeMode != 0)
                {
                    inkCanvas1.ForceCursor = true;
                }
                else
                {
                    inkCanvas1.ForceCursor = false;
                }
            }
            else
            {
                inkCanvas1.ForceCursor = false;
            }
            // 原"编辑模式变回画笔时 toggle 擦除方式"已废弃（不可见的隐晦切换）；
            // 擦除方式只在橡皮设置面板里切（SetEraserMode，见 MW_EraserSettings.cs）

            if (inkCanvas.EditingMode == InkCanvasEditingMode.Select)
            {
                // 高亮色统一取自 InkSpec.xaml 的 HighlightBrush（原为硬编码 Color.FromRgb(0,136,255)）
                SetSelectToolColor(ResolveToolHighlightColor());
            }
            else
            {
                SetSelectToolColor(FloatBarForegroundColor);
            }

            // 悬浮条笔图标高亮同步：画笔模式亮、橡皮/选择等熄灭（图形模式由 drawingShapeMode 判断）
            UpdatePenIconHighlight();

            // 选择图标高亮同步：选择模式亮淡蓝底+蓝边，其他模式熄灭
            UpdateSelectIconHighlight();
        }

        /// <summary>更新选择图标的颜色（虚线框描边 + 箭头填充共用同一画刷）</summary>
        private void SetSelectToolColor(Color color)
        {
            var brush = new SolidColorBrush(color);
            PathSelectBox.Stroke = brush;
            PathSelectCursor.Fill = brush;
        }

        #endregion Ink Canvas
    }
}
