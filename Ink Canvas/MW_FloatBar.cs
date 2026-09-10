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
using System.Windows.Shapes;
using System.Windows.Threading;
using Application = System.Windows.Application;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using Timer = System.Timers.Timer;

namespace Ink_Canvas
{
    /// <summary>MainWindow 分部类：浮动工具栏（含拖动）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Float Bar

        private void HideSubPanels()
        {
            BorderTools.Visibility = Visibility.Collapsed;
            ClosePenSettingsPanel(); // 笔设置面板（替代原 BorderPenWidth 小面板）
            CloseSelectionModePanel(); // 选择方式面板（矩形框选/自由选择）
            CloseEraserSettingsPanel(); // 橡皮设置面板（含滑动清屏，原 BorderClearInDelete 已废弃）
        }


        private void BorderPenColorBlack_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorBlack_Click(BtnColorBlack, null);
            HideSubPanels();
        }

        private void BorderPenColorRed_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorRed_Click(BtnColorRed, null);
            HideSubPanels();
        }

        private void BorderPenColorGreen_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorGreen_Click(BtnColorGreen, null);
            HideSubPanels();
        }

        private void BorderPenColorBlue_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorBlue_Click(BtnColorBlue, null);
            HideSubPanels();
        }

        private void BorderPenColorYellow_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnColorYellow_Click(BtnColorYellow, null);
            HideSubPanels();
        }

        private void BorderPenColorWhite_MouseUp(object sender, MouseButtonEventArgs e)
        {
            inkCanvas.DefaultDrawingAttributes.Color = StringToColor("#FFFEFEFE");
            inkColor = 5;
            ColorSwitchCheck();
            HideSubPanels();
        }

        #region 画笔图标 + 笔粗细面板

        // 笔粗细三档对应的实际笔宽（设置页滑条值 = 笔宽 × 2，范围 0.5~10）
        private const double PenWidthThin = 1.5;
        private const double PenWidthMedium = 3;
        private const double PenWidthThick = 6;

        // 工具选中高亮色：与选择工具高亮一致的项目蓝（#0088FF）
        private static readonly SolidColorBrush PenToolHighlightBrush = new SolidColorBrush(Color.FromRgb(0, 136, 255));
        // 淡蓝选中底（约 15% 不透明度），用于笔图标高亮底和粗细档位选中底
        private static readonly SolidColorBrush PenToolHighlightSoftBrush = new SolidColorBrush(Color.FromArgb(38, 0, 136, 255));

        /// <summary>笔图标单击：切回画笔模式 + 展开/收起笔设置面板（笔种类/粗细/颜色/笔锋）</summary>
        private void PenIcon_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // 非画笔状态（橡皮/选择/图形）时先切回画笔：
            // ColorSwitchCheck 会无条件恢复 Ink 模式、熄灭图形图标高亮，并刷新笔图标状态
            if (inkCanvas.EditingMode != InkCanvasEditingMode.Ink || drawingShapeMode != 0)
                ColorSwitchCheck();

            // 面板显隐 + 定位（MW_PenSettings.cs）
            TogglePenSettingsPanel();
        }

        /// <summary>更新笔图标的笔身颜色，使其始终显示当前笔颜色</summary>
        private void UpdatePenIconColor()
        {
            Brush brush;
            switch (inkColor)
            {
                case 0: brush = BtnColorBlack.Background; break;
                case 2: brush = BtnColorGreen.Background; break;
                case 3: brush = BtnColorBlue.Background; break;
                case 4: brush = BtnColorYellow.Background; break;
                case 5: brush = new SolidColorBrush(StringToColor("#FFFEFEFE")); break;
                case 6: brush = new SolidColorBrush(ExtraColors[0]); break; // 橙
                case 7: brush = new SolidColorBrush(ExtraColors[1]); break; // 品红（紫红）
                case 8: brush = new SolidColorBrush(ExtraColors[2]); break; // 青
                case 9: // 自定义色（未设置过时兜底红色，笔图标不至于无色）
                    var custom = GetSavedCustomColor();
                    brush = custom.HasValue ? new SolidColorBrush(custom.Value) : BtnColorRed.Background;
                    break;
                default: brush = BtnColorRed.Background; break; // inkColor == 1（红色，默认）
            }
            PathPenIconBody.Fill = brush;

            // 快捷换色条同步：色点填充（红绿蓝黄随配色）+ 选中环（当前色亮蓝环）。
            // UpdatePenIconColor 是换色/白黑板配色切换/启动恢复的公共汇合点，在此搭车最省接线
            RefreshQuickColorStrip();
        }

        /// <summary>画笔模式时给笔图标加淡蓝底+蓝边高亮，其他工具（橡皮/选择/图形）时熄灭</summary>
        private void UpdatePenIconHighlight()
        {
            bool isPenActive = inkCanvas.EditingMode == InkCanvasEditingMode.Ink && drawingShapeMode == 0;
            BorderPenIconHighlight.Background = isPenActive ? PenToolHighlightSoftBrush : Brushes.Transparent;
            BorderPenIconHighlight.BorderBrush = isPenActive ? PenToolHighlightBrush : Brushes.Transparent;
        }

        /// <summary>同步笔设置面板颜色选中环：当前色中心显示细蓝环（替代原浮动条 6 色点）</summary>
        private void UpdateFloatBarColorDots()
        {
            PenSettingsPanel.UpdateColorHighlight(inkColor);
        }

        #region 快捷换色条（常驻浮动条：对色格——黑⇄白 / 红⇄青 / 蓝⇄黄 / 绿⇄品红）

        /// <summary>
        /// 快捷换色条色点单击：Tag = 格编号 → 在对色间循环切换。
        /// 切换规则（当前色是这格的两色之一 → 切到另一个；否则 → 切到主色）：
        /// - 黑白格：白板下当前黑→切白、当前白→切黑；黑板下主副互换（主=白）
        /// - 红青格：红⇄青；蓝黄格：蓝⇄黄；绿品红格：绿⇄品红
        /// 切换走 SelectPenColorByIndex：任何工具状态下点色点都会自动切回画笔。
        /// </summary>
        private void QuickColorDot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is Grid grid) || !int.TryParse(grid.Tag?.ToString(), out int slot)) return;

            var (primary, secondary) = GetQuickColorPair(slot);
            int target;
            if (inkColor == primary) target = secondary;      // 当前是主色 → 切副色
            else if (inkColor == secondary) target = primary; // 当前是副色 → 切主色
            else target = primary;                            // 当前是别的色 → 先选主色

            SelectPenColorByIndex(target);
        }

        /// <summary>
        /// 取对色格的（主色索引，副色索引）。
        /// 索引语义与面板色格一致：0黑 1红 2绿 3蓝 4黄 5白 7品红 8青。
        /// 黑白格主色随板面：白板黑(0)/黑板白(5)——黑板下黑看不见，主书写色必须是白。
        /// </summary>
        private (int primary, int secondary) GetQuickColorPair(int slot)
        {
            bool isLight = currentMode != 0 && !Settings.Canvas.UsingWhiteboard;
            switch (slot)
            {
                case 0: return isLight ? (0, 5) : (5, 0); // 黑⇄白（主色随板面）
                case 1: return (1, 8);                    // 红⇄青
                case 2: return (3, 4);                    // 蓝⇄黄
                default: return (2, 7);                   // 绿⇄品红
            }
        }

        /// <summary>当前笔色在这对色中的位置：0=主色 1=副色 -1=都不是（大区/角标/选中环共用）</summary>
        private int GetPairActiveIndex((int primary, int secondary) pair)
        {
            if (inkColor == pair.primary) return 0;
            if (inkColor == pair.secondary) return 1;
            return -1;
        }

        /// <summary>
        /// 套用一格的显示状态（PS 前景色/背景色模式——面积比例 + 上下层级指示当前色）：
        /// - 默认（当前是主色或不在本对，如橙/自定义色）→ 大色区主色 + 右下小角标副色
        /// - 当前正在用副色 → 大区副色 + 小角主色（一点即回主色）
        /// 小角标常显、永不隐藏：每格恒定可见两色，用户随时一眼读出格内有什么颜色可用。
        /// 面积对比（大区 11px : 角标 5px ≈ 5:1）醒目易读，"当前色是否在这格"由选中环单独表达。
        /// </summary>
        private void ApplyPairDot(Ellipse mainDot, Ellipse miniDot,
            (int primary, int secondary) pair, Color primaryC, Color secondaryC)
        {
            if (GetPairActiveIndex(pair) == 1)
            {
                // 当前用副色 → 大区显示副色、小角显示主色
                mainDot.Fill = new SolidColorBrush(secondaryC);
                miniDot.Fill = new SolidColorBrush(primaryC);
            }
            else
            {
                // 当前是主色或不在本对 → 大区主色 + 小角副色（默认布局）
                mainDot.Fill = new SolidColorBrush(primaryC);
                miniDot.Fill = new SolidColorBrush(secondaryC);
            }
            miniDot.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 刷新快捷换色条状态（由 UpdatePenIconColor 搭车调用——换色/配色切换/启动恢复都会走到）：
        /// 1. 大色区 = 主色（仅当前用副色时对调为副色）；
        /// 2. 右下小角标 = 另一色（常显，随时可读出这格的两个颜色）；
        /// 3. 黑白格主色随板面（白板黑/黑板白）+ 细灰描边防隐身；
        /// 4. 选中环：当前色是这格任一色 → 亮环；橙/自定义色不在四格内 → 全灭。
        /// </summary>
        private void RefreshQuickColorStrip()
        {
            try
            {
                // 板面判断：与 SetColors 完全同款表达式（Light 配色 ⇔ 白板）
                bool isLight = currentMode != 0 && !Settings.Canvas.UsingWhiteboard;

                var bw = GetQuickColorPair(0);
                var rc = GetQuickColorPair(1);
                var by = GetQuickColorPair(2);
                var gm = GetQuickColorPair(3);

                // ---- 黑白格：主色随板面（白板黑/黑板白），主点细灰描边防隐身 ----
                Color bwPrimary = isLight ? Colors.Black : StringToColor("#FFFEFEFE");
                Color bwSecondary = isLight ? StringToColor("#FFFEFEFE") : Colors.Black;
                ApplyPairDot(QuickColorDotBW, QuickColorMiniBW, bw, bwPrimary, bwSecondary);
                QuickColorDotBW.Stroke = new SolidColorBrush(Color.FromArgb(128, 128, 128, 128));
                QuickColorDotBW.StrokeThickness = 0.75;

                // ---- 红青格：红与右面板同源（白板纯红/黑板亮红），青为固定扩展色 ----
                ApplyPairDot(QuickColorDotRC, QuickColorMiniRC, rc,
                    ((SolidColorBrush)BtnColorRed.Background).Color, ExtraColors[2]);

                // ---- 蓝黄格：蓝黄都与右面板同源（黑板下黄为纯黄、蓝提亮） ----
                ApplyPairDot(QuickColorDotBY, QuickColorMiniBY, by,
                    ((SolidColorBrush)BtnColorBlue.Background).Color,
                    ((SolidColorBrush)BtnColorYellow.Background).Color);

                // ---- 绿品红格：绿与右面板同源，品红为固定扩展色 ----
                ApplyPairDot(QuickColorDotGM, QuickColorMiniGM, gm,
                    ((SolidColorBrush)BtnColorGreen.Background).Color, ExtraColors[1]);

                // ---- 选中环：当前色 = 这层任一色 → 亮环（与大区纯色互为印证） ----
                QuickColorRingBW.Visibility = GetPairActiveIndex(bw) >= 0 ? Visibility.Visible : Visibility.Collapsed;
                QuickColorRingRC.Visibility = GetPairActiveIndex(rc) >= 0 ? Visibility.Visible : Visibility.Collapsed;
                QuickColorRingBY.Visibility = GetPairActiveIndex(by) >= 0 ? Visibility.Visible : Visibility.Collapsed;
                QuickColorRingGM.Visibility = GetPairActiveIndex(gm) >= 0 ? Visibility.Visible : Visibility.Collapsed;
            }
            catch { }
        }

        #endregion

        #endregion

        private void SymbolIconUndo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnUndo_Click(BtnUndo, null);
            HideSubPanels();
        }

        private void SymbolIconRedo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnRedo_Click(BtnRedo, null);
            HideSubPanels();
        }

        private async void SymbolIconCursor_Click(object sender, RoutedEventArgs e)
        {
            if (currentMode != 0)
            {
                ImageBlackboard_MouseUp(null, null);
            }
            else
            {
                BtnHideInkCanvas_Click(BtnHideInkCanvas, null);

                if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                {
                    if (ViewboxFloatingBar.Margin == new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200))
                    {
                        await Task.Delay(100);
                        ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                    }
                }
            }

            SetColors();
        }

        private void SymbolIconDelete_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // 垃圾桶图标已移除（清屏迁入橡皮面板的滑动清屏）；保留处理器避免旧事件绑定报错，
            // 直接转调同一逻辑体（ExecuteSlideClearScreen，见 MW_EraserSettings.cs）
            if (sender != lastBorderMouseDownObject) return;
            ExecuteSlideClearScreen();
        }

        private void SymbolIconSettings_Click(object sender, RoutedEventArgs e)
        {
            BtnSettings_Click(BtnSettings, null);
            HideSubPanels();
        }

        /// <summary>上次点选择图标的时间（双击检测：500ms 内两击 = 全选，恢复经典交互）</summary>
        private DateTime _lastSelectIconClickTime = DateTime.MinValue;

        private void SymbolIconSelect_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // 双击（500ms 内两击）→ 全选当前页墨迹（经典交互回归：点两下选择图标=全选）
            var now = DateTime.Now;
            if ((now - _lastSelectIconClickTime).TotalMilliseconds < 500)
            {
                _lastSelectIconClickTime = DateTime.MinValue;
                CloseSelectionModePanel();
                SelectAllStrokes();
                return;
            }
            _lastSelectIconClickTime = now;

            // 面板已打开 → 再点图标仅收起面板（toggle），不重复切换模式
            if (SelectionModePanel.Visibility == Visibility.Visible)
            {
                CloseSelectionModePanel();
                return;
            }

            BtnSelect_Click(BtnSelect, null);

            ImageEraser.Visibility = Visibility.Visible;
            ViewboxBtnColorBlackContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorBlueContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorGreenContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorRedContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorYellowContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));

            HideSubPanels();

            // 选择方式面板随选择模式弹出（矩形框选/自由选择/全选）
            ToggleSelectionModePanel();
        }

        private void SymbolIconScreenshot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnScreenshot_Click(BtnScreenshot, null);
        }

        Point pointDesktop = new Point(-1, -1); //用于记录上次进入PPT或白板时的坐标
        Point pointPPT = new Point(-1, -1); //用于记录上次在PPT中打开白板时的坐标

        private void ImageBlackboard_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (currentMode == 0)
            {
                //进入黑板
                if (BtnPPTSlideShowEnd.Visibility == Visibility.Collapsed)
                {
                    pointDesktop = new Point(ViewboxFloatingBar.Margin.Left, ViewboxFloatingBar.Margin.Top);
                }
                else
                {
                    pointPPT = new Point(ViewboxFloatingBar.Margin.Left, ViewboxFloatingBar.Margin.Top);
                }
                //ViewboxFloatingBar.Margin = new Thickness(10, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);

                new Thread(new ThreadStart(() =>
                {
                    Thread.Sleep(100);
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                    });
                })).Start();
                if (Settings.Canvas.UsingWhiteboard)
                {
                    BorderPenColorBlack_MouseUp(null, null);
                }
                else
                {
                    BorderPenColorWhite_MouseUp(null, null);
                }
            }
            else
            {
                //关闭黑板
                if (isInMultiTouchMode) BorderMultiTouchMode_MouseUp(null, null);

                if (BtnPPTSlideShowEnd.Visibility == Visibility.Collapsed)
                {
                    if (pointDesktop != new Point(-1, -1))
                    {
                        ViewboxFloatingBar.Margin = new Thickness(pointDesktop.X, pointDesktop.Y, -2000, -200);
                        pointDesktop = new Point(-1, -1);
                    }
                }
                else
                {
                    new Thread(new ThreadStart(() =>
                    {
                        Thread.Sleep(100);
                        Application.Current.Dispatcher.Invoke(() =>
                        {
                            ViewboxFloatingBar.Margin = new Thickness((SystemParameters.PrimaryScreenWidth - ViewboxFloatingBar.ActualWidth) / 2, SystemParameters.PrimaryScreenHeight - 60, -2000, -200);
                        });
                    })).Start();
                }
                BorderPenColorRed_MouseUp(null, null);
            }
            BtnSwitch_Click(BtnSwitch, null);

            BtnExit.Foreground = Brushes.White;
            ThemeManager.Current.ApplicationTheme = ApplicationTheme.Dark;
            SetColors();
            SetColorByIndex();
            if (currentMode == 0 && inkCanvas.Strokes.Count == 0 && BtnPPTSlideShowEnd.Visibility != Visibility.Visible)
            {
                BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
            }
        }

        /// <summary>
        /// 橡皮图标单击（新交互，与笔图标同款模式）：
        /// 进入上次的擦除方式（不再 toggle）+ 展开/收起橡皮设置面板。
        /// 擦除方式/大小只在面板里切（原"点图标 toggle 擦除方式"已废弃——不可见、易误切）。
        /// </summary>
        private void ImageEraser_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            // 进入橡皮模式（BtnErase_Click 已改为固定进入当前方式，不 toggle）
            BtnErase_Click(BtnErase, e);

            // 面板显隐 + 定位（MW_EraserSettings.cs）
            ToggleEraserSettingsPanel();
        }

        private void ImageCountdownTimer_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BtnCountdownTimer_Click(BtnCountdownTimer, null);
        }

        private void SymbolIconRand_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BtnRand_Click(BtnRand, null);
        }

        private void SymbolIconRandOne_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            new RandWindow(true).ShowDialog();
        }

        private void GridInkReplayButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BorderDrawShape.Visibility = Visibility.Collapsed;

            InkCanvasForInkReplay.Visibility = Visibility.Visible;
            inkCanvas.Visibility = Visibility.Collapsed;
            isStopInkReplay = false;
            InkCanvasForInkReplay.Strokes.Clear();
            StrokeCollection strokes = inkCanvas.Strokes.Clone();
            if (inkCanvas.GetSelectedStrokes().Count != 0)
            {
                strokes = inkCanvas.GetSelectedStrokes().Clone();
            }
            int k = 1, i = 0;
            new Thread(new ThreadStart(() =>
            {
                foreach (Stroke stroke in strokes)
                {
                    //Thread.Sleep(100);
                    //Application.Current.Dispatcher.Invoke(() =>
                    //{
                    //    InkCanvasForInkReplay.Strokes.Add(stroke);
                    //});
                    StylusPointCollection stylusPoints = new StylusPointCollection();
                    if (stroke.StylusPoints.Count == 629) //圆或椭圆
                    {
                        Stroke s = null;
                        foreach (StylusPoint stylusPoint in stroke.StylusPoints)
                        {
                            if (i++ >= 50)
                            {
                                i = 0;
                                Thread.Sleep(10);
                                if (isStopInkReplay) return;
                            }
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    InkCanvasForInkReplay.Strokes.Remove(s);
                                }
                                catch { }
                                stylusPoints.Add(stylusPoint);
                                s = new Stroke(stylusPoints.Clone());
                                s.DrawingAttributes = stroke.DrawingAttributes;
                                InkCanvasForInkReplay.Strokes.Add(s);
                            });
                        }
                    }
                    else
                    {
                        Stroke s = null;
                        foreach (StylusPoint stylusPoint in stroke.StylusPoints)
                        {
                            if (i++ >= k)
                            {
                                i = 0;
                                Thread.Sleep(10);
                                if (isStopInkReplay) return;
                            }
                            Application.Current.Dispatcher.Invoke(() =>
                            {
                                try
                                {
                                    InkCanvasForInkReplay.Strokes.Remove(s);
                                }
                                catch { }
                                stylusPoints.Add(stylusPoint);
                                s = new Stroke(stylusPoints.Clone());
                                s.DrawingAttributes = stroke.DrawingAttributes;
                                InkCanvasForInkReplay.Strokes.Add(s);
                            });
                        }
                    }
                }
                Thread.Sleep(100);
                Application.Current.Dispatcher.Invoke(() =>
                {
                    InkCanvasForInkReplay.Visibility = Visibility.Collapsed;
                    inkCanvas.Visibility = Visibility.Visible;
                });
            })).Start();
        }
        bool isStopInkReplay = false;
        private void InkCanvasForInkReplay_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                InkCanvasForInkReplay.Visibility = Visibility.Collapsed;
                inkCanvas.Visibility = Visibility.Visible;
                isStopInkReplay = true;
            }
        }

        private void SymbolIconTools_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (BorderTools.Visibility == Visibility.Visible)
            {
                BorderTools.Visibility = Visibility.Collapsed;
            }
            else
            {
                // 板面组跟随白板状态：非白板模式整组灰显（"不适用功能灰显"规范，与笔面板笔锋区同款）。
                // 主栏那两个图标是靠 Visibility 绑定直接隐藏，但面板里若也隐藏会让布局跳动，故改为灰显。
                bool boardActive = GridBackgroundCover.Visibility == Visibility.Visible;
                MoreToolsBoardSection.IsEnabled = boardActive;
                MoreToolsBoardSection.Opacity = boardActive ? 1.0 : 0.35;

                BorderTools.Visibility = Visibility.Visible;
            }
        }


        #region 更多面板条目（BorderTools 扩充内容）

        /// <summary>
        /// 更多面板里的条目尽量直接复用现成 handler（多人物写 BorderMultiTouchMode_MouseUp、重做 SymbolIconRedo_MouseUp 等）。
        /// 仅以下两个需要包一层：
        /// </summary>

        /// <summary>更多面板·白板底纹：菜单以本条目自身为锚点弹出
        /// （不写死某个元素当锚点，这样菜单会从被点击的条目位置弹出，而不是跑到主栏或别处）</summary>
        private void MoreToolsWhiteboardPattern_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ShowWhiteboardPatternMenu(sender as FrameworkElement, placeBelow: false);
        }

        /// <summary>更多面板·打开设置</summary>
        private void MoreToolsSettings_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BorderTools.Visibility = Visibility.Collapsed;
            BtnSettings_Click(BtnSettings, null);
        }

        /// <summary>更多面板·检查更新（手动检查：已是最新版本时会给出提示）</summary>
        private void MoreToolsCheckUpdate_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BorderTools.Visibility = Visibility.Collapsed;
            (Application.Current as App)?.CheckForUpdate(true);
        }

        /// <summary>更多面板·白板 / 黑板切换（从白板底纹菜单顶部拆出来的独立入口；
        /// 复用 BtnSwitchTheme_Click 的全套联动：板面色、UI 深浅主题、白板/黑板文案）</summary>
        private void MoreToolsBoardTheme_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnSwitchTheme_Click(null, null);
        }

        /// <summary>更多面板·查看快捷键（原笑脸右键菜单 →「快捷键 → 查看快捷键」）</summary>
        private void MoreToolsShowShortcuts_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BorderTools.Visibility = Visibility.Collapsed;
            MenuItemShowShortcuts_Click(null, null);
        }

        /// <summary>更多面板·退出（原主栏退出图标 + 笑脸右键菜单「退出」，内部有二次确认）</summary>
        private void MoreToolsExit_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnExit_Click(null, null);
        }

        #endregion


        #region Drag

        bool isDragDropInEffect = false;
        Point pos = new Point();
        Point downPos = new Point();

        void Element_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragDropInEffect)
            {
                FrameworkElement currEle = sender as FrameworkElement;
                double xPos = e.GetPosition(null).X - pos.X + currEle.Margin.Left;
                double yPos = e.GetPosition(null).Y - pos.Y + currEle.Margin.Top;
                currEle.Margin = new Thickness(xPos, yPos, 0, 0);
                pos = e.GetPosition(null);
            }
        }

        void Element_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {

            FrameworkElement fEle = sender as FrameworkElement;
            isDragDropInEffect = true;
            pos = e.GetPosition(null);
            fEle.CaptureMouse();
            fEle.Cursor = Cursors.Hand;
        }

        void Element_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (isDragDropInEffect)
            {
                FrameworkElement ele = sender as FrameworkElement;
                isDragDropInEffect = false;
                ele.ReleaseMouseCapture();
            }
        }


        void SymbolIconEmoji_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragDropInEffect)
            {
                double xPos = e.GetPosition(null).X - pos.X + ViewboxFloatingBar.Margin.Left;
                double yPos = e.GetPosition(null).Y - pos.Y + ViewboxFloatingBar.Margin.Top;
                SetFloatingBarPosition(xPos, yPos); // 唯一入口：坐标合法性校验，防拖飞
                pos = e.GetPosition(null);
            }
        }

        void SymbolIconEmoji_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //只处理左键：右键按下交由 ContextMenu 处理，若不过滤会误入拖动模式
            //（右键的 MouseUp 被 ContextMenu 弹出吞掉，isDragDropInEffect 将卡在 true，悬浮条会跟着鼠标飞）
            if (e.ChangedButton != MouseButton.Left) return;

            isDragDropInEffect = true;
            pos = e.GetPosition(null);
            downPos = e.GetPosition(null);
            GridForFloatingBarDraging.Visibility = Visibility.Visible;

            //浮动条开始拖动时收起笔设置面板：面板按笔图标位置定位，
            //条一动面板就悬空错位，不如直接收起（重开时重新定位）
            ClosePenSettingsPanel();
            //选择方式面板同理（按选择图标位置定位）
            CloseSelectionModePanel();

            SymbolIconEmoji.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Emoji;
        }

        void SymbolIconEmoji_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;

            isDragDropInEffect = false;

            if (e is null || (downPos.X == e.GetPosition(null).X && downPos.Y == e.GetPosition(null).Y))
            {
                SetBorderFloatingBarMainControlsVisibility(!borderFloatingBarMainControlsVisibility);
            }

            GridForFloatingBarDraging.Visibility = Visibility.Collapsed;
            SymbolIconEmoji.Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Emoji2;
        }

        //初始 false + XAML ScaleX=0/Opacity=0：启动默认收起为单笑脸图标（单击展开）
        bool borderFloatingBarMainControlsVisibility = false;

        #region 悬浮条存活性防御（bug：右键菜单弹出时点左键，isDragDropInEffect 卡死导致 Margin 被拖出屏幕）

        /// <summary>
        /// 拖动悬浮条时设置 Margin 的唯一入口：坐标合法性校验。
        /// 悬浮条飞出屏幕工作区（含缓冲）即拒绝写入并复位拖动状态——运行中悬浮条绝不允许被移出屏幕。
        /// </summary>
        private void SetFloatingBarPosition(double x, double y)
        {
            double w = ViewboxFloatingBar.ActualWidth > 0 ? ViewboxFloatingBar.ActualWidth : 50;
            double h = ViewboxFloatingBar.ActualHeight > 0 ? ViewboxFloatingBar.ActualHeight : 50;
            double wa = SystemParameters.WorkArea.Width, ha = SystemParameters.WorkArea.Height;
            const double margin = 200; // 缓冲：允许稍微拖出边缘，但不允许完全飞出

            if (x < -w - margin || x > wa + margin || y < -h - margin || y > ha + margin)
            {
                // 非法位置：拒绝写入 + 复位拖动状态 + 拉回默认位，防 isDragDropInEffect 卡死连锁拖飞
                isDragDropInEffect = false;
                ViewboxFloatingBar.Margin = new Thickness((wa - 284) / 2, ha - 80, -2000, -200);
                LogHelper.WriteLogToFile($"[FloatBar] 拦截非法位置 ({x:F0},{y:F0})，已复位", LogHelper.LogType.Event);
                return;
            }
            ViewboxFloatingBar.Margin = new Thickness(x, y, -2000, -200);
        }

        /// <summary>自愈：拖动状态卡死检测。运行中悬浮条必须可见（本机制唯一例外是程序关闭）。</summary>
        private void InitFloatingBarWatchdog()
        {
            var timer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(500) };
            timer.Tick += (s, e) =>
            {
                // 1) 拖动状态卡死自愈：左键未按下却仍在拖动 → 强制复位（右键菜单吞 MouseUp 的卡死场景）
                if (isDragDropInEffect && System.Windows.Input.Mouse.LeftButton != MouseButtonState.Pressed)
                    isDragDropInEffect = false;

                // 2) 可见性自愈：运行中（非旧UI模式）悬浮条必须可见，除非程序正在关闭
                if (!CloseIsFromButton && !App.StartArgs.Contains("-o")
                    && ViewboxFloatingBar.Visibility != Visibility.Visible)
                {
                    LogHelper.WriteLogToFile("[FloatBar] 检测到悬浮条不可见，已自愈", LogHelper.LogType.Event);
                    ViewboxFloatingBar.Visibility = Visibility.Visible;
                }
            };
            timer.Start();
        }

        #endregion

        void SetBorderFloatingBarMainControlsVisibility(bool isVisible, bool isAnimated = true)
        {
            borderFloatingBarMainControlsVisibility = isVisible;
            if (!isVisible)
            {
                BorderFloatingBarMainControls.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(isAnimated ? 100 : 0))
                {
                    EasingFunction = new PowerEase() { Power = 4, EasingMode = EasingMode.EaseOut },
                });
                BorderFloatingBarMainControls.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(isAnimated ? 100 : 0)));

                //工具条收起成笑脸时，图形面板一并收起：面板已"解挂"到主窗口根层
                //（方案B），不再是工具条的后代，不会随祖先缩放自动消失——不补这行
                //会出现"工具条没了、面板还孤零零浮在屏幕上"的状态
                try { BorderDrawShape.Visibility = Visibility.Collapsed; } catch { }
                //笔设置面板同理（已解挂到 Main_Grid 顶层，不随工具条缩放消失）
                try { ClosePenSettingsPanel(); } catch { }
                //选择方式面板同理（也已解挂到 Main_Grid 顶层）
                try { CloseSelectionModePanel(); } catch { }
            }
            else
            {
                BorderFloatingBarMainControls.RenderTransform.BeginAnimation(ScaleTransform.ScaleXProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(isAnimated ? 160 : 0))
                {
                    EasingFunction = new PowerEase() { Power = 4, EasingMode = EasingMode.EaseOut },
                });
                BorderFloatingBarMainControls.BeginAnimation(OpacityProperty, new DoubleAnimation(1, TimeSpan.FromMilliseconds(isAnimated ? 160 : 0)));
            }
        }

        #endregion


        private void GridPPTControlPrevious_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlidesUp_Click(BtnPPTSlidesUp, null);
        }

        private void GridPPTControlNext_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlidesDown_Click(BtnPPTSlidesDown, null);
        }

        private void ImagePPTControlEnd_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnPPTSlideShowEnd_Click(BtnPPTSlideShowEnd, null);
        }

        #endregion
    }
}
