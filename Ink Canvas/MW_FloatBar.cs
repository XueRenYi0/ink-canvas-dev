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

        /// <summary>弹层关闭范围 —— 两个场景要关的东西不一样，别混用</summary>
        private enum PopupScope
        {
            /// <summary>切工具/换色/撤销等"正在操作画布"：只关会让路的子面板（更多 + 笔 + 选择 + 橡皮）。
            /// 图形面板刻意不关 —— 用户可能正钉着它连续画图。</summary>
            ToolSwitch,

            /// <summary>工具条整体收起：关掉所有"不随工具条缩放消失"的浮层（笔 + 选择 + 橡皮 + 图形）。</summary>
            ToolbarCollapse,
        }

        /// <summary>
        /// 统一弹层关闭入口 —— 全项目 7 个弹层的唯一登记处。
        /// ★ 今后新增弹层，先判断它属于下面哪一组，再决定要不要在本函数登记；
        ///   不要在调用点各自手写 Visibility = Collapsed（漏一个就静默出 bug ——
        ///   2026-09-11 就是"收起工具条时漏关橡皮设置面板"）。
        ///
        /// 【A 组】宿主在 BorderFloatingBarMainControls 子树内 → 工具条收起时随缩放动画一起消失，
        ///         不需要显式关；只有"切工具让路"时才主动关 BorderTools。
        ///   1 更多面板 BorderTools      （MainWindow.xaml:1490，注意用的是 Name= 不是 x:Name=）
        ///   2 截图菜单 BorderImageMenu  （MainWindow.xaml:1326）
        ///
        /// 【B 组】★不随工具条缩放消失，凡"收起工具条"必须显式关，漏一个就留在屏幕上：
        ///   3 笔设置   PenSettingsPanel    （MainWindow.xaml:1855，Main_Grid 直挂）
        ///   4 选择方式 SelectionModePanel  （MainWindow.xaml:1861，Main_Grid 直挂）
        ///   5 橡皮设置 EraserSettingsPanel （MainWindow.xaml:1867，Main_Grid 直挂）
        ///   6 图形绘制 BorderDrawShape     （XAML:1118，运行期 DetachShapePanelToRoot 搬到 Main_Grid）
        ///   7 数学面板 原生 COM 窗口（micaut），非 WPF 元素，永不受工具条缩放影响
        ///      ⚠ 已知缺口：本函数目前**不关它**。因为"收起工具条腾地方、继续用手写数学面板"
        ///        是合理用法，关掉会打断正在输入的公式，是否要关待定。
        /// </summary>
        private void ClosePopupLayers(PopupScope scope)
        {
            // A 组：只有"切工具让路"才关更多面板
            if (scope == PopupScope.ToolSwitch)
            {
                BorderTools.Visibility = Visibility.Collapsed;
            }

            // B 组：解挂在 Main_Grid 顶部（或运行期 Detach）的浮层
            ClosePenSettingsPanel(); // 笔设置面板（替代原 BorderPenWidth 小面板）
            CloseSelectionModePanel(); // 选择方式面板（矩形框选/自由选择）
            CloseEraserSettingsPanel(); // 橡皮设置面板（含滑动清屏，原 BorderClearInDelete 已废弃）

            if (scope == PopupScope.ToolbarCollapse)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        /// <summary>切工具 / 换色时收起挡路的子面板（= ClosePopupLayers(ToolSwitch)，行为与改动前一致）</summary>
        private void HideSubPanels()
        {
            ClosePopupLayers(PopupScope.ToolSwitch);
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

        // 工具选中高亮色：色值统一取自 InkSpec.xaml 的 HighlightBrush（#FF0088FF），
        // 不再在本文件硬编码（2026-09-11 收口，原为 Color.FromRgb(0,136,255) 的字面量）。
        // 解析失败时退回同色兜底，保证高亮不会因资源缺失而莫名消失。
        private static readonly Color ToolHighlightFallbackColor = Color.FromRgb(0, 136, 255);

        /// <summary>解析当前工具高亮主色（= InkSpec.xaml 的 HighlightBrush 颜色；解析不到时退回兜底蓝）</summary>
        private static Color ResolveToolHighlightColor()
        {
            var brush = Application.Current?.TryFindResource("HighlightBrush") as SolidColorBrush;
            return brush != null ? brush.Color : ToolHighlightFallbackColor;
        }

        /// <summary>按给定不透明度生成高亮蓝画刷（色值取自 HighlightBrush）。
        /// 各处"激活态底色 / 蓝边"统一走这里，避免同一色值在不同文件里各写一遍颜色字面量。</summary>
        private static Brush CreateHighlightBrush(byte alpha)
        {
            var c = ResolveToolHighlightColor();
            return new SolidColorBrush(Color.FromArgb(alpha, c.R, c.G, c.B));
        }

        /// <summary>工具选中高亮边框色（笔图标 / 选择图标 / 粗细档位的蓝边）</summary>
        private static Brush PenToolHighlightBrush
        {
            get { return CreateHighlightBrush(255); }
        }

        /// <summary>淡蓝选中底（约 15% 不透明度），用于笔图标高亮底和粗细档位选中底</summary>
        private static Brush PenToolHighlightSoftBrush
        {
            get { return CreateHighlightBrush(38); }
        }

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

        #region 快捷换色条（常驻浮动条：4 格 × 2 色，色值固定，不跟白板/黑板联动）

        /// <summary>
        /// 色槽表：4 格 × 2 槽 = 8 色（固定，不随板面变）。
        /// 每格 [0] = 大色区（色 A），[1] = 右下小角（色 B）。
        /// 索引语义与全局一致：0黑 1红 2绿 3蓝 4黄 5白 7品红 8青。
        /// 配对仍用标准互补对（黑/白 · 红/青 · 蓝/黄 · 绿/品红）只是为了好记；
        /// 两色是**平等**的：点哪个用哪个，不做"当前色 → 对色"的跳转。
        /// </summary>
        private static readonly int[,] QuickColorSlots = new int[4, 2]
        {
            { 0, 5 }, // 黑白
            { 1, 8 }, // 红青
            { 3, 4 }, // 蓝黄
            { 2, 7 }, // 绿品红
        };

        /// <summary>取色槽颜色：红/绿/蓝/黄与右面板色块同源，黑白青品红取常量/扩展色。
        /// 右面板色块已不再按板面换两套（见 SetColors），所以这里读到的色是稳定的。</summary>
        private Color GetQuickColorValue(int slot, int sub)
        {
            switch (QuickColorSlots[slot, sub])
            {
                case 0: return Colors.Black;
                case 1: return ((SolidColorBrush)BtnColorRed.Background).Color;
                case 2: return ((SolidColorBrush)BtnColorGreen.Background).Color;
                case 3: return ((SolidColorBrush)BtnColorBlue.Background).Color;
                case 4: return ((SolidColorBrush)BtnColorYellow.Background).Color;
                case 5: return StringToColor("#FFFEFEFE");
                case 7: return ExtraColors[1]; // 品红
                default: return ExtraColors[2]; // 8 = 青
            }
        }

        /// <summary>大色区单击：用大区显示的那个颜色（不跳色、不看板面）；任何工具状态下都切回画笔</summary>
        private void QuickColorMainDot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is FrameworkElement fe) || !int.TryParse(fe.Tag?.ToString(), out int slot)) return;
            if (slot < 0 || slot > 3) return;
            ApplyQuickColor(QuickColorSlots[slot, 0], GetQuickColorValue(slot, 0));
        }

        /// <summary>右下小角单击：用小角显示的那个颜色</summary>
        private void QuickColorMiniDot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!(sender is FrameworkElement fe) || !int.TryParse(fe.Tag?.ToString(), out int slot)) return;
            if (slot < 0 || slot > 3) return;
            ApplyQuickColor(QuickColorSlots[slot, 1], GetQuickColorValue(slot, 1));
        }

        /// <summary>
        /// 应用快捷色：设色 + 记索引 + 清橡皮意图 + 走 ColorSwitchCheck（与点色块完全同一条路径），
        /// 保证"右面板高亮 / 笔身颜色 / 条上选中环"三处同步；除颜色外不做任何别的事。
        /// </summary>
        private void ApplyQuickColor(int index, Color color)
        {
            inkCanvas.DefaultDrawingAttributes.Color = color;
            inkColor = index;
            forceEraser = false;
            ColorSwitchCheck();
        }

        /// <summary>当前笔色在本格的哪个槽：0=大色区 1=右下小角 -1=都不在（选中环共用）</summary>
        private int GetQuickColorSlotActive(int slot)
        {
            if (inkColor == QuickColorSlots[slot, 0]) return 0;
            if (inkColor == QuickColorSlots[slot, 1]) return 1;
            return -1;
        }

        /// <summary>
        /// 刷新快捷换色条：四格两色**恒定显示**（位置与色值都不随板面/当前色变），
        /// 只有"当前笔色落在哪一槽"用亮蓝环表达 ——
        /// 落在色 A → 大区外圈环；落在色 B → 小角外圈环；橙/自定义色不在四格内 → 四格都不亮环。
        /// 由 UpdatePenIconColor 搭车调用（换色/启动恢复的公共汇合点；函数名没提色条，改调用链时注意）。
        /// </summary>
        private void RefreshQuickColorStrip()
        {
            try
            {
                var mainDots = new[] { QuickColorDotBW, QuickColorDotRC, QuickColorDotBY, QuickColorDotGM };
                var miniDots = new[] { QuickColorMiniBW, QuickColorMiniRC, QuickColorMiniBY, QuickColorMiniGM };
                var mainRings = new[] { QuickColorRingBW, QuickColorRingRC, QuickColorRingBY, QuickColorRingGM };
                var miniRings = new[] { QuickColorMiniRingBW, QuickColorMiniRingRC, QuickColorMiniRingBY, QuickColorMiniRingGM };

                for (int slot = 0; slot < 4; slot++)
                {
                    mainDots[slot].Fill = new SolidColorBrush(GetQuickColorValue(slot, 0));
                    miniDots[slot].Fill = new SolidColorBrush(GetQuickColorValue(slot, 1));
                    miniDots[slot].Visibility = Visibility.Visible; // 小角常显，永不隐藏

                    int active = GetQuickColorSlotActive(slot);
                    mainRings[slot].Visibility = active == 0 ? Visibility.Visible : Visibility.Collapsed;
                    miniRings[slot].Visibility = active == 1 ? Visibility.Visible : Visibility.Collapsed;
                }
            }
            catch { }
        }

        #endregion

        #endregion

        private void SymbolIconUndo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BtnUndo_Click(BtnUndo, null);
            HideSubPanels();
        }

        private void SymbolIconRedo_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

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

        /// <summary>
        /// 选择图标单击（两段式，2026-09-11 改）：
        /// 从别的工具切进来 → 只进入选择模式，**不弹面板**；
        /// 已经在选择模式时再点图标 → 才弹/收选择方式面板（矩形框选/自由选择/全选）。
        /// 双击（500ms 内两击）仍为全选当前页墨迹。
        /// </summary>
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

            // 两段式：面板只在"点击当下已经是选择模式"时才弹/收。
            // 已在选择模式 → 再点图标 = 调选择方式（弹/收面板，ToggleSelectionModePanel 自带 toggle）；
            // 从别的工具切进来 → 只进入选择模式，面板不弹（切工具是为了框选，不是为了调方式；
            // 弹出面板会挡住板书，还得多点一次才能收起）。
            // 顺带修掉旧实现的视觉瑕疵：双击全选时第一击会先弹面板、第二击才关掉 → 面板闪一下；
            // 两段式下第一击不再弹面板，双击路径变干净。
            if (inkCanvas.EditingMode == InkCanvasEditingMode.Select)
            {
                ToggleSelectionModePanel();
                return;
            }

            BtnSelect_Click(BtnSelect, null);

            ImageEraser.Visibility = Visibility.Visible;
            ViewboxBtnColorBlackContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorBlueContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorGreenContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorRedContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));
            ViewboxBtnColorYellowContent.BeginAnimation(OpacityProperty, new DoubleAnimation(0, TimeSpan.FromMilliseconds(ColorSwiftOpacityDurationOff)));

            // HideSubPanels 内含 CloseSelectionModePanel：切进来时不残留任何面板。
            // （选择方式面板改由"已在选择模式时再点图标"打开，见上方两段式注释）
            HideSubPanels();
        }

        private void SymbolIconScreenshot_MouseUp(object sender, MouseButtonEventArgs e)
        {
            BtnScreenshot_Click(BtnScreenshot, null);
        }

        Point pointDesktop = new Point(-1, -1); //用于记录上次进入PPT或白板时的坐标
        Point pointPPT = new Point(-1, -1); //用于记录上次在PPT中打开白板时的坐标

        private void ImageBlackboard_MouseUp(object sender, MouseButtonEventArgs e)
        {
            // 防误触：按下与松开须在同一元素。
            // sender == null 是"代码直调"（MW_FloatBar.cs 里光标图标收起白板时调用），此时跳过守卫。
            if (sender != null && lastBorderMouseDownObject != sender) return;

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
                // 不再按板面自动改笔色（2026-09-12 改）：
                // 原"白板→黑笔 / 黑板→白笔"会让用户刚选的色被悄悄换掉，
                // 现在进板面前后保持用户当前笔色，需要换色自己点色块/色点。
                // （原来这里顺带收起子面板，改色逻辑删掉后单独保留这个副作用）
                HideSubPanels();
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
                // 关黑板不再强制红笔（2026-09-12）：保持用户当前笔色（原来是 BorderPenColorRed_MouseUp）
                HideSubPanels();
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
        /// 橡皮图标单击（两段式，2026-09-11 改）：
        /// 从别的工具切进来 → 只进入橡皮，**不弹面板**；
        /// 已经在橡皮时再点图标 → 才弹/收橡皮设置面板。
        /// 擦除方式/大小只在面板里切（原"点图标 toggle 擦除方式"已废弃——不可见、易误切）。
        /// </summary>
        private void ImageEraser_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            // 点击当下是否已在橡皮（BtnErase_Click 会按记忆的方式二选一设置 EditingMode：
            // 面积擦 EraseByPoint / 笔画擦 EraseByStroke）
            bool alreadyEraser = inkCanvas.EditingMode == InkCanvasEditingMode.EraseByPoint
                              || inkCanvas.EditingMode == InkCanvasEditingMode.EraseByStroke;

            // 进入橡皮模式（BtnErase_Click 已改为固定进入当前方式，不 toggle）
            BtnErase_Click(BtnErase, e);

            if (alreadyEraser)
            {
                // 已在橡皮 → 再点图标 = 调擦除方式/大小（面板显隐 + 定位，MW_EraserSettings.cs）
                ToggleEraserSettingsPanel();
            }
            else
            {
                // 刚从别的工具切进来 → 不弹面板，并兜底关闭（避免任何残留面板挡住板书）
                CloseEraserSettingsPanel();
            }
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
            if (lastBorderMouseDownObject != sender) return;

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
                // 手指模式图标跟随真实状态刷新（它可能被别处改过：设置页开关、启动时
                // MW_DefinitionsLoading 恢复、设置面板自动进入手指模式的联动）
                UpdateFingerModeIcon();
                // 定位：以三滑杆图标为锚做绝对定位（解挂到 Main_Grid 后，不随悬浮条缩放漂移）
                PositionBorderTools();
            }
        }

        /// <summary>
        /// 手指模式（「更多」面板内）：点击整格切换。
        /// 必须经 ToggleSwitchModeFinger.IsOn 走 —— 它是这条设置的**唯一落盘通道**，
        /// 面板里原来的 ToggleSwitch 只是靠 IsOn 双向绑定挂在它上面。改了它的 IsOn
        /// 就会触发 MW_RightPanel.cs 的 ToggleSwitchModeFinger_Toggled（同步
        /// 「自动进入手指模式」+ 保存设置），绕开它则设置存不下来。
        /// </summary>
        private void SymbolIconFingerMode_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ToggleSwitchModeFinger.IsOn = !ToggleSwitchModeFinger.IsOn;
            UpdateFingerModeIcon();
        }

        /// <summary>
        /// 手指模式图标状态：关闭=显示斜杠+手指变淡；开启=隐藏斜杠+手指正常。
        /// 与「双指手势」锁的 UpdateGestureLockIcon 同款表达（斜杠 = 该功能未启用）。
        /// </summary>
        private void UpdateFingerModeIcon()
        {
            bool on = ToggleSwitchModeFinger.IsOn;
            PathFingerModeSlash.Visibility = on ? Visibility.Collapsed : Visibility.Visible;
            PathFingerSingle.Opacity = on ? 1.0 : 0.45;
        }

        /// <summary>
        /// 把「更多」面板从悬浮条视觉树解挂到主窗口根层（启动时调用一次，与 DetachShapePanelToRoot 同款）。
        /// 为什么必须解挂：面板原先是 BorderFloatingBarMainControls（缩放容器）内的子元素，
        /// 位置只能用"相对 20×24 图标格"的负 Margin 表达 —— 既做不到居中，
        /// 又会随工具条收起时的 ScaleX 缩放一起变形；解挂到 Main_Grid 后与笔/橡皮/选择三面板
        /// 归入同一套绝对定位体系（位置可控、不受缩放影响）。
        /// 失败时面板留在原处（功能不受影响，只是位置不理想）。
        /// </summary>
        private void DetachToolsPanelToRoot()
        {
            try
            {
                var oldParent = BorderTools.Parent as System.Windows.Controls.Grid;
                if (oldParent == null || oldParent == Main_Grid) return; // 已在根层，防重入
                oldParent.Children.Remove(BorderTools);
                Main_Grid.Children.Add(BorderTools); // 追加到最后 = z 序最高，盖在画布之上
                BorderTools.HorizontalAlignment = HorizontalAlignment.Left;
                BorderTools.VerticalAlignment = VerticalAlignment.Top;
            }
            catch (Exception ex)
            {
                Ink_Canvas.Helpers.LogHelper.WriteLogToFile("[更多面板] 解挂失败 " + ex, Ink_Canvas.Helpers.LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// 「更多」面板定位：与 PositionPenSettingsPanel 完全同款算法 ——
        /// 以三滑杆图标为锚，图标上方空间够就弹在图标正上方，顶不够时弹到图标下方；
        /// 水平居中对齐图标，并夹在屏幕（Main_Grid）内。
        /// </summary>
        private void PositionBorderTools()
        {
            // 先量出面板实际尺寸（刚设为 Visible，强制布局立即生效）
            Main_Grid.UpdateLayout();
            double panelW = BorderTools.ActualWidth;
            double panelH = BorderTools.ActualHeight;
            if (panelW < 1 || panelH < 1) return; // 布局未就绪，放弃定位（保持上次位置）

            // 三滑杆图标在 Main_Grid 坐标系中的位置与渲染尺寸
            var transform = GridToolsEntry.TransformToAncestor(Main_Grid);
            Point iconTopLeft = transform.Transform(new Point(0, 0));
            Point iconBottomRight = transform.Transform(new Point(GridToolsEntry.ActualWidth, GridToolsEntry.ActualHeight));
            double iconCenterX = (iconTopLeft.X + iconBottomRight.X) / 2;
            double iconTopY = iconTopLeft.Y;
            double iconBottomY = iconBottomRight.Y;

            const double Gap = 8; // 面板与悬浮条间距（与笔面板一致）
            double gridW = Main_Grid.ActualWidth;
            double gridH = Main_Grid.ActualHeight;

            // 垂直：图标上方放得下就放上方（贴近手、不挡下方板书），否则放图标下方
            double y = iconTopY - panelH - Gap;
            if (y < Gap) y = iconBottomY + Gap;
            if (y + panelH > gridH - Gap) y = Math.Max(Gap, gridH - panelH - Gap);

            // 水平：中心对齐图标，夹在屏幕内
            double x = iconCenterX - panelW / 2;
            if (x < Gap) x = Gap;
            if (x + panelW > gridW - Gap) x = Math.Max(Gap, gridW - panelW - Gap);

            BorderTools.Margin = new Thickness(x, y, 0, 0);
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
            if (lastBorderMouseDownObject != sender) return;

            ShowWhiteboardPatternMenu(sender as FrameworkElement, placeBelow: false);
        }

        /// <summary>更多面板·打开设置</summary>
        private void MoreToolsSettings_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            BtnSettings_Click(BtnSettings, null);
        }

        /// <summary>更多面板·检查更新（手动检查：已是最新版本时会给出提示）</summary>
        private void MoreToolsCheckUpdate_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            (Application.Current as App)?.CheckForUpdate(true);
        }

        /// <summary>更多面板·白板 / 黑板切换（从白板底纹菜单顶部拆出来的独立入口；
        /// 复用 BtnSwitchTheme_Click 的全套联动：板面色、UI 深浅主题、白板/黑板文案）</summary>
        private void MoreToolsBoardTheme_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BtnSwitchTheme_Click(null, null);
        }

        /// <summary>更多面板·查看快捷键（原笑脸右键菜单 →「快捷键 → 查看快捷键」）</summary>
        private void MoreToolsShowShortcuts_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            BorderTools.Visibility = Visibility.Collapsed;
            MenuItemShowShortcuts_Click(null, null);
        }

        /// <summary>更多面板·退出（原主栏退出图标 + 笑脸右键菜单「退出」，内部有二次确认）</summary>
        private void MoreToolsExit_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

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

                //工具条收起：关掉所有"不随工具条缩放消失"的浮层（笔/选择/橡皮/图形）。
                //它们已"解挂"到主窗口根层（方案B）或运行期 Detach，不再是工具条的后代，
                //不会随祖先缩放自动消失——不关就会出现"工具条没了、面板还孤零零浮在屏幕上"。
                //★ 统一走 ClosePopupLayers，新增浮层只需在其中登记一处
                //  （2026-09-11 起因：此处曾漏关橡皮设置面板）
                try { ClosePopupLayers(PopupScope.ToolbarCollapse); } catch { }
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
            if (lastBorderMouseDownObject != sender) return;

            BtnPPTSlideShowEnd_Click(BtnPPTSlideShowEnd, null);
        }

        #endregion
    }
}
