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
using System.Windows.Data;
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
    /// <summary>MainWindow 分部类：图形绘制（形状按钮）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Shape Drawing

        #region Floating Bar Control

        private void ImageDrawShape_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (BorderDrawShape.Visibility == Visibility.Visible)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
            else
            {
                ShowShapePanel(); //统一入口：先定位（会话记忆或图标上方）再显示
            }
        }

        #endregion Floating Bar Control

        #region 解挂的图形面板：自由拖动小窗口（方案B）

        // ==================== 原理 ====================
        // 面板原本是悬浮条视觉树里的后代（负 Margin 吊在图形图标上方），
        // 启动时把它移到主窗口根层 Main_Grid，改用 Margin 绝对定位：
        //   · 主窗口本身全屏 → 面板可在整屏自由拖动，效果等同"独立小窗口"
        //   · 仍在同一个 Window 的命名域 → FindName 高亮、26 个按钮事件、
        //     14 处显隐控制全部原样有效，零逻辑改动
        // 位置记忆与数学面板同策略：会话级——本次运行拖到哪就停在哪，重启回默认。

        /// <summary>会话级位置记忆：拖动后面板停在哪，下次打开就在哪（软件重启自动回默认）</summary>
        private Point? _shapePanelSessionPos = null;

        /// <summary>会话级大小记忆：拖角柄后面板多宽多高，下次打开就保持（软件重启回默认）</summary>
        private Size? _shapePanelSessionSize = null;

        //面板默认尺寸 / 限制（角柄：宽=每行图标数，高=图标区可视行数）
        private const double ShapePanelDefaultW = 420;
        private const double ShapePanelDefaultH = 400;
        private const double ShapePanelMinW = 340;
        private const double ShapePanelMinH = 280;

        private bool _isDraggingShapePanel = false;      // 是否正在拖动面板
        private Point _shapePanelDragStartMousePos;      // 拖动开始时鼠标在主窗口的坐标
        private Point _shapePanelDragStartMargin;        // 拖动开始时面板 Margin 的 (Left, Top)

        private bool _isResizingShapePanel = false;      // 是否正在调整面板大小
        private Point _resizeStartMousePos;              // 调整开始时鼠标位置
        private Size _resizeStartSize;                   // 调整开始时面板尺寸
        //图标区默认可视高度（48 格 + 6 间距 = 每行 54；4 行 = 216）
        private const double ShapeIconScrollDefaultMax = 216;
        //图库区默认可视高度（与 XAML 中 LibraryScroll 的 MaxHeight="88" 对应）
        private const double LibraryScrollDefaultMax = 88;

        /// <summary>
        /// 把图形面板从悬浮条视觉树"解挂"到主窗口根层（构造函数里调用一次）。
        /// 失败时面板留在原处继续用旧的吊挂方式，功能不受影响（优雅降级）。
        /// </summary>
        private void DetachShapePanelToRoot()
        {
            try
            {
                var oldParent = BorderDrawShape.Parent as System.Windows.Controls.Grid;
                if (oldParent == null || oldParent == Main_Grid) return; //已在根层，防重入
                oldParent.Children.Remove(BorderDrawShape);
                Main_Grid.Children.Add(BorderDrawShape); //追加到最后 = z 序最高，盖在画布之上
                //从"负 Margin 吊在图标上方"改为根层左上角绝对定位
                BorderDrawShape.HorizontalAlignment = HorizontalAlignment.Left;
                BorderDrawShape.VerticalAlignment = VerticalAlignment.Top;
            }
            catch (Exception ex)
            {
                Ink_Canvas.Helpers.LogHelper.WriteLogToFile("[图形面板] 解挂失败 " + ex, Ink_Canvas.Helpers.LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// 把面板相对默认高度多出的部分分配给两个内部滚动区：
        /// 图标区拿 60%（内置图形多、查找频率高），图库区拿 40%（存的自定义图形也要能多看几行）。
        /// 之前只喂图标区，图库固定 88px——面板拉得再高图库也只能上下滚动，不符合直觉。
        /// </summary>
        private void ApplyShapePanelScrollHeights(double panelHeight)
        {
            double extra = Math.Max(0, panelHeight - ShapePanelDefaultH);
            ShapeIconScroll.MaxHeight = ShapeIconScrollDefaultMax + extra * 0.6;
            LibraryScroll.MaxHeight = LibraryScrollDefaultMax + extra * 0.4;
        }

        /// <summary>显示图形面板的统一入口：先恢复大小（会话记忆 → 默认值）和位置再显示</summary>
        private void ShowShapePanel()
        {
            //大小记忆：本次运行拉到多宽多高就保持（高度喂给图标区/图库区滚动 → 可视行数）
            if (_shapePanelSessionSize.HasValue)
            {
                BorderDrawShape.Width = _shapePanelSessionSize.Value.Width;
                BorderDrawShape.Height = _shapePanelSessionSize.Value.Height;
                ApplyShapePanelScrollHeights(_shapePanelSessionSize.Value.Height);
            }
            else
            {
                BorderDrawShape.Width = ShapePanelDefaultW;
                BorderDrawShape.Height = ShapePanelDefaultH;
                ApplyShapePanelScrollHeights(ShapePanelDefaultH);
            }
            ApplyShapePanelPosition();
            BorderDrawShape.Visibility = Visibility.Visible;
        }

        /// <summary>
        /// 计算并应用面板位置。优先级：会话记忆（本次运行拖过）→ 默认在图形图标正上方。
        /// 原来的 -315 偏移会让面板底部伸到图标下方 85px（面板高 400），和悬浮条叠在一起；
        /// 现改为面板完全悬在图标上方（底部留 12px 间距），悬浮条拖到哪都不重叠。
        /// </summary>
        private void ApplyShapePanelPosition()
        {
            if (_shapePanelSessionPos.HasValue)
            {
                SetShapePanelPosition(_shapePanelSessionPos.Value.X, _shapePanelSessionPos.Value.Y);
                return;
            }
            try
            {
                //图形图标当前屏幕位置 → 主窗口根层坐标（PointToScreen/FromScreen 自带全部变换换算，
                //悬浮条无论拖到哪、缩放多少，算出来的都是准确位置）
                var iconScreen = ImageDrawShape.PointToScreen(new Point(0, 0));
                var iconLocal = Main_Grid.PointFromScreen(iconScreen);
                //面板高度用代码设置的显式值（ShowShapePanel 里已先设好），不用 ActualHeight——
                //首次显示时布局未跑完 ActualHeight 可能为 0
                double h = BorderDrawShape.Height > 0 ? BorderDrawShape.Height : ShapePanelDefaultH;
                SetShapePanelPosition(iconLocal.X - 100, iconLocal.Y - h - 12);
            }
            catch
            {
                SetShapePanelPosition(200, 150); //拿不到图标坐标时的安全默认值
            }
        }

        /// <summary>设置面板位置的唯一入口（带越界钳制，面板绝不允许被拖飞出屏幕）</summary>
        private void SetShapePanelPosition(double x, double y)
        {
            double w = BorderDrawShape.ActualWidth > 0 ? BorderDrawShape.ActualWidth : 400;
            double h = BorderDrawShape.ActualHeight > 0 ? BorderDrawShape.ActualHeight : 290;
            double wa = SystemParameters.WorkArea.Width, ha = SystemParameters.WorkArea.Height;
            const double keep = 60; //面板至少留 60px 在屏幕内，保证随时能抓回
            if (x < -w + keep) x = -w + keep;
            if (y < 0) y = 0;
            if (x > wa - keep) x = wa - keep;
            if (y > ha - keep) y = ha - keep;
            BorderDrawShape.Margin = new Thickness(x, y, 0, 0);
        }

        // ---------- 标题行拖动（三件套：Down 记起点抓鼠标 / Move 增量平移 / Up 释放+记忆） ----------

        private void ShapePanelTitle_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;

            //双击标题栏 = 复位（Grid 没有 MouseDoubleClick 事件，用 ClickCount 手动判定）
            if (e.ClickCount == 2)
            {
                ShapePanelTitle_DoubleClick(sender, e);
                return;
            }

            _isDraggingShapePanel = true;
            _shapePanelDragStartMousePos = e.GetPosition(Main_Grid);
            _shapePanelDragStartMargin = new Point(BorderDrawShape.Margin.Left, BorderDrawShape.Margin.Top);
            ((UIElement)sender).CaptureMouse(); //抓住鼠标：移出标题区也持续跟随
        }

        private void ShapePanelTitle_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isDraggingShapePanel) return;
            var pos = e.GetPosition(Main_Grid);
            SetShapePanelPosition(
                _shapePanelDragStartMargin.X + pos.X - _shapePanelDragStartMousePos.X,
                _shapePanelDragStartMargin.Y + pos.Y - _shapePanelDragStartMousePos.Y);
        }

        private void ShapePanelTitle_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isDraggingShapePanel) return;
            _isDraggingShapePanel = false;
            try { ((UIElement)sender).ReleaseMouseCapture(); } catch { }
            //拖动结束 → 记住位置（会话级；软件重启后自动回默认）
            _shapePanelSessionPos = new Point(BorderDrawShape.Margin.Left, BorderDrawShape.Margin.Top);
        }

        /// <summary>双击标题栏 = 复位：清掉会话记忆，位置和大小都回默认（面板拖飞后的逃生门）</summary>
        private void ShapePanelTitle_DoubleClick(object sender, MouseButtonEventArgs e)
        {
            _shapePanelSessionPos = null;
            _shapePanelSessionSize = null;
            BorderDrawShape.Width = ShapePanelDefaultW;
            BorderDrawShape.Height = ShapePanelDefaultH;
            ApplyShapePanelScrollHeights(ShapePanelDefaultH); //两个滚动区一并回默认
            ApplyShapePanelPosition(); //回到图形图标正上方的默认位置
            ShowToastNotification("面板已复位到默认位置和大小");
        }

        // ---------- 右下角角柄：调整面板大小 ----------
        // 宽 → WrapPanel 换行（每行图标数）；高 → 图标区 ScrollViewer 可视高度（行数）

        private void ShapePanelResizeGrip_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            _isResizingShapePanel = true;
            _resizeStartMousePos = e.GetPosition(Main_Grid);
            _resizeStartSize = new Size(BorderDrawShape.ActualWidth, BorderDrawShape.ActualHeight);
            ((UIElement)sender).CaptureMouse();
        }

        private void ShapePanelResizeGrip_MouseMove(object sender, MouseEventArgs e)
        {
            if (!_isResizingShapePanel) return;
            var pos = e.GetPosition(Main_Grid);
            double w = _resizeStartSize.Width + pos.X - _resizeStartMousePos.X;
            double h = _resizeStartSize.Height + pos.Y - _resizeStartMousePos.Y;
            //钳制范围：宽最小保证每行至少 6 个图标，高最小保证标题+一行图标+通栏条可见
            double maxW = SystemParameters.WorkArea.Width * 0.9;
            double maxH = SystemParameters.WorkArea.Height * 0.9;
            BorderDrawShape.Width = Math.Max(ShapePanelMinW, Math.Min(maxW, w));
            BorderDrawShape.Height = Math.Max(ShapePanelMinH, Math.Min(maxH, h));
            //多出的高度实时分给图标区和图库区（即时反馈，不用松手才看到）
            ApplyShapePanelScrollHeights(BorderDrawShape.Height);
        }

        private void ShapePanelResizeGrip_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_isResizingShapePanel) return;
            _isResizingShapePanel = false;
            try { ((UIElement)sender).ReleaseMouseCapture(); } catch { }
            //调整结束 → 记住大小（会话级；软件重启后自动回默认）
            _shapePanelSessionSize = new Size(BorderDrawShape.ActualWidth, BorderDrawShape.ActualHeight);
        }

        #endregion

        int drawingShapeMode = 0;
        //长按锁定功能已移除（克隆重做已覆盖连续绘制需求）

        #region Buttons

        /// <summary>图钉专用按下处理：记录 + 拦截冒泡。
        /// 不拦的话事件会冒泡到标题栏拖动逻辑，被 CaptureMouse 抢走 → 图钉自己的 MouseUp 收不到 → 左键点不中。
        /// 触摸屏的触摸事件会提升为左键鼠标事件，这里拦住后三路输入（左键/右键/触摸）都能正常点图钉。</summary>
        private void SymbolIconPinDrawShape_MouseDown(object sender, MouseButtonEventArgs e)
        {
            lastBorderMouseDownObject = sender;
            e.Handled = true; //关键：不让标题栏收到这次按下（否则会被当成"开始拖动面板"）
        }

        private void SymbolIconPinBorderDrawShape_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ToggleSwitchDrawShapeBorderAutoHide.IsOn = !ToggleSwitchDrawShapeBorderAutoHide.IsOn;

            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                ((iNKORE.UI.WPF.Modern.Controls.SymbolIcon)sender).Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.Pin;
            }
            else
            {
                ((iNKORE.UI.WPF.Modern.Controls.SymbolIcon)sender).Symbol = iNKORE.UI.WPF.Modern.Controls.Symbol.UnPin;
            }
        }

        object lastMouseDownSender = null;
        DateTime lastMouseDownTime = DateTime.MinValue;

        //长按锁定功能已移除（克隆重做已覆盖连续绘制需求，所有图形统一"点一次画一次"）。
        //此方法保留为空实现：XAML 中 5 个图标仍绑定了 MouseDown 事件，删除绑定需改 XAML，留空壳成本最低。
        private void Image_MouseDown(object sender, MouseButtonEventArgs e)
        {
            //记录按下对象供 MouseUp 处理器判定"同一次点击"（Touch 也走这里）
            lastMouseDownSender = sender;
            lastMouseDownTime = DateTime.Now;
        }

        private void BtnPen_Click(object sender, RoutedEventArgs e)
        {
            forceEraser = false;
            drawingShapeMode = 0;
            inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight(); //切回笔模式时熄灭所有图形图标高亮
        }

        // ============ 图标三态视觉（悬停 / 激活） ============
        // 状态设计：
        //   ① 悬停态：鼠标（或手指按下瞬间，触摸会提升为鼠标事件）划过图标 → 浅灰底，提示"这个可以点"
        //   ② 激活态：点击选中工具（或长按锁定）→ 蓝色高亮，提示"现在用的是这个工具"
        // 实现方式（重要，新手注意）：
        //   三个状态全部由一个 Style 的触发器管理，代码只负责打 Tag="Active" 标记。
        //   不能在代码里直接改 Background——本地值优先级高于样式触发器，会把悬停效果"焊死"。
        private void InitShapeIconStyles()
        {
            //① 先造样式：默认透明 → 悬停浅灰 → 激活蓝（触发器按顺序叠加，后面的赢）
            var style = new Style(typeof(Border));
            style.Setters.Add(new Setter(Border.BackgroundProperty, Brushes.Transparent)); //默认：透明

            //悬停触发器：鼠标移到图标范围内时生效，离开自动还原（WPF 自动管理，无需代码）
            var hoverTrigger = new Trigger
            {
                Property = UIElement.IsMouseOverProperty, //Border 自身（含子元素 Image）的悬停判断
                Value = true
            };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(45, 128, 128, 128)))); //半透明中灰，深浅主题下都可见
            style.Triggers.Add(hoverTrigger);

            //激活触发器：代码给 Border.Tag 赋 "Active" 字符串时生效，清空 Tag 则还原
            var activeTrigger = new DataTrigger
            {
                Binding = new Binding("Tag") { RelativeSource = new RelativeSource(RelativeSourceMode.Self) },
                Value = "Active"
            };
            activeTrigger.Setters.Add(new Setter(Border.BackgroundProperty,
                new SolidColorBrush(Color.FromArgb(140, 0, 136, 255)))); //半透明蓝，与整体高亮风格一致
            style.Triggers.Add(activeTrigger);

            //② 把样式套到全部图标 Border 上（名字规则见 UpdateShapeIconHighlight）
            //范围说明：1~26 是拖拽绘制工具（激活态由 drawingShapeMode 决定）；
            //27 是 fx 函数识别按钮——它不是绘制模式，激活蓝由"面板是否打开"单独控制
            //（见 ShowMathPanel/CloseMathPanel），这里只负责给它悬停灰。
            for (int mode = 1; mode <= 27; mode++)
            {
                var border = FindName("BorderShapeIcon_" + mode) as Border;
                if (border == null) continue;
                //关键修复：XAML 里写了 Background="Transparent"（本地值），
                //本地值优先级高于样式触发器，不清掉会把悬停/激活颜色全部盖住（永远透明）。
                //ClearValue 删除本地值后，样式里的"默认 Transparent"会接管，视觉效果不变，
                //但触发器从此能正常改色。
                border.ClearValue(Border.BackgroundProperty);
                border.Style = style;
            }
        }

        private void UpdateShapeIconHighlight()
        {
            //图标 Border 名 = "BorderShapeIcon_" + 图形编号（drawingShapeMode 的值）
            //只打/清 Tag 标记，颜色交给样式的 DataTrigger 处理（见 InitShapeIconStyles）

            //搭车收口：本方法被全部工具激活点（30 处图形按钮 + 笔 + 选择工具）调用，
            //在这里统一结束"插入后一次性选中"，保证激活任何工具时画布遮罩立即收起。
            //（否则选区遮罩会吃掉第一笔的鼠标事件，图形画不出来）
            EndOneShotSelectionNow();

            int activeMode = drawingShapeMode;
            //只遍历 1~26（绘制工具）。27 号 fx 按钮的激活蓝表示"面板开着"，
            //与 drawingShapeMode 无关，由 ShowMathPanel/CloseMathPanel 单独管理，这里不碰它
            for (int mode = 1; mode <= 26; mode++)
            {
                var border = FindName("BorderShapeIcon_" + mode) as Border;
                if (border == null) continue; //该编号没有对应图标
                border.Tag = (mode == activeMode) ? "Active" : null;
            }
        }

        private void BtnDrawLine_Click(object sender, EventArgs e)
        {
            if (lastMouseDownSender == sender)
            {
                forceEraser = true;
                drawingShapeMode = 1;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = true;
                CancelSingleFingerDragMode();
                UpdateShapeIconHighlight();
            }
            lastMouseDownSender = null;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnDrawDashedLine_Click(object sender, EventArgs e)
        {
            if (lastMouseDownSender == sender)
            {
                forceEraser = true;
                drawingShapeMode = 8;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = true;
                CancelSingleFingerDragMode();
                UpdateShapeIconHighlight();
            }
            lastMouseDownSender = null;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnDrawDotLine_Click(object sender, EventArgs e)
        {
            if (lastMouseDownSender == sender)
            {
                forceEraser = true;
                drawingShapeMode = 18;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = true;
                CancelSingleFingerDragMode();
                UpdateShapeIconHighlight();
            }
            lastMouseDownSender = null;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnDrawArrow_Click(object sender, EventArgs e)
        {
            if (lastMouseDownSender == sender)
            {
                forceEraser = true;
                drawingShapeMode = 2;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = true;
                CancelSingleFingerDragMode();
                UpdateShapeIconHighlight();
            }
            lastMouseDownSender = null;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnDrawParallelLine_Click(object sender, EventArgs e)
        {
            if (lastMouseDownSender == sender)
            {
                forceEraser = true;
                drawingShapeMode = 15;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = true;
                CancelSingleFingerDragMode();
                UpdateShapeIconHighlight();
            }
            lastMouseDownSender = null;
            if (ToggleSwitchDrawShapeBorderAutoHide.IsOn)
            {
                BorderDrawShape.Visibility = Visibility.Collapsed;
            }
        }

        private void BtnDrawCoordinate1_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 11;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCoordinate2_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 12;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCoordinate3_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 13;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCoordinate4_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 14;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCoordinate5_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 17;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawRectangle_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 3;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawRectangleCenter_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 19;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawEllipse_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 4;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCircle_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 5;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCenterEllipse_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 16;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCenterEllipseWithFocalPoint_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 23;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawDashedCircle_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 10;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawHyperbola_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 24;
            drawMultiStepShapeCurrentStep = 0;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawHyperbolaWithFocalPoint_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 25;
            drawMultiStepShapeCurrentStep = 0;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawParabola1_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 20;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawParabolaWithFocalPoint_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 22;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawParabola2_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 21;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        // 正弦曲线按钮：进入"画 sin"模式（drawingShapeMode = 26）
        private void BtnDrawSin_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 26;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCylinder_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 6;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCone_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 7;
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        private void BtnDrawCuboid_Click(object sender, EventArgs e)
        {
            forceEraser = true;
            drawingShapeMode = 9;
            isFirstTouchCuboid = true;
            CuboidFrontRectIniP = new Point();
            CuboidFrontRectEndP = new Point();
            inkCanvas.EditingMode = InkCanvasEditingMode.None;
            inkCanvas.IsManipulationEnabled = true;
            CancelSingleFingerDragMode();
            UpdateShapeIconHighlight();
        }

        #endregion

        private void inkCanvas_TouchMove(object sender, TouchEventArgs e)
        {
            if (isSingleFingerDragMode) return;
            if (drawingShapeMode != 0)
            {
                if (isLastTouchEraser)
                {
                    return;
                }
                //EraserContainer.Background = null;
                ImageEraser.Visibility = Visibility.Visible;
                if (isWaitUntilNextTouchDown) return;
                if (dec.Count > 1)
                {
                    isWaitUntilNextTouchDown = true;
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                        inkCanvas.Strokes.Remove(lastTempCenterDotStroke); //多指取消绘制：中心圆点也要一并清掉
                    }
                    catch { }
                    return;
                }
                if (inkCanvas.EditingMode != InkCanvasEditingMode.None)
                {
                    inkCanvas.EditingMode = InkCanvasEditingMode.None;
                }
            }
            MouseTouchMove(e.GetTouchPoint(inkCanvas).Position);
        }

        int drawMultiStepShapeCurrentStep = 0; //多笔完成的图形 当前所处在的笔画
        StrokeCollection drawMultiStepShapeSpecialStrokeCollection = new StrokeCollection(); //多笔完成的图形 当前所处在的笔画
        //double drawMultiStepShapeSpecialParameter1 = 0.0; //多笔完成的图形 特殊参数 通常用于表示a
        //double drawMultiStepShapeSpecialParameter2 = 0.0; //多笔完成的图形 特殊参数 通常用于表示b
        double drawMultiStepShapeSpecialParameter3 = 0.0; //多笔完成的图形 特殊参数 通常用于表示k

        private void MouseTouchMove(Point endP)
        {
            List<System.Windows.Point> pointList;
            StylusPointCollection point;
            Stroke stroke;
            StrokeCollection strokes = new StrokeCollection();
            Point newIniP = iniP;
            switch (drawingShapeMode)
            {
                case 1:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point(iniP.X, iniP.Y),
                        new System.Windows.Point(endP.X, endP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    inkCanvas.Strokes.Add(stroke);
                    break;
                case 8:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    strokes.Add(GenerateDashedLineStrokeCollection(iniP, endP));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 18:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    strokes.Add(GenerateDotLineStrokeCollection(iniP, endP));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 2:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    double w = 30, h = 10;
                    double theta = Math.Atan2(iniP.Y - endP.Y, iniP.X - endP.X);
                    double sint = Math.Sin(theta);
                    double cost = Math.Cos(theta);

                    pointList = new List<Point>
                    {
                        new Point(iniP.X, iniP.Y),
                        new Point(endP.X , endP.Y),
                        new Point(endP.X + (w * cost - h * sint), endP.Y + (w * sint + h * cost)),
                        new Point(endP.X,endP.Y),
                        new Point(endP.X + (w * cost + h * sint), endP.Y - (h * cost - w * sint))
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    inkCanvas.Strokes.Add(stroke);
                    break;
                case 15:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    double d = GetDistance(iniP, endP);
                    if (d == 0) return;
                    double sinTheta = (iniP.Y - endP.Y) / d;
                    double cosTheta = (endP.X - iniP.X) / d;
                    double tanTheta = Math.Abs(sinTheta / cosTheta);
                    double x = 25;
                    if (Math.Abs(tanTheta) < 1.0 / 12)
                    {
                        sinTheta = 0;
                        cosTheta = 1;
                        endP.Y = iniP.Y;
                    }
                    if (tanTheta < 0.63 && tanTheta > 0.52) //30
                    {
                        sinTheta = sinTheta / Math.Abs(sinTheta) * 0.5;
                        cosTheta = cosTheta / Math.Abs(cosTheta) * 0.866;
                        endP.Y = iniP.Y - d * sinTheta;
                        endP.X = iniP.X + d * cosTheta;
                    }
                    if (tanTheta < 1.08 && tanTheta > 0.92) //45
                    {
                        sinTheta = sinTheta / Math.Abs(sinTheta) * 0.707;
                        cosTheta = cosTheta / Math.Abs(cosTheta) * 0.707;
                        endP.Y = iniP.Y - d * sinTheta;
                        endP.X = iniP.X + d * cosTheta;
                    }
                    if (tanTheta < 1.95 && tanTheta > 1.63) //60
                    {
                        sinTheta = sinTheta / Math.Abs(sinTheta) * 0.866;
                        cosTheta = cosTheta / Math.Abs(cosTheta) * 0.5;
                        endP.Y = iniP.Y - d * sinTheta;
                        endP.X = iniP.X + d * cosTheta;
                    }
                    if (Math.Abs(cosTheta / sinTheta) < 1.0 / 12)
                    {
                        endP.X = iniP.X;
                        sinTheta = 1;
                        cosTheta = 0;
                    }
                    strokes.Add(GenerateLineStroke(new Point(iniP.X - 3 * x * sinTheta, iniP.Y - 3 * x * cosTheta), new Point(endP.X - 3 * x * sinTheta, endP.Y - 3 * x * cosTheta)));
                    strokes.Add(GenerateLineStroke(new Point(iniP.X - x * sinTheta, iniP.Y - x * cosTheta), new Point(endP.X - x * sinTheta, endP.Y - x * cosTheta)));
                    strokes.Add(GenerateLineStroke(new Point(iniP.X + x * sinTheta, iniP.Y + x * cosTheta), new Point(endP.X + x * sinTheta, endP.Y + x * cosTheta)));
                    strokes.Add(GenerateLineStroke(new Point(iniP.X + 3 * x * sinTheta, iniP.Y + 3 * x * cosTheta), new Point(endP.X + 3 * x * sinTheta, endP.Y + 3 * x * cosTheta)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 11:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    strokes.Add(GenerateArrowLineStroke(new Point(2 * iniP.X - (endP.X - 20), iniP.Y), new Point(endP.X, iniP.Y)));
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, 2 * iniP.Y - (endP.Y + 20)), new Point(iniP.X, endP.Y)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 12:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    if (Math.Abs(iniP.X - endP.X) < 0.01) return;
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X + (iniP.X - endP.X) / Math.Abs(iniP.X - endP.X) * 25, iniP.Y), new Point(endP.X, iniP.Y)));
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, 2 * iniP.Y - (endP.Y + 20)), new Point(iniP.X, endP.Y)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 13:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    if (Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    strokes.Add(GenerateArrowLineStroke(new Point(2 * iniP.X - (endP.X - 20), iniP.Y), new Point(endP.X, iniP.Y)));
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, iniP.Y + (iniP.Y - endP.Y) / Math.Abs(iniP.Y - endP.Y) * 25), new Point(iniP.X, endP.Y)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 14:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X + (iniP.X - endP.X) / Math.Abs(iniP.X - endP.X) * 25, iniP.Y), new Point(endP.X, iniP.Y)));
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, iniP.Y + (iniP.Y - endP.Y) / Math.Abs(iniP.Y - endP.Y) * 25), new Point(iniP.X, endP.Y)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 17:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, iniP.Y), new Point(iniP.X + Math.Abs(endP.X - iniP.X), iniP.Y)));
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, iniP.Y), new Point(iniP.X, iniP.Y - Math.Abs(endP.Y - iniP.Y))));
                    d = (Math.Abs(iniP.X - endP.X) + Math.Abs(iniP.Y - endP.Y)) / 2;
                    strokes.Add(GenerateArrowLineStroke(new Point(iniP.X, iniP.Y), new Point(iniP.X - d / 1.76, iniP.Y + d / 1.76)));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 3:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point(iniP.X, iniP.Y),
                        new System.Windows.Point(iniP.X, endP.Y),
                        new System.Windows.Point(endP.X, endP.Y),
                        new System.Windows.Point(endP.X, iniP.Y),
                        new System.Windows.Point(iniP.X, iniP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    inkCanvas.Strokes.Add(stroke);
                    break;
                case 19:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    double a = iniP.X - endP.X;
                    double b = iniP.Y - endP.Y;
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point(iniP.X - a, iniP.Y - b),
                        new System.Windows.Point(iniP.X - a, iniP.Y + b),
                        new System.Windows.Point(iniP.X + a, iniP.Y + b),
                        new System.Windows.Point(iniP.X + a, iniP.Y - b),
                        new System.Windows.Point(iniP.X - a, iniP.Y - b)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                        inkCanvas.Strokes.Remove(lastTempCenterDotStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    lastTempCenterDotStroke = GenerateCenterDotStroke(iniP); //中心圆点标记（见方法注释）
                    inkCanvas.Strokes.Add(stroke);
                    inkCanvas.Strokes.Add(lastTempCenterDotStroke);
                    break;
                case 4:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    pointList = GenerateEllipseGeometry(iniP, endP);
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    inkCanvas.Strokes.Add(stroke);
                    break;
                case 5:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    double R = GetDistance(iniP, endP);
                    pointList = GenerateEllipseGeometry(new Point(iniP.X - R, iniP.Y - R), new Point(iniP.X + R, iniP.Y + R));
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                        inkCanvas.Strokes.Remove(lastTempCenterDotStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    lastTempCenterDotStroke = GenerateCenterDotStroke(iniP); //圆心圆点标记
                    inkCanvas.Strokes.Add(stroke);
                    inkCanvas.Strokes.Add(lastTempCenterDotStroke);
                    break;
                case 16:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    double halfA = endP.X - iniP.X;
                    double halfB = endP.Y - iniP.Y;
                    pointList = GenerateEllipseGeometry(new Point(iniP.X - halfA, iniP.Y - halfB), new Point(iniP.X + halfA, iniP.Y + halfB));
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStroke);
                        inkCanvas.Strokes.Remove(lastTempCenterDotStroke);
                    }
                    catch { }
                    lastTempStroke = stroke;
                    lastTempCenterDotStroke = GenerateCenterDotStroke(iniP); //椭圆中心圆点标记
                    inkCanvas.Strokes.Add(stroke);
                    inkCanvas.Strokes.Add(lastTempCenterDotStroke);
                    break;
                case 23:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    a = Math.Abs(endP.X - iniP.X);
                    b = Math.Abs(endP.Y - iniP.Y);
                    pointList = GenerateEllipseGeometry(new Point(iniP.X - a, iniP.Y - b), new Point(iniP.X + a, iniP.Y + b));
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke);
                    double c = Math.Sqrt(Math.Abs(a * a - b * b));
                    StylusPoint stylusPoint;
                    if (a > b)
                    {
                        stylusPoint = new StylusPoint(iniP.X + c, iniP.Y, (float)1.0);
                        point = new StylusPointCollection();
                        point.Add(stylusPoint);
                        stroke = new Stroke(point)
                        {
                            DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                        };
                        strokes.Add(stroke.Clone());
                        stylusPoint = new StylusPoint(iniP.X - c, iniP.Y, (float)1.0);
                        point = new StylusPointCollection();
                        point.Add(stylusPoint);
                        stroke = new Stroke(point)
                        {
                            DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                        };
                        strokes.Add(stroke.Clone());
                    }
                    else if (a < b)
                    {
                        stylusPoint = new StylusPoint(iniP.X, iniP.Y - c, (float)1.0);
                        point = new StylusPointCollection();
                        point.Add(stylusPoint);
                        stroke = new Stroke(point)
                        {
                            DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                        };
                        strokes.Add(stroke.Clone());
                        stylusPoint = new StylusPoint(iniP.X, iniP.Y + c, (float)1.0);
                        point = new StylusPointCollection();
                        point.Add(stylusPoint);
                        stroke = new Stroke(point)
                        {
                            DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                        };
                        strokes.Add(stroke.Clone());
                    }
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 10:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    R = GetDistance(iniP, endP);
                    strokes = GenerateDashedLineEllipseStrokeCollection(new Point(iniP.X - R, iniP.Y - R), new Point(iniP.X + R, iniP.Y + R));
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 24:
                case 25:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    //双曲线 x^2/a^2 - y^2/b^2 = 1
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    var pointList2 = new List<Point>();
                    var pointList3 = new List<Point>();
                    var pointList4 = new List<Point>();
                    if (drawMultiStepShapeCurrentStep == 0)
                    {
                        //第一笔：画渐近线
                        double k = Math.Abs((endP.Y - iniP.Y) / (endP.X - iniP.X));
                        strokes.Add(GenerateDashedLineStrokeCollection(new Point(2 * iniP.X - endP.X, 2 * iniP.Y - endP.Y), endP));
                        strokes.Add(GenerateDashedLineStrokeCollection(new Point(2 * iniP.X - endP.X, endP.Y), new Point(endP.X, 2 * iniP.Y - endP.Y)));
                        drawMultiStepShapeSpecialParameter3 = k;
                        drawMultiStepShapeSpecialStrokeCollection = strokes;
                    }
                    else
                    {
                        //第二笔：画双曲线
                        double k = drawMultiStepShapeSpecialParameter3;
                        bool isHyperbolaFocalPointOnXAxis = Math.Abs((endP.Y - iniP.Y) / (endP.X - iniP.X)) < k;
                        if (isHyperbolaFocalPointOnXAxis)
                        { // 焦点在 x 轴上
                            a = Math.Sqrt(Math.Abs((endP.X - iniP.X) * (endP.X - iniP.X) - (endP.Y - iniP.Y) * (endP.Y - iniP.Y) / (k * k)));
                            b = a * k;
                            pointList = new List<Point>();
                            for (double i = a; i <= Math.Abs(endP.X - iniP.X); i += 0.5)
                            {
                                double rY = Math.Sqrt(Math.Abs(k * k * i * i - b * b));
                                pointList.Add(new Point(iniP.X + i, iniP.Y - rY));
                                pointList2.Add(new Point(iniP.X + i, iniP.Y + rY));
                                pointList3.Add(new Point(iniP.X - i, iniP.Y - rY));
                                pointList4.Add(new Point(iniP.X - i, iniP.Y + rY));
                            }
                        }
                        else
                        { // 焦点在 y 轴上
                            a = Math.Sqrt(Math.Abs((endP.Y - iniP.Y) * (endP.Y - iniP.Y) - (endP.X - iniP.X) * (endP.X - iniP.X) * (k * k)));
                            b = a / k;
                            pointList = new List<Point>();
                            for (double i = a; i <= Math.Abs(endP.Y - iniP.Y); i += 0.5)
                            {
                                double rX = Math.Sqrt(Math.Abs(i * i / k / k - b * b));
                                pointList.Add(new Point(iniP.X - rX, iniP.Y + i));
                                pointList2.Add(new Point(iniP.X + rX, iniP.Y + i));
                                pointList3.Add(new Point(iniP.X - rX, iniP.Y - i));
                                pointList4.Add(new Point(iniP.X + rX, iniP.Y - i));
                            }
                        }
                        try
                        {
                            point = new StylusPointCollection(pointList);
                            stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                            strokes.Add(stroke.Clone());
                            point = new StylusPointCollection(pointList2);
                            stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                            strokes.Add(stroke.Clone());
                            point = new StylusPointCollection(pointList3);
                            stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                            strokes.Add(stroke.Clone());
                            point = new StylusPointCollection(pointList4);
                            stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                            strokes.Add(stroke.Clone());
                            if (drawingShapeMode == 25)
                            {
                                //画焦点
                                c = Math.Sqrt(a * a + b * b);
                                stylusPoint = isHyperbolaFocalPointOnXAxis ? new StylusPoint(iniP.X + c, iniP.Y, (float)1.0) : new StylusPoint(iniP.X, iniP.Y + c, (float)1.0);
                                point = new StylusPointCollection();
                                point.Add(stylusPoint);
                                stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                                strokes.Add(stroke.Clone());
                                stylusPoint = isHyperbolaFocalPointOnXAxis ? new StylusPoint(iniP.X - c, iniP.Y, (float)1.0) : new StylusPoint(iniP.X, iniP.Y - c, (float)1.0);
                                point = new StylusPointCollection();
                                point.Add(stylusPoint);
                                stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                                strokes.Add(stroke.Clone());
                            }
                        }
                        catch
                        {
                            return;
                        }
                    }
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 20:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    //抛物线 y=ax^2
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    a = (iniP.Y - endP.Y) / ((iniP.X - endP.X) * (iniP.X - endP.X));
                    pointList = new List<Point>();
                    pointList2 = new List<Point>();
                    for (double i = 0.0; i <= Math.Abs(endP.X - iniP.X); i += 0.5)
                    {
                        pointList.Add(new Point(iniP.X + i, iniP.Y - a * i * i));
                        pointList2.Add(new Point(iniP.X - i, iniP.Y - a * i * i));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    point = new StylusPointCollection(pointList2);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 21:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    //抛物线 y^2=ax
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    a = (iniP.X - endP.X) / ((iniP.Y - endP.Y) * (iniP.Y - endP.Y));
                    pointList = new List<Point>();
                    pointList2 = new List<Point>();
                    for (double i = 0.0; i <= Math.Abs(endP.Y - iniP.Y); i += 0.5)
                    {
                        pointList.Add(new Point(iniP.X - a * i * i, iniP.Y + i));
                        pointList2.Add(new Point(iniP.X - a * i * i, iniP.Y - i));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    point = new StylusPointCollection(pointList2);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 22:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    //抛物线 y^2=ax, 含焦点
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    double p = (iniP.Y - endP.Y) * (iniP.Y - endP.Y) / (2 * (iniP.X - endP.X));
                    a = 0.5 / p;
                    pointList = new List<Point>();
                    pointList2 = new List<Point>();
                    for (double i = 0.0; i <= Math.Abs(endP.Y - iniP.Y); i += 0.5)
                    {
                        pointList.Add(new Point(iniP.X - a * i * i, iniP.Y + i));
                        pointList2.Add(new Point(iniP.X - a * i * i, iniP.Y - i));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    point = new StylusPointCollection(pointList2);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    stylusPoint = new StylusPoint(iniP.X - p / 2, iniP.Y, (float)1.0);
                    point = new StylusPointCollection();
                    point.Add(stylusPoint);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 26:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    //正弦曲线 y = A·sin(2πx/T)
                    //拖拽交互：按下点 = 曲线起点（相当于原点），
                    //          水平拖动距离 = 一个周期 T，纵向拖动距离 = 振幅 A
                    if (Math.Abs(iniP.X - endP.X) < 0.01 || Math.Abs(iniP.Y - endP.Y) < 0.01) return;
                    double sinPeriod = Math.Abs(endP.X - iniP.X); //周期 T = 水平拖动距离
                    double sinAmplitude = iniP.Y - endP.Y;        //振幅 A（向上拖为正，符合数学 y 轴向上）
                    double sinDirX = endP.X > iniP.X ? 1 : -1;   //曲线延伸方向：向右拖向右画，向左拖向左画
                    pointList = new List<Point>();
                    for (double i = 0.0; i <= sinPeriod; i += 0.5)
                    {
                        //屏幕坐标 y 向下、数学 y 向上，因此 sin 项取负号
                        pointList.Add(new Point(iniP.X + sinDirX * i, iniP.Y - sinAmplitude * Math.Sin(2 * Math.PI * i / sinPeriod)));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 6:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    newIniP = iniP;
                    if (iniP.Y > endP.Y)
                    {
                        newIniP = new Point(iniP.X, endP.Y);
                        endP = new Point(endP.X, iniP.Y);
                    }
                    double topA = Math.Abs(newIniP.X - endP.X);
                    double topB = topA / 2.646;
                    //顶部椭圆
                    pointList = GenerateEllipseGeometry(new Point(newIniP.X, newIniP.Y - topB / 2), new Point(endP.X, newIniP.Y + topB / 2));
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    //底部椭圆
                    pointList = GenerateEllipseGeometry(new Point(newIniP.X, endP.Y - topB / 2), new Point(endP.X, endP.Y + topB / 2), false, true);
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    strokes.Add(GenerateDashedLineEllipseStrokeCollection(new Point(newIniP.X, endP.Y - topB / 2), new Point(endP.X, endP.Y + topB / 2), true, false));
                    //左侧
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point(newIniP.X, newIniP.Y),
                        new System.Windows.Point(newIniP.X, endP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    //右侧
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point(endP.X, newIniP.Y),
                        new System.Windows.Point(endP.X, endP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 7:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    if (iniP.Y > endP.Y)
                    {
                        newIniP = new Point(iniP.X, endP.Y);
                        endP = new Point(endP.X, iniP.Y);
                    }
                    double bottomA = Math.Abs(newIniP.X - endP.X);
                    double bottomB = bottomA / 2.646;
                    //底部椭圆
                    pointList = GenerateEllipseGeometry(new Point(newIniP.X, endP.Y - bottomB / 2), new Point(endP.X, endP.Y + bottomB / 2), false, true);
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    strokes.Add(GenerateDashedLineEllipseStrokeCollection(new Point(newIniP.X, endP.Y - bottomB / 2), new Point(endP.X, endP.Y + bottomB / 2), true, false));
                    //左侧
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point((newIniP.X + endP.X) / 2, newIniP.Y),
                        new System.Windows.Point(newIniP.X, endP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    //右侧
                    pointList = new List<System.Windows.Point>{
                        new System.Windows.Point((newIniP.X + endP.X) / 2, newIniP.Y),
                        new System.Windows.Point(endP.X, endP.Y)
                    };
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                    try
                    {
                        inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                    }
                    catch { }
                    lastTempStrokeCollection = strokes;
                    inkCanvas.Strokes.Add(strokes);
                    break;
                case 9:
                    _currentCommitType = CommitReason.ShapeDrawing;
                    if (isFirstTouchCuboid)
                    {
                        //分开画线条方便后期单独擦除某一条棱
                        strokes.Add(GenerateLineStroke(new Point(iniP.X, iniP.Y), new Point(iniP.X, endP.Y)));
                        strokes.Add(GenerateLineStroke(new Point(iniP.X, endP.Y), new Point(endP.X, endP.Y)));
                        strokes.Add(GenerateLineStroke(new Point(endP.X, endP.Y), new Point(endP.X, iniP.Y)));
                        strokes.Add(GenerateLineStroke(new Point(iniP.X, iniP.Y), new Point(endP.X, iniP.Y)));
                        try
                        {
                            inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                        }
                        catch { }
                        lastTempStrokeCollection = strokes;
                        inkCanvas.Strokes.Add(strokes);
                        CuboidFrontRectIniP = iniP;
                        CuboidFrontRectEndP = endP;
                    }
                    else
                    {
                        d = CuboidFrontRectIniP.Y - endP.Y;
                        if (d < 0) d = -d; //就是懒不想做反向的，不要让我去做，想做自己做好之后 Pull Request
                        a = CuboidFrontRectEndP.X - CuboidFrontRectIniP.X; //正面矩形长
                        b = CuboidFrontRectEndP.Y - CuboidFrontRectIniP.Y; //正面矩形宽

                        //横上
                        Point newLineIniP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectIniP.Y - d);
                        Point newLineEndP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectIniP.Y - d);
                        pointList = new List<System.Windows.Point> { newLineIniP, newLineEndP };
                        point = new StylusPointCollection(pointList);
                        stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                        strokes.Add(stroke.Clone());
                        //横下 (虚线)
                        newLineIniP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectEndP.Y - d);
                        newLineEndP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectEndP.Y - d);
                        strokes.Add(GenerateDashedLineStrokeCollection(newLineIniP, newLineEndP));
                        //斜左上
                        newLineIniP = new Point(CuboidFrontRectIniP.X, CuboidFrontRectIniP.Y);
                        newLineEndP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectIniP.Y - d);
                        pointList = new List<System.Windows.Point> { newLineIniP, newLineEndP };
                        point = new StylusPointCollection(pointList);
                        stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                        strokes.Add(stroke.Clone());
                        //斜右上
                        newLineIniP = new Point(CuboidFrontRectEndP.X, CuboidFrontRectIniP.Y);
                        newLineEndP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectIniP.Y - d);
                        pointList = new List<System.Windows.Point> { newLineIniP, newLineEndP };
                        point = new StylusPointCollection(pointList);
                        stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                        strokes.Add(stroke.Clone());
                        //斜左下 (虚线)
                        newLineIniP = new Point(CuboidFrontRectIniP.X, CuboidFrontRectEndP.Y);
                        newLineEndP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectEndP.Y - d);
                        strokes.Add(GenerateDashedLineStrokeCollection(newLineIniP, newLineEndP));
                        //斜右下
                        newLineIniP = new Point(CuboidFrontRectEndP.X, CuboidFrontRectEndP.Y);
                        newLineEndP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectEndP.Y - d);
                        pointList = new List<System.Windows.Point> { newLineIniP, newLineEndP };
                        point = new StylusPointCollection(pointList);
                        stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                        strokes.Add(stroke.Clone());
                        //竖左 (虚线)
                        newLineIniP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectIniP.Y - d);
                        newLineEndP = new Point(CuboidFrontRectIniP.X + d, CuboidFrontRectEndP.Y - d);
                        strokes.Add(GenerateDashedLineStrokeCollection(newLineIniP, newLineEndP));
                        //竖右
                        newLineIniP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectIniP.Y - d);
                        newLineEndP = new Point(CuboidFrontRectEndP.X + d, CuboidFrontRectEndP.Y - d);
                        pointList = new List<System.Windows.Point> { newLineIniP, newLineEndP };
                        point = new StylusPointCollection(pointList);
                        stroke = new Stroke(point) { DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone() };
                        strokes.Add(stroke.Clone());

                        try
                        {
                            inkCanvas.Strokes.Remove(lastTempStrokeCollection);
                        }
                        catch { }
                        lastTempStrokeCollection = strokes;
                        inkCanvas.Strokes.Add(strokes);
                    }
                    break;
            }
        }

        bool isFirstTouchCuboid = true;
        Point CuboidFrontRectIniP = new Point();
        Point CuboidFrontRectEndP = new Point();

        private void Main_Grid_TouchUp(object sender, TouchEventArgs e)
        {
            inkCanvas_MouseUp(sender, null);
            if (dec.Count == 0)
            {
                isWaitUntilNextTouchDown = false;
            }
        }
        Stroke lastTempStroke = null;
        //中心图形（圆/中心椭圆/中心矩形）的中心圆点标记：拖拽过程中与图形本体一起"擦了重画"的临时笔迹
        Stroke lastTempCenterDotStroke = null;

        /// <summary>
        /// 生成"中心圆点"标记笔迹：圆/中心椭圆/中心矩形画完后在中心（按下点）位置点一个圆点，
        /// 让"按下点 = 图形中心"这个交互有明确的视觉锚点（教学上也方便讲圆心/对称中心）。
        /// 圆点比当前画笔稍微粗一点（1.8 倍，保证在图形线条旁清晰可见），颜色跟随画笔
        /// ——与图形笔迹"选中什么笔画出什么样"同一大原则。
        /// </summary>
        private Stroke GenerateCenterDotStroke(Point center)
        {
            var attrs = inkCanvas.DefaultDrawingAttributes.Clone();
            double dotSize = attrs.Width * 1.8; //稍微比当前笔迹粗一点
            attrs.Width = dotSize;
            attrs.Height = dotSize;
            attrs.FitToCurve = false; //单个点，无需平滑
            var point = new StylusPointCollection();
            point.Add(new StylusPoint(center.X, center.Y, (float)1.0));
            return new Stroke(point) { DrawingAttributes = attrs };
        }
        StrokeCollection lastTempStrokeCollection = new StrokeCollection();
        bool isWaitUntilNextTouchDown = false;
        private List<System.Windows.Point> GenerateEllipseGeometry(System.Windows.Point st, System.Windows.Point ed, bool isDrawTop = true, bool isDrawBottom = true)
        {
            double a = 0.5 * (ed.X - st.X);
            double b = 0.5 * (ed.Y - st.Y);
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            if (isDrawTop && isDrawBottom)
            {
                for (double r = 0; r <= 2 * Math.PI; r = r + 0.01)
                {
                    pointList.Add(new System.Windows.Point(0.5 * (st.X + ed.X) + a * Math.Cos(r), 0.5 * (st.Y + ed.Y) + b * Math.Sin(r)));
                }
            }
            else
            {
                if (isDrawBottom)
                {
                    for (double r = 0; r <= Math.PI; r = r + 0.01)
                    {
                        pointList.Add(new System.Windows.Point(0.5 * (st.X + ed.X) + a * Math.Cos(r), 0.5 * (st.Y + ed.Y) + b * Math.Sin(r)));
                    }
                }
                if (isDrawTop)
                {
                    for (double r = Math.PI; r <= 2 * Math.PI; r = r + 0.01)
                    {
                        pointList.Add(new System.Windows.Point(0.5 * (st.X + ed.X) + a * Math.Cos(r), 0.5 * (st.Y + ed.Y) + b * Math.Sin(r)));
                    }
                }
            }
            return pointList;
        }

        private StrokeCollection GenerateDashedLineEllipseStrokeCollection(System.Windows.Point st, System.Windows.Point ed, bool isDrawTop = true, bool isDrawBottom = true)
        {
            double a = 0.5 * (ed.X - st.X);
            double b = 0.5 * (ed.Y - st.Y);
            double step = 0.05;
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            StylusPointCollection point;
            Stroke stroke;
            StrokeCollection strokes = new StrokeCollection();
            if (isDrawBottom)
            {
                for (double i = 0.0; i < 1.0; i += step * 1.66)
                {
                    pointList = new List<Point>();
                    for (double r = Math.PI * i; r <= Math.PI * (i + step); r = r + 0.01)
                    {
                        pointList.Add(new System.Windows.Point(0.5 * (st.X + ed.X) + a * Math.Cos(r), 0.5 * (st.Y + ed.Y) + b * Math.Sin(r)));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                }
            }
            if (isDrawTop)
            {
                for (double i = 1.0; i < 2.0; i += step * 1.66)
                {
                    pointList = new List<Point>();
                    for (double r = Math.PI * i; r <= Math.PI * (i + step); r = r + 0.01)
                    {
                        pointList.Add(new System.Windows.Point(0.5 * (st.X + ed.X) + a * Math.Cos(r), 0.5 * (st.Y + ed.Y) + b * Math.Sin(r)));
                    }
                    point = new StylusPointCollection(pointList);
                    stroke = new Stroke(point)
                    {
                        DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                    };
                    strokes.Add(stroke.Clone());
                }
            }
            return strokes;
        }

        private Stroke GenerateLineStroke(System.Windows.Point st, System.Windows.Point ed)
        {
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            StylusPointCollection point;
            Stroke stroke;
            pointList = new List<System.Windows.Point>{
                new System.Windows.Point(st.X, st.Y),
                new System.Windows.Point(ed.X, ed.Y)
            };
            point = new StylusPointCollection(pointList);
            stroke = new Stroke(point)
            {
                DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
            };
            return stroke;
        }

        private Stroke GenerateArrowLineStroke(System.Windows.Point st, System.Windows.Point ed)
        {
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            StylusPointCollection point;
            Stroke stroke;

            double w = 20, h = 7;
            double theta = Math.Atan2(st.Y - ed.Y, st.X - ed.X);
            double sint = Math.Sin(theta);
            double cost = Math.Cos(theta);

            pointList = new List<Point>
            {
                new Point(st.X, st.Y),
                new Point(ed.X , ed.Y),
                new Point(ed.X + (w * cost - h * sint), ed.Y + (w * sint + h * cost)),
                new Point(ed.X,ed.Y),
                new Point(ed.X + (w * cost + h * sint), ed.Y - (h * cost - w * sint))
            };
            point = new StylusPointCollection(pointList);
            stroke = new Stroke(point)
            {
                DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
            };
            return stroke;
        }

        private StrokeCollection GenerateDashedLineStrokeCollection(System.Windows.Point st, System.Windows.Point ed)
        {
            double step = 5;
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            StylusPointCollection point;
            Stroke stroke;
            StrokeCollection strokes = new StrokeCollection();
            double d = GetDistance(st, ed);
            double sinTheta = (ed.Y - st.Y) / d;
            double cosTheta = (ed.X - st.X) / d;
            for (double i = 0.0; i < d; i += step * 2.76)
            {
                pointList = new List<System.Windows.Point>{
                    new System.Windows.Point(st.X + i * cosTheta, st.Y + i * sinTheta),
                    new System.Windows.Point(st.X + Math.Min(i + step, d) * cosTheta, st.Y + Math.Min(i + step, d) * sinTheta)
                };
                point = new StylusPointCollection(pointList);
                stroke = new Stroke(point)
                {
                    DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                };
                strokes.Add(stroke.Clone());
            }
            return strokes;
        }

        private StrokeCollection GenerateDotLineStrokeCollection(System.Windows.Point st, System.Windows.Point ed)
        {
            double step = 3;
            List<System.Windows.Point> pointList = new List<System.Windows.Point>();
            StylusPointCollection point;
            Stroke stroke;
            StrokeCollection strokes = new StrokeCollection();
            double d = GetDistance(st, ed);
            double sinTheta = (ed.Y - st.Y) / d;
            double cosTheta = (ed.X - st.X) / d;
            for (double i = 0.0; i < d; i += step * 2.76)
            {
                var stylusPoint = new StylusPoint(st.X + i * cosTheta, st.Y + i * sinTheta, (float)0.8);
                point = new StylusPointCollection();
                point.Add(stylusPoint);
                stroke = new Stroke(point)
                {
                    DrawingAttributes = inkCanvas.DefaultDrawingAttributes.Clone()
                };
                strokes.Add(stroke.Clone());
            }
            return strokes;
        }

        bool isMouseDown = false;
        private void inkCanvas_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isMouseDown = true;
            if (NeedUpdateIniP())
            {
                iniP = e.GetPosition(inkCanvas);
            }
        }

        private void inkCanvas_MouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseDown)
            {
                MouseTouchMove(e.GetPosition(inkCanvas));
            }
        }

        private void inkCanvas_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (drawingShapeMode == 5)
            {
                Circle circle = new Circle(new Point(), 0, lastTempStroke);
                circle.R = GetDistance(circle.Stroke.StylusPoints[0].ToPoint(), circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].ToPoint()) / 2;
                circle.Centroid = new Point((circle.Stroke.StylusPoints[0].X + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].X) / 2,
                                            (circle.Stroke.StylusPoints[0].Y + circle.Stroke.StylusPoints[circle.Stroke.StylusPoints.Count / 2].Y) / 2);
                circles.Add(circle);
            }
            if (drawingShapeMode != 9 && drawingShapeMode != 0 && drawingShapeMode != 24 && drawingShapeMode != 25)
            {
                BtnPen_Click(null, null); //画完一次还原到笔模式
            }
            if (drawingShapeMode == 9)
            {
                if (isFirstTouchCuboid)
                {
                    if (CuboidStrokeCollection == null) CuboidStrokeCollection = new StrokeCollection();
                    isFirstTouchCuboid = false;
                    Point newIniP = new Point(Math.Min(CuboidFrontRectIniP.X, CuboidFrontRectEndP.X), Math.Min(CuboidFrontRectIniP.Y, CuboidFrontRectEndP.Y));
                    Point newEndP = new Point(Math.Max(CuboidFrontRectIniP.X, CuboidFrontRectEndP.X), Math.Max(CuboidFrontRectIniP.Y, CuboidFrontRectEndP.Y));
                    CuboidFrontRectIniP = newIniP;
                    CuboidFrontRectEndP = newEndP;
                    CuboidStrokeCollection.Add(lastTempStrokeCollection);
                }
                else
                {
                    BtnPen_Click(null, null); //画完还原到笔模式
                    if (_currentCommitType == CommitReason.ShapeDrawing)
                    {
                        CuboidStrokeCollection.Add(lastTempStrokeCollection);
                        _currentCommitType = CommitReason.UserInput;
                        timeMachine.CommitStrokeUserInputHistory(CuboidStrokeCollection);
                        //长方体两步画完 → 打标签 + 自动选中（统一入口，见 MW_GraphStrokes.cs）
                        InsertGraphStrokes(CuboidStrokeCollection);
                        CuboidStrokeCollection = null;
                    }
                }
            }
            if (drawingShapeMode == 24 || drawingShapeMode == 25)
            {
                if (drawMultiStepShapeCurrentStep == 0)
                {
                    //第一笔（渐近线）完成：只进入"等待第二笔"状态，不提交历史、不自动选中
                    drawMultiStepShapeCurrentStep = 1;
                }
                else
                {
                    //第二笔（双曲线本体）完成：渐近线 + 双曲线合并提交历史、合并选中
                    drawMultiStepShapeCurrentStep = 0;
                    StrokeCollection hyperbolaGraph = null;
                    if (drawMultiStepShapeSpecialStrokeCollection != null)
                    {
                        bool opFlag = false;
                        switch (Settings.Canvas.HyperbolaAsymptoteOption)
                        {
                            case OptionalOperation.Yes:
                                opFlag = true;
                                break;
                            case OptionalOperation.No:
                                opFlag = false;
                                break;
                            case OptionalOperation.Ask:
                                opFlag = MessageBox.Show("是否移除渐近线？", "Inkboard", MessageBoxButton.YesNo) != MessageBoxResult.Yes;
                                break;
                        };
                        if (opFlag)
                        {
                            //保留渐近线：并入双曲线一起提交（一个图形、一个撤销单元）
                            hyperbolaGraph = new StrokeCollection();
                            foreach (Stroke s in drawMultiStepShapeSpecialStrokeCollection) hyperbolaGraph.Add(s);
                        }
                        else
                        {
                            inkCanvas.Strokes.Remove(drawMultiStepShapeSpecialStrokeCollection);
                        }
                    }
                    //并入双曲线本体（第二笔画出的笔迹 = lastTempStrokeCollection）
                    if (hyperbolaGraph == null) hyperbolaGraph = new StrokeCollection();
                    if (lastTempStrokeCollection != null)
                    {
                        foreach (Stroke s in lastTempStrokeCollection) hyperbolaGraph.Add(s);
                    }
                    if (hyperbolaGraph.Count > 0)
                    {
                        //渐近线 + 双曲线作为一个整体提交历史：Ctrl+Z 一步全撤
                        _currentCommitType = CommitReason.UserInput;
                        timeMachine.CommitStrokeUserInputHistory(hyperbolaGraph);
                        //整组打标签 + 自动选中（渐近线和曲线一起动）
                        InsertGraphStrokes(hyperbolaGraph);
                    }
                    BtnPen_Click(null, null); //画完还原到笔模式
                }
            }
            isMouseDown = false;
            if (ReplacedStroke != null || AddedStroke != null)
            {
                timeMachine.CommitStrokeEraseHistory(ReplacedStroke, AddedStroke);
                AddedStroke = null;
                ReplacedStroke = null;
            }
            //注意：多步图形（24/25 双曲线）第一笔只是中间产物（渐近线），不能走通用提交口——
            //否则第一笔就触发自动选中，用户点空白取消选中时收尾逻辑会切回笔模式，
            //drawingShapeMode 被清零，第二笔（双曲线本体）就画不出来了。
            //渐近线与双曲线的合并提交在下方 drawingShapeMode==24/25 的完成分支里做。
            //（长方体 case 9 天然不走此口：它的第一笔在自己的 MouseUp 分支内处理）
            if (_currentCommitType == CommitReason.ShapeDrawing && drawingShapeMode != 9
                && drawingShapeMode != 24 && drawingShapeMode != 25)
            {
                _currentCommitType = CommitReason.UserInput;
                StrokeCollection collection = null;
                if (lastTempStrokeCollection != null && lastTempStrokeCollection.Count > 0)
                {
                    collection = lastTempStrokeCollection;
                }
                else if (lastTempStroke != null)
                {
                    collection = new StrokeCollection() { lastTempStroke };
                    //中心圆点与图形本体一起进撤销历史、一起打标签选中（拖动/缩放/整组擦除都整体走）
                    if (lastTempCenterDotStroke != null) collection.Add(lastTempCenterDotStroke);
                }
                if (collection != null)
                {
                    timeMachine.CommitStrokeUserInputHistory(collection);

                    //图形画完 → 打标签 + 自动选中（统一入口，见 MW_GraphStrokes.cs）
                    InsertGraphStrokes(collection);
                }
            }
            lastTempStroke = null;
            lastTempStrokeCollection = null;
            lastTempCenterDotStroke = null;
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

        private bool NeedUpdateIniP()
        {
            if (drawingShapeMode == 24 || drawingShapeMode == 25)
            {
                if (drawMultiStepShapeCurrentStep == 1) return false;
            }
            return true;
        }

        #endregion Shape Drawing
    }
}
