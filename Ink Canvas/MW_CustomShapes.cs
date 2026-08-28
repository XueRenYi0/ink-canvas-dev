using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using MessageBox = System.Windows.MessageBox;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：自定义图形（选中墨迹存为图形，加入图形快捷栏"我的图形"行）
    /// 存储：ISF 墨迹格式（StrokeCollection.Save，零第三方依赖），文件位于 App.RootPath\CustomShapes\
    /// 交互：选中控制条"存为图形" → 图形面板顶部新增缩略图按钮 → 点击进入插入模式 → 画布落笔处以原始大小插入并自动选中
    /// </summary>
    public partial class MainWindow
    {
        #region 自定义图形

        static string CustomShapesDir
        {
            get { try { return App.RootPath + "CustomShapes\\"; } catch { return "CustomShapes\\"; } }
        }

        //待插入的自定义图形（非空时处于"落笔插入"模式）
        StrokeCollection pendingCustomShape = null;
        //防触摸 promoted 鼠标事件双触发（StylusDown 与 MouseDown 会先后到达）
        DateTime lastCustomShapeInsertTime = DateTime.MinValue;

        /// <summary>初始化：挂载落笔事件、加载已存图形</summary>
        private void InitCustomShapes()
        {
            inkCanvas.PreviewStylusDown += InkCanvas_PreviewStylusDownForCustomShape;
            inkCanvas.PreviewMouseDown += InkCanvas_PreviewMouseDownForCustomShape;
            LoadCustomShapes();
        }

        #region 存入

        private void BorderStrokeSelectionSaveAsShape_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count == 0) return;

            try
            {
                Directory.CreateDirectory(CustomShapesDir);
                string file = Path.Combine(CustomShapesDir, Guid.NewGuid().ToString("N") + ".isc");
                using (FileStream fs = new FileStream(file, FileMode.Create))
                {
                    strokes.Save(fs); //ISF 二进制格式，含点数据与笔迹属性
                }
                LoadCustomShapes(); //立即刷新面板，立等可见
                //点击提示可直接打开图形面板（用户"不知道存到哪了"时一步直达）
                ShowToastNotification("已存入图库 — 点击此处查看", () =>
                {
                    BorderDrawShape.Visibility = Visibility.Visible;
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存图形失败：" + ex.Message);
            }
        }

        #endregion 存入

        #region 面板加载

        /// <summary>扫描 CustomShapes 目录，重建"我的图形"按钮行；无图形时整区折叠隐藏</summary>
        private void LoadCustomShapes()
        {
            if (StackPanelCustomShapes == null) return;
            StackPanelCustomShapes.Children.Clear();
            try
            {
                if (Directory.Exists(CustomShapesDir))
                {
                    bool isFirst = true;
                    //按保存时间排序：旧图形在前，新存的排后面（与"排在原图形后面"的增长方向一致）
                    foreach (string file in Directory.GetFiles(CustomShapesDir, "*.isc").OrderBy(f => File.GetLastWriteTime(f)))
                    {
                        StrokeCollection strokes;
                        using (FileStream fs = File.OpenRead(file))
                        {
                            strokes = new StrokeCollection(fs);
                        }
                        if (strokes.Count == 0) continue;

                        ImageSource thumb = RenderShapeThumbnail(strokes);
                        if (thumb == null) continue;

                        Image img = new Image();
                        img.Source = thumb;
                        img.MaxWidth = 40;
                        img.MaxHeight = 26;
                        img.Stretch = Stretch.Uniform;
                        //对齐内置图形行：行高50、图标垂直居中、首个图标左Margin=16（与内置行起点一致），后续图标左Margin=0（Spacing=10 控制间距）
                        img.Margin = isFirst ? new Thickness(16, 12, 0, 12) : new Thickness(0, 12, 0, 12);
                        isFirst = false;
                        img.VerticalAlignment = VerticalAlignment.Center;
                        img.Cursor = Cursors.Hand;
                        img.Tag = file;
                        img.ToolTip = "左键点击后到画布上落笔插入\n右键可删除此图形";
                        img.MouseDown += Border_MouseDown; //复用按下记录，配合 MouseUp 判定同一次点击
                        img.MouseUp += CustomShapeButton_MouseUp;
                        img.ContextMenu = BuildCustomShapeContextMenu(file);
                        StackPanelCustomShapes.Children.Add(img);
                    }
                }
            }
            catch { }

            //空状态折叠整区（标题+行），不占面板空间
            StackPanelCustomShapesRow.Visibility =
                StackPanelCustomShapes.Children.Count == 0 ? Visibility.Collapsed : Visibility.Visible;
        }

        /// <summary>缩略图右键菜单：删除自定义图形（带确认）</summary>
        private ContextMenu BuildCustomShapeContextMenu(string file)
        {
            ContextMenu menu = new ContextMenu();
            MenuItem deleteItem = new MenuItem();
            deleteItem.Header = "删除此图形";
            deleteItem.Click += (s, args) =>
            {
                try
                {
                    if (MessageBox.Show("确定删除这个自定义图形吗？\n（画布上已插入的内容不受影响）",
                        "删除自定义图形", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                    {
                        File.Delete(file);
                        LoadCustomShapes(); //立即刷新面板
                        ShowToastNotification("已删除该图形");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("删除失败：" + ex.Message);
                }
            };
            menu.Items.Add(deleteItem);
            return menu;
        }

        /// <summary>渲染墨迹集合缩略图（DrawingVisual + Stroke.Draw → RenderTargetBitmap）</summary>
        private ImageSource RenderShapeThumbnail(StrokeCollection strokes)
        {
            try
            {
                Rect bounds = strokes.GetBounds();
                if (bounds.IsEmpty || bounds.Width <= 0 || bounds.Height <= 0) return null;

                const double margin = 6;

                //缩放系数：图形整体（含 margin）必须装进 120px 内，笔画绘制与位图尺寸同步缩放
                //（bug 修复：此前只缩了位图尺寸没缩绘制内容，大于120px的图形只截到包围盒左上角空白区域，缩略图全空）
                const int maxSize = 120;
                double rawW = bounds.Width + margin * 2;
                double rawH = bounds.Height + margin * 2;
                double k = Math.Min(1.0, Math.Min(maxSize / rawW, maxSize / rawH));

                DrawingVisual dv = new DrawingVisual();
                using (DrawingContext dc = dv.RenderOpen())
                {
                    //变换：p' = (p - bounds左上角) * k + margin，缩放与平移合并进一个矩阵
                    Matrix mtx = new Matrix();
                    mtx.Scale(k, k);
                    mtx.Translate(margin - k * bounds.Left, margin - k * bounds.Top);
                    dc.PushTransform(new MatrixTransform(mtx));
                    foreach (Stroke s in strokes)
                    {
                        s.Draw(dc);
                    }
                    dc.Pop();
                }

                int w = Math.Max(1, (int)Math.Ceiling(rawW * k));
                int h = Math.Max(1, (int)Math.Ceiling(rawH * k));

                RenderTargetBitmap rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(dv);
                rtb.Freeze();
                return rtb;
            }
            catch { return null; }
        }

        #endregion 面板加载

        #region 插入

        private void CustomShapeButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            //只响应左键：右键释放会先于 ContextMenu 弹出触发本事件，若不拦截会误入插入模式
            if (e.ChangedButton != MouseButton.Left) return;
            if (lastBorderMouseDownObject != sender) return;

            string file = (sender as FrameworkElement)?.Tag as string;
            if (file == null || !File.Exists(file)) return;

            try
            {
                StrokeCollection strokes;
                using (FileStream fs = File.OpenRead(file))
                {
                    strokes = new StrokeCollection(fs);
                }
                if (strokes.Count == 0) return;

                pendingCustomShape = strokes;

                //进入插入模式：停用墨迹收集（同内置形状按钮的做法），等待画布落笔
                forceEraser = true;
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.IsManipulationEnabled = false;

                //自动隐藏面板（与图形面板 AutoHide 开关一致）
                if (ToggleSwitchDrawShapeBorderAutoHide.IsOn) BorderDrawShape.Visibility = Visibility.Collapsed;
            }
            catch { }
        }

        private void InkCanvas_PreviewStylusDownForCustomShape(object sender, StylusDownEventArgs e)
        {
            TryInsertPendingCustomShape(e.GetPosition(inkCanvas));
        }

        private void InkCanvas_PreviewMouseDownForCustomShape(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton != MouseButton.Left) return;
            TryInsertPendingCustomShape(e.GetPosition(inkCanvas));
        }

        /// <summary>落笔处插入待插入图形（原始大小），插入后自动选中并恢复笔模式</summary>
        private void TryInsertPendingCustomShape(Point p)
        {
            if (pendingCustomShape == null) return;
            if ((DateTime.Now - lastCustomShapeInsertTime).TotalMilliseconds < 300) return; //去重
            lastCustomShapeInsertTime = DateTime.Now;

            StrokeCollection strokes = pendingCustomShape;
            pendingCustomShape = null;

            //平移：图形包围盒左上角对齐落点
            Rect b = strokes.GetBounds();
            Matrix m = new Matrix();
            m.Translate(p.X - b.Left, p.Y - b.Top);
            strokes.Transform(m, false);

            inkCanvas.Strokes.Add(strokes); //走 StrokesChanged 常规历史，Ctrl+Z 一步撤掉

            //自动选中（屏蔽 SelectionChanged 的快照副作用，快照由插入后重新捕获）
            isProgramChangeStrokeSelection = true;
            try { inkCanvas.Select(strokes); } catch { }
            isProgramChangeStrokeSelection = false;

            //恢复笔输入
            forceEraser = false;
            inkCanvas.EditingMode = InkCanvasEditingMode.Ink;

            //显示选区遮罩与控制条（此时快照 = 插入状态）
            GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
            inkCanvas_SelectionChanged(inkCanvas, EventArgs.Empty);
            updateBorderStrokeSelectionControlLocation();
        }

        #endregion 插入

        #region 操作反馈

        /// <summary>
        /// 轻量 toast 通知：屏幕底部居中淡入，停留约1.6秒后淡出移除。
        /// 传入 onClick 则提示可点击（手型光标），点击执行动作并立即收起；
        /// 未传 onClick 时不拦截鼠标（IsHitTestVisible=false），不打断书写。失败静默（仅反馈增强，非关键路径）。
        /// </summary>
        private void ShowToastNotification(string message, Action onClick = null)
        {
            try
            {
                var rootGrid = Main_Grid;
                if (rootGrid == null) return;

                Border toast = new Border();
                toast.Background = new SolidColorBrush(Color.FromArgb(190, 0, 0, 0)); //半透明黑底白字，深浅背景都清晰
                toast.CornerRadius = new CornerRadius(18);
                toast.Padding = new Thickness(18, 8, 18, 8);
                toast.Opacity = 0;
                toast.HorizontalAlignment = HorizontalAlignment.Center;
                toast.VerticalAlignment = VerticalAlignment.Bottom;
                toast.Margin = new Thickness(0, 0, 0, SystemParameters.WorkArea.Height * 0.10);

                TextBlock tb = new TextBlock();
                tb.Text = message;
                tb.Foreground = Brushes.White;
                tb.FontSize = 14;
                toast.Child = tb;

                DispatcherTimer timer = null;

                //立即收起：快速淡出并移除（点击后/超时共用）
                Action dismiss = () =>
                {
                    try
                    {
                        if (timer != null) timer.Stop();
                        DoubleAnimation fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(250));
                        fadeOut.Completed += (s2, e2) => rootGrid.Children.Remove(toast);
                        toast.BeginAnimation(OpacityProperty, fadeOut);
                    }
                    catch { }
                };

                if (onClick != null)
                {
                    toast.IsHitTestVisible = true;
                    toast.Cursor = Cursors.Hand;
                    toast.MouseLeftButtonUp += (s, e) =>
                    {
                        try { onClick(); } catch { }
                        dismiss();
                    };
                }
                else
                {
                    toast.IsHitTestVisible = false;
                }

                rootGrid.Children.Add(toast); //加在最后 = 顶层渲染

                //淡入
                toast.BeginAnimation(OpacityProperty,
                    new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(150)));

                //停留后淡出并移除
                timer = new DispatcherTimer { Interval = TimeSpan.FromSeconds(1.6) };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    dismiss();
                };
                timer.Start();
            }
            catch { }
        }

        #endregion 操作反馈

        #endregion 自定义图形
    }
}
