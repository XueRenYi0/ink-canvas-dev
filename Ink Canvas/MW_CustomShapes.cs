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

        //整理模式（触摸屏删除入口）：开启后缩略图显示 × 角标，点角标即删；
        //此模式下点缩略图本体不进入插入模式（防整理时误插入）
        private bool _libraryEditMode = false;

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
                //走 ShowShapePanel 统一入口：面板已解挂为自由小窗口，显示前需要先定位
                ShowToastNotification("已存入图库 — 点击此处查看", () =>
                {
                    ShowShapePanel();
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存图形失败：" + ex.Message);
            }
        }

        #endregion 存入

        #region 面板加载

        /// <summary>扫描 CustomShapes 目录，重建"我的图形"缩略图墙；无图形时整区折叠隐藏</summary>
        private void LoadCustomShapes()
        {
            if (StackPanelCustomShapes == null) return;
            StackPanelCustomShapes.Children.Clear();
            try
            {
                if (Directory.Exists(CustomShapesDir))
                {
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

                        StackPanelCustomShapes.Children.Add(BuildShapeThumbCell(thumb, file));
                    }
                }
            }
            catch (Exception ex)
            {
                //不再静默吞掉（上次"整理后图全消失"就是因为异常被吞、面板误判为空）
                System.Diagnostics.Debug.WriteLine("[LoadCustomShapes] 加载失败: " + ex.Message);
            }

            int count = StackPanelCustomShapes.Children.Count;

            //图数统计（0 个时隐藏计数，只留"图库"二字也不显示——整区都折叠了）
            TextBlockLibraryCount.Text = count > 0 ? $"· {count} 个" : "";

            //图删空时自动退出整理模式（角标无目标可挂，按钮状态回"整理"）
            if (count == 0 && _libraryEditMode)
            {
                _libraryEditMode = false;
                UpdateLibraryEditVisual();
            }

            //空状态折叠整区（标题+行），不占面板空间
            StackPanelCustomShapesRow.Visibility =
                count == 0 ? Visibility.Collapsed : Visibility.Visible;
        }

        /// <summary>
        /// 构建单个图库缩略图格：64×56 格子（用户存的图形比内置图标金贵，格子放大一档）。
        /// 缩略图一律填满格子（全貌可见）；整理模式下叠加红色 × 角标（触摸屏删除入口）。
        /// </summary>
        private Border BuildShapeThumbCell(ImageSource thumb, string file)
        {
            Border cell = new Border();
            cell.Width = 64;
            cell.Height = 56;
            cell.Margin = new Thickness(2);
            cell.CornerRadius = new CornerRadius(4);
            cell.Background = Brushes.Transparent;
            cell.BorderThickness = new Thickness(1);
            cell.BorderBrush = new SolidColorBrush(Color.FromArgb(40, 128, 128, 128)); //淡灰描边让格子可见
            cell.Cursor = Cursors.Hand;
            cell.Tag = file;
            cell.ToolTip = _libraryEditMode
                ? "整理模式：点 × 删除此图形"
                : "左键点击后到画布上落笔插入\n右键可删除此图形";
            cell.MouseDown += Border_MouseDown; //复用按下记录，配合 MouseUp 判定同一次点击
            cell.MouseUp += CustomShapeButton_MouseUp;
            if (!_libraryEditMode) cell.ContextMenu = BuildCustomShapeContextMenu(file); //整理模式右键菜单多余，不挂

            Image img = new Image();
            img.Source = thumb;
            img.Stretch = Stretch.Uniform; //填满格子显示全貌（渲染时已按 maxSize 等比缩放）
            img.HorizontalAlignment = HorizontalAlignment.Center;
            img.VerticalAlignment = VerticalAlignment.Center;

            if (_libraryEditMode)
            {
                //× 角标：红底白字小圆，叠在格子右上角；点击即删（Handled 拦住冒泡，不触发插入）
                Border badge = new Border();
                badge.Width = 16;
                badge.Height = 16;
                badge.CornerRadius = new CornerRadius(8);
                badge.Background = new SolidColorBrush(Color.FromRgb(229, 72, 77));
                badge.HorizontalAlignment = HorizontalAlignment.Right;
                badge.VerticalAlignment = VerticalAlignment.Top;
                badge.Margin = new Thickness(0, 1, 1, 0);
                badge.Cursor = Cursors.Hand;
                badge.ToolTip = "删除此图形";
                TextBlock x = new TextBlock();
                x.Text = "×";
                x.Foreground = Brushes.White;
                x.FontSize = 11;
                x.FontWeight = FontWeights.Bold;
                x.HorizontalAlignment = HorizontalAlignment.Center;
                x.VerticalAlignment = VerticalAlignment.Center;
                badge.Child = x;
                badge.MouseLeftButtonUp += (s, args) =>
                {
                    args.Handled = true; //不冒泡到格子本体（否则会触发插入逻辑）
                    try
                    {
                        File.Delete(file);
                        LoadCustomShapes(); //立即刷新（角标全重建）
                        ShowToastNotification("已删除该图形");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("删除失败：" + ex.Message);
                    }
                };

                //格子内容 = 缩略图 + 角标（Grid 一次挂载，不要先 cell.Child=img 再加 Grid——WPF 元素只能有一个父级）
                Grid g = new Grid();
                img.Margin = new Thickness(2); //留出角标空间
                g.Children.Add(img);
                g.Children.Add(badge);
                cell.Child = g;
            }
            else
            {
                cell.Child = img;
            }

            return cell;
        }

        /// <summary>"整理"按钮：切换整理模式（iOS 桌面式——缩略图显示 × 角标，触摸屏删除入口）</summary>
        private void BorderLibraryEdit_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            _libraryEditMode = !_libraryEditMode;
            UpdateLibraryEditVisual();
            LoadCustomShapes(); //重挂缩略图（生成/移除 × 角标）
        }

        /// <summary>整理按钮的视觉状态：开启 = 蓝底白字"完成"，关闭 = 透明"整理"</summary>
        private void UpdateLibraryEditVisual()
        {
            TextBlockLibraryEdit.Text = _libraryEditMode ? "完成" : "整理";
            BorderLibraryEdit.Background = _libraryEditMode
                ? new SolidColorBrush(Color.FromArgb(140, 0, 136, 255))
                : Brushes.Transparent;
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

        /// <summary>渲染墨迹集合缩略图：按 maxSize 等比缩放填满（小图形放大、大图形缩小，全貌一律可见）</summary>
        private ImageSource RenderShapeThumbnail(StrokeCollection strokes)
        {
            try
            {
                Rect bounds = strokes.GetBounds();
                if (bounds.IsEmpty || bounds.Width <= 0 || bounds.Height <= 0) return null;

                const double margin = 6;
                const int maxSize = 120;

                //缩放系数：一律填满 maxSize（不设 1.0 上限——小图形也放大到满格，全貌可见且视觉一致；
                //位图渲染 maxSize 显示在格子里相当于超采样，清晰度无损）
                double rawW = bounds.Width + margin * 2;
                double rawH = bounds.Height + margin * 2;
                double k = Math.Min(maxSize / rawW, maxSize / rawH);

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
            //整理模式下点缩略图本体不进入插入模式（此时点击的意图是管理而非使用）
            if (_libraryEditMode) return;

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

            //统一走图形管理入口（打标签 + 自动选中 + 控制条），与函数图像等行为一致
            InsertGraphStrokes(strokes);
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
