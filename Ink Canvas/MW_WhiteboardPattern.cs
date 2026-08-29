using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：白板底纹（方格/横线）。
    /// 设计要点：
    /// 1. 底纹画在 GridBackgroundCover 内的 BorderWhiteboardPattern 上（DrawingBrush 平铺 + Freeze），
    ///    不是 Stroke——橡皮擦不掉、撤销不涉及、选择选不中、99 页共用；
    /// 2. 底纹固定在屏幕上：笔记滚动（坐标物化）与 Ctrl+滚轮缩放只动墨迹，底纹不动；
    /// 3. 类型与间距即改即存（Settings.Canvas.WhiteboardPattern / WhiteboardGridSize），下次启动继承；
    /// 4. 唯一设置入口：白板空白处右键菜单（无/方格/横线 + 间距滑块），黑板/白板按钮右键同菜单；
    ///    菜单每次现建，勾选与滑块值实时准确。线粗固定 0.5px（辅助参考线，不提供调节）；
    ///    菜单顶部提供白板/黑板快速切换（复用 BtnSwitchTheme_Click 全套联动）。
    /// </summary>
    public partial class MainWindow
    {
        #region Whiteboard Pattern（白板底纹）

        /// <summary>当前板面是否为浅色（白板）。按 GridBackgroundCover 底色亮度判断</summary>
        private bool IsWhiteboardBoardLight()
        {
            var bg = GridBackgroundCover.Background as SolidColorBrush;
            if (bg != null)
            {
                var c = bg.Color;
                return (c.R * 299 + c.G * 587 + c.B * 114) / 1000 > 128;
            }
            return true;
        }

        /// <summary>
        /// 按当前设置重建底纹 Brush 并应用到底纹层。
        /// 仅在切换类型/间距/深浅色时调用，开销可忽略。
        /// </summary>
        private void ApplyWhiteboardPattern()
        {
            if (BorderWhiteboardPattern == null) return; // 早于 InitializeComponent 的极端调用路径防御

            int pattern = Settings.Canvas.WhiteboardPattern;
            if (pattern <= 0)
            {
                BorderWhiteboardPattern.Background = null;
                return;
            }

            double size = Settings.Canvas.WhiteboardGridSize;
            if (size < 16) size = 16;
            if (size > 240) size = 240;

            // 线粗 0.5px（辅助参考线，纤细不抢视觉）；线色随板面明暗自适应
            bool boardIsLight = IsWhiteboardBoardLight();
            Color lineColor = boardIsLight
                ? Color.FromArgb(0x99, 0x99, 0x99, 0x99)   // 白板：浅灰（约 60% 不透明度）
                : Color.FromArgb(0x33, 0xFF, 0xFF, 0xFF);  // 黑板：约 20% 白

            var pen = new Pen(new SolidColorBrush(lineColor), 0.5);
            pen.Freeze();

            var geometry = new GeometryGroup();
            if (pattern == 1)
            {
                // 方格：整格矩形轮廓。相邻格的共享边在同一位置重叠描边，视觉上仍是 1px 线
                geometry.Children.Add(new RectangleGeometry(new Rect(0, 0, size, size)));
            }
            else
            {
                // 横线：每格画上、下两条边——下边缘与下一格上边缘重叠，视觉只有一组横线。
                // 关键：必须让图案包围盒撑满整格（高=size），否则 DrawingBrush 默认 Stretch=Fill
                // 会把仅 1px 高的图案纵向拉伸填满整格，导致"整屏变灰、格界出竖线"
                geometry.Children.Add(new LineGeometry(new Point(0, 0), new Point(size, 0)));
                geometry.Children.Add(new LineGeometry(new Point(0, size), new Point(size, size)));
            }
            geometry.Freeze();

            // 0.5px 参考线：1px 线对齐物理像素，避免发虚
            var drawing = new GeometryDrawing(null, pen, geometry);
            var group = new DrawingGroup();
            group.Children.Add(drawing);
            group.GuidelineSet = new GuidelineSet(new[] { 0.5 }, new[] { 0.5 });
            group.Freeze();

            var brush = new DrawingBrush(group)
            {
                TileMode = TileMode.Tile,
                Viewport = new Rect(0, 0, size, size),
                ViewportUnits = BrushMappingMode.Absolute
            };
            brush.Freeze();
            BorderWhiteboardPattern.Background = brush;
        }

        /// <summary>设置底纹类型（0=无 1=方格 2=横线）：写设置、存盘、应用</summary>
        private void SetWhiteboardPattern(int pattern)
        {
            if (pattern < 0 || pattern > 2) return;
            Settings.Canvas.WhiteboardPattern = pattern;
            SaveSettingsToFile();
            ApplyWhiteboardPattern();
        }

        /// <summary>设置底纹间距：写设置、存盘、应用</summary>
        private void SetWhiteboardGridSize(double size)
        {
            Settings.Canvas.WhiteboardGridSize = size;
            SaveSettingsToFile();
            ApplyWhiteboardPattern();
        }

        /// <summary>启动时应用底纹（由设置加载流程调用）</summary>
        private void InitWhiteboardPatternUI()
        {
            ApplyWhiteboardPattern();
        }

        /// <summary>
        /// 白板空白处右键：弹出底纹快捷设置菜单。
        /// 仅白板/黑板模式（GridBackgroundCover 可见）且无选区时生效；
        /// 选中墨迹时右键保留原有行为，屏幕/PPT 批注模式右键行为不变。
        /// </summary>
        private void inkCanvas_PreviewMouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            //有选中笔迹时：弹"识别为函数"菜单（框选识别入口，见 MW_MathGraph.cs）
            if (inkCanvas.GetSelectedStrokes().Count > 0)
            {
                ShowMathContextMenu();
                e.Handled = true;
                return;
            }
            if (GridBackgroundCover.Visibility != Visibility.Visible) return;
            ShowWhiteboardPatternMenu(inkCanvas);
            e.Handled = true;
        }

        /// <summary>
        /// 框选笔迹后的右键菜单：识别为函数（选中笔迹 → 数学识别 → 原位替换成函数图像）
        /// </summary>
        private void ShowMathContextMenu()
        {
            var menu = new ContextMenu();
            var item = new MenuItem { Header = "识别为函数" };
            item.Click += (s, args) => RecognizeSelectedAsFunction();
            menu.Items.Add(item);
            menu.IsOpen = true;
        }

        /// <summary>
        /// 黑板/白板切换按钮上右键：同样弹出底纹菜单（双入口）。
        /// 白板守卫：该按钮在所有模式常驻（含屏幕批注），但底纹只画在板面上——
        /// 板面收起（GridBackgroundCover 隐藏）时右键不弹菜单，避免"改了设置却看不到效果"的困惑。
        /// </summary>
        private void BtnSwitchTheme_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (GridBackgroundCover.Visibility != Visibility.Visible) return;
            ShowWhiteboardPatternMenu(BtnSwitchTheme);
            e.Handled = true;
        }

        /// <summary>
        /// 悬浮栏底纹图标点击：触摸/手写笔可达入口（右键白板空白处为鼠标入口，同一菜单）。
        /// 菜单从按钮锚点弹出（不依赖鼠标位置），白板模式下随时可用。
        /// </summary>
        private void GridWhiteboardPatternIcon_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ShowWhiteboardPatternMenu(GridWhiteboardPattern, placeBelow: true);
        }

        /// <summary>
        /// 构建并弹出底纹菜单（每次现建，勾选与滑块值实时准确）。
        /// 顶部为白板/黑板切换（复用 BtnSwitchTheme_Click 全套联动：板面色、按钮文案、UI 主题、底纹线色）；
        /// 类型用 MenuItem（点选后菜单关闭）；间距滑块作为普通控件直接放进菜单——
        /// 非 MenuItem 项不会触发菜单关闭，可连续拖动实时预览，点空白处/Esc 收起。
        /// </summary>
        private void ShowWhiteboardPatternMenu(FrameworkElement target, bool placeBelow = false)
        {
            var menu = new ContextMenu();

            // 白板/黑板切换：按当前板面显示"切换到对方"
            bool boardIsLightNow = IsWhiteboardBoardLight();
            var toggleItem = new MenuItem
            {
                Header = boardIsLightNow ? "切换为黑板" : "切换为白板"
            };
            toggleItem.Click += (s, ev) => BtnSwitchTheme_Click(BtnSwitchTheme, null);
            menu.Items.Add(toggleItem);

            menu.Items.Add(new Separator());

            string[] patternNames = { "无底纹", "方格", "横线" };
            for (int i = 0; i < patternNames.Length; i++)
            {
                int p = i;
                var item = new MenuItem
                {
                    Header = patternNames[i],
                    IsCheckable = true,
                    IsChecked = Settings.Canvas.WhiteboardPattern == p
                };
                item.Click += (s, ev) => SetWhiteboardPattern(p);
                menu.Items.Add(item);
            }

            menu.Items.Add(new Separator());

            // 间距滑块行：标签 + 数值 + 滑块（实时生效即存盘）
            var valueText = new TextBlock
            {
                Text = Settings.Canvas.WhiteboardGridSize.ToString("0"),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0, 0, 0),
                MinWidth = 20
            };
            var slider = new Slider
            {
                Minimum = 16, Maximum = 240,
                Width = 150,
                Value = Settings.Canvas.WhiteboardGridSize,
                TickFrequency = 8,
                IsSnapToTickEnabled = true,
                VerticalAlignment = VerticalAlignment.Center
            };
            slider.ValueChanged += (s, ev) =>
            {
                valueText.Text = slider.Value.ToString("0");
                SetWhiteboardGridSize(slider.Value);
            };
            var row = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(12, 4, 12, 8) };
            row.Children.Add(new TextBlock { Text = "间距", VerticalAlignment = VerticalAlignment.Center });
            row.Children.Add(valueText);
            row.Children.Add(slider);
            menu.Items.Add(row);

            // 定位：按钮触发时弹在锚点下方（触摸/笔场景无可靠鼠标位置）；右键触发时跟随鼠标
            menu.PlacementTarget = target;
            menu.Placement = placeBelow ? PlacementMode.Bottom : PlacementMode.MousePoint;
            if (placeBelow) menu.VerticalOffset = 4;
            menu.IsOpen = true;
        }

        #endregion Whiteboard Pattern
    }
}
