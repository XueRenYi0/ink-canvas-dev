using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：页面图片层（截图/本地图片的存储与显示，最小版）。
    /// 设计原则（简单版，不自造交互）：
    /// 1. 图片直接作为 inkCanvas 的子元素——WPF InkCanvas 原生支持子元素渲染，
    ///    墨迹天然画在图片上方；图片本身只读（IsHitTestVisible=false），
    ///    需要"选中拖动/删除"等编辑能力时后期再从备份分支并入原生选中方案。
    /// 2. 图片永远存到"白板抽屉"：不管截图时在哪个模式，都挂在 CurrentWhiteboardIndex
    ///    对应的白板页上；屏幕注释模式不显示图片（桌面批注层贴图场景少，用户定稿）。
    /// 3. 图片只在白板/黑板模式且页号匹配时可见；换页/切模式只切显隐，元素常驻不销毁。
    /// 4. 清屏/擦除不影响图片（导入的教学素材防误删）。
    /// 5. 滚动同步：滚动是"坐标物化"（矩阵直接改墨迹点数据），图片用 SetTop 同步平移；
    ///    最大可滚深度把图片底边也计入（纯图片页也能滚）。
    /// 6. 模块自包含：所有入口 try/catch 包裹，异常只通知不抛出，绝不影响主流程。
    /// </summary>
    public partial class MainWindow
    {
        #region 页面图片层

        /// <summary>按白板页号存储的图片列表：1~99 = 白板页</summary>
        readonly Dictionary<int, List<Image>> _pageImages = new Dictionary<int, List<Image>>();

        /// <summary>图片 → 页号 反查表（删除时快速定位归属页）</summary>
        readonly Dictionary<Image, int> _imagePageKey = new Dictionary<Image, int>();

        /// <summary>上次刷新时的页号缓存（避免低频兜底定时器无谓遍历）</summary>
        int _lastImageVisibleKey = -2; // -2 = 未初始化；-1 = 注释模式（无可见图片）

        /// <summary>
        /// 初始化图片层（MainWindow 构造流程调用）：
        /// 挂 EditingModeChanged——只有"选择工具"激活时图片才可被点选/拖动，
        /// 笔/橡皮等模式下图片完全不参与命中（书写擦除零干扰）。
        /// </summary>
        private void InitImageLayer()
        {
            try
            {
                inkCanvas.EditingModeChanged += (s, e) => ImageLayer_UpdateHitTest();
            }
            catch { }
        }

        /// <summary>
        /// 删除当前选中的图片（Delete 键调用）。
        /// 原生选中的元素通过 GetSelectedElements 拿到，再逐个从页表与画布移除。
        /// </summary>
        internal void ImageLayer_DeleteSelectedImages()
        {
            try
            {
                var selected = inkCanvas.GetSelectedElements()
                    .OfType<Image>()
                    .Where(img => _imagePageKey.ContainsKey(img))
                    .ToList();
                if (selected.Count == 0) return;

                foreach (var img in selected)
                {
                    int key = _imagePageKey[img];
                    _imagePageKey.Remove(img);
                    if (_pageImages.TryGetValue(key, out var list))
                    {
                        list.Remove(img);
                        if (list.Count == 0) _pageImages.Remove(key);
                    }
                    inkCanvas.Children.Remove(img);
                }
                _lastImageVisibleKey = -2; // 强制下次刷新重新同步
            }
            catch { }
        }

        /// <summary>同步所有图片的命中开关：选择模式=可选中可拖，其他模式=完全穿透</summary>
        private void ImageLayer_UpdateHitTest()
        {
            try
            {
                bool selectable = inkCanvas.EditingMode == InkCanvasEditingMode.Select;
                foreach (var img in _imagePageKey.Keys)
                    img.IsHitTestVisible = selectable;
            }
            catch { }
        }

        /// <summary>级联偏移计数器：连续插入的多张图错开摆放（每张右下偏移 24px，5 张一轮），避免完全叠死</summary>
        int _imageCascadeCounter = 0;

        /// <summary>
        /// 插入一张图片到白板当前页（自动按画布适配缩小、落画布左上角）。
        /// 落点页 = 当前视图：白板模式挂当前白板页；注释模式（第 0 页）挂批注层（键 -1，
        /// 与墨迹 strokeCollections[0] 同层同生死）——粘贴在哪个视图就出现在哪个视图。
        /// cascade=true 时应用级联偏移（连续插入多张不完全叠死）。
        /// </summary>
        private Image ImageLayer_AddImage(BitmapSource source, bool cascade = false, Point? center = null)
        {
            try
            {
                if (source == null || source.PixelWidth < 1) return null;

                // 适配缩放：显示尺寸不超过画布 80%；小图不放大（保持原尺寸清晰）
                double hostW = inkCanvas.ActualWidth, hostH = inkCanvas.ActualHeight;
                if (hostW <= 0) hostW = SystemParameters.WorkArea.Width;
                if (hostH <= 0) hostH = SystemParameters.WorkArea.Height;
                double scale = Math.Min(hostW * 0.8 / source.PixelWidth, hostH * 0.8 / source.PixelHeight);
                if (scale > 1) scale = 1;

                var img = new Image
                {
                    Source = source,
                    Width = source.PixelWidth * scale,
                    Height = source.PixelHeight * scale,
                    Stretch = Stretch.Fill,
                    // 默认不参与命中（书写穿透）；选择工具激活时由 ImageLayer_UpdateHitTest 打开
                    IsHitTestVisible = inkCanvas.EditingMode == InkCanvasEditingMode.Select
                };

                // 图片本体拖动三件套（见下方 region）：
                // WPF InkCanvas 原生只支持"抓选框边缘"拖动元素，按在图片本体（中央）
                // 只会重复点选——此处补齐"抓图片任意位置即可拖动"的行业通用交互
                img.MouseLeftButtonDown += ImageLayer_DragMouseDown;
                img.MouseMove += ImageLayer_DragMouseMove;
                img.MouseLeftButtonUp += ImageLayer_DragMouseUp;

                // 落画布左上角（留 20px 边距，保证图片边缘可被抓取拖动）；
                // 级联时每张右下错开 24px（一轮 5 张，超出画布自动收边）
                double left = 20, top = 20;
                if (cascade)
                {
                    double off = (_imageCascadeCounter++ % 5) * 24;
                    left = Math.Min(20 + off, Math.Max(0, hostW - img.Width));
                    top = Math.Min(20 + off, Math.Max(0, hostH - img.Height));
                }
                if (center.HasValue)
                {
                    // 指定位置（粘贴气泡场景）：图片中心对准点击点，钳制在画布内
                    left = Math.Max(4, Math.Min(center.Value.X - img.Width / 2, hostW - img.Width - 4));
                    top = Math.Max(4, Math.Min(center.Value.Y - img.Height / 2, hostH - img.Height - 4));
                }
                InkCanvas.SetLeft(img, left);
                InkCanvas.SetTop(img, top);

                // 登记到页表（InkCanvas 继承自 Canvas，子元素用 Left/Top 绝对定位）
                // 键 = 当前视图：白板模式=当前白板页；注释模式（第 0 页）=-1 批注层
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                if (!_pageImages.TryGetValue(key, out var list))
                {
                    list = new List<Image>();
                    _pageImages[key] = list;
                }
                list.Add(img);
                _imagePageKey[img] = key;

                inkCanvas.Children.Add(img);
                // 刚插入的就是当前视图，直接可见（翻页/切模式由 RefreshVisibility 接管）
                img.Visibility = Visibility.Visible;

                // 插入即选中（与"墨迹变图形"同一交互口径，见 InsertGraphStrokes）：
                // 截图/粘贴到达画布的图默认带选框+操作条+手柄，拖动/缩放/适配白板宽度立即可用；
                // 做其他操作（点空白/激活工具）时取消选中——复用一次性选中机制，
                // 取消后笔模式自动还原，如同没有这回事
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    try
                    {
                        // ① 先定型笔输入（若已是 Ink 模式，这行是幂等的安全垫）
                        forceEraser = false;
                        inkCanvas.EditingMode = InkCanvasEditingMode.Ink;

                        // ② 选中刚插入的图（屏蔽 SelectionChanged 的快照副作用，快照由选中后重新捕获）
                        // 注意：Select() 会把 EditingMode 切成 Select（WPF 行为），
                        // 靠 _isOneShotGraphSelection 标志在取消选中时把笔模式换回来
                        isProgramChangeStrokeSelection = true;
                        try { inkCanvas.Select(new StrokeCollection(), new List<UIElement> { img }); } catch { }
                        isProgramChangeStrokeSelection = false;
                        _isOneShotGraphSelection = inkCanvas.GetSelectedElements().Count > 0;

                        // ③ 显示选区遮罩与控制条（拖动/缩放/操作条立即可用）
                        GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                        inkCanvas_SelectionChanged(inkCanvas, EventArgs.Empty);
                        updateBorderStrokeSelectionControlLocation();
                    }
                    catch { }
                }));

                _lastImageVisibleKey = -2; // 强制下次 RefreshVisibility 重新同步
                return img;
            }
            catch (Exception ex)
            {
                ShowNotification("插入图片失败：" + ex.Message);
                return null;
            }
        }

        #region 图片本体拖动（补齐 WPF 原生缺口）

        /// <summary>
        /// 【为什么需要这段代码】WPF InkCanvas 的元素拖动只能从"选区装饰器"
        /// （选框边框、手柄、选框内非元素区域）发起——装饰器的可命中区域是
        /// "选框减去元素本体"（WPF 源码 DrawBackgound 用透明画刷只覆盖该区域）。
        /// 按在图片本体（中央）时事件交给 Image 元素，只走"点选"路径不会移动。
        /// 本 region 补齐：已选中的图片，按住图片任意位置（含中央）即可直接拖动。
        /// 注意：混合选中（墨迹+图片）时选区覆盖层会先拦下事件，走
        /// MW_SelectionGestures.cs 的覆盖层拖动路径（那里已同步平移元素），
        /// 与本路径互斥不冲突。
        /// </summary>

        /// <summary>图片本体拖动是否进行中（MouseLeftButtonDown 到 Up 之间为 true）</summary>
        private bool _imageBodyDragActive = false;

        /// <summary>拖动起点（inkCanvas 坐标系——SetLeft/SetTop 同坐标系，直接做差）</summary>
        private Point _imageBodyDragStart;

        /// <summary>拖动开始时各选中图片的原始位置快照（整体平移基准，避免累积误差）</summary>
        private readonly Dictionary<Image, Point> _imageBodyDragOrigins = new Dictionary<Image, Point>();

        /// <summary>
        /// 图片上按下：仅"选择工具 + 图片已被选中"时接管为拖动。
        /// 未选中的图片不抢事件——保持原生"单击选中"体验（第一次点选中、第二次按住拖）。
        /// </summary>
        private void ImageLayer_DragMouseDown(object sender, MouseButtonEventArgs e)
        {
            try
            {
                if (e.ChangedButton != MouseButton.Left) return; // 右键等不接管
                if (inkCanvas.EditingMode != InkCanvasEditingMode.Select) return; // 仅选择工具
                if (!(sender is Image img)) return;

                // Ctrl/Shift+点击 = 原生多选切换（往选区里加/减这张图），交回原生逻辑，不接管成拖动
                if ((Keyboard.Modifiers & (ModifierKeys.Control | ModifierKeys.Shift)) != ModifierKeys.None) return;

                // 按下的这张必须在当前选中集合里（点未选中图片 = 原生点选语义，放行）
                var selected = inkCanvas.GetSelectedElements().OfType<Image>().ToList();
                if (!selected.Contains(img)) return;

                // 记录起点 + 所有选中图片的原始位置（多选时整体拖动）
                _imageBodyDragActive = true;
                _imageBodyDragStart = e.GetPosition(inkCanvas);
                _imageBodyDragOrigins.Clear();
                foreach (var s in selected)
                {
                    _imageBodyDragOrigins[s] = new Point(
                        InkCanvas.GetLeft(s), InkCanvas.GetTop(s));
                }

                // 捕获鼠标：拖出窗口边界也能持续收到 Move/Up（行业惯例）
                img.CaptureMouse();
                e.Handled = true; // 阻止事件继续冒泡触发原生的重复点选逻辑
            }
            catch { }
        }

        /// <summary>拖动中：位移量同步应用到所有选中图片（整体平移）；顺带做悬停光标反馈</summary>
        private void ImageLayer_DragMouseMove(object sender, MouseEventArgs e)
        {
            if (!(sender is Image hoverImg)) return;

            // 悬停反馈：已选中的图片显示四向移动光标（告诉用户"这里按住可拖"）；
            // 未选中显示默认光标（首次点击是"选中"语义，还不可直接拖）。
            // 鼠标移入/选中状态变化后的第一次移动即自动纠正，无需额外事件
            if (!_imageBodyDragActive)
            {
                try
                {
                    bool isSelected = inkCanvas.GetSelectedElements().Contains(hoverImg);
                    hoverImg.Cursor = isSelected ? Cursors.SizeAll : null;
                }
                catch { }
                return;
            }

            try
            {
                // 左键已松开（异常路径）→ 直接结束，防悬空状态
                if (e.LeftButton != MouseButtonState.Pressed)
                {
                    ImageLayer_EndBodyDrag(sender as Image);
                    return;
                }

                var pos = e.GetPosition(inkCanvas);
                double dx = pos.X - _imageBodyDragStart.X;
                double dy = pos.Y - _imageBodyDragStart.Y;

                // 用"原始位置 + 总位移"而不是逐次累加，避免浮点误差累积
                foreach (var kv in _imageBodyDragOrigins)
                {
                    InkCanvas.SetLeft(kv.Key, kv.Value.X + dx);
                    InkCanvas.SetTop(kv.Key, kv.Value.Y + dy);
                }
            }
            catch { }
        }

        /// <summary>抬起：结束拖动，释放鼠标捕获</summary>
        private void ImageLayer_DragMouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!_imageBodyDragActive) return;
            ImageLayer_EndBodyDrag(sender as Image);
        }

        /// <summary>收尾公共路径：清状态 + 释放捕获。图片移动不进 TimeMachine（与图片删除一致）</summary>
        private void ImageLayer_EndBodyDrag(Image img)
        {
            _imageBodyDragActive = false;
            _imageBodyDragOrigins.Clear();
            try { if (img != null && img.IsMouseCaptured) img.ReleaseMouseCapture(); } catch { }
        }

        #endregion

        /// <summary>
        /// 页号/模式变化时切换图片可见性（换页 UpdateIndexInfoDisplay 即时调用）。
        /// 可见条件：白板/黑板模式 且 页号匹配。注释模式（键 -1）一律隐藏。
        /// </summary>
        internal void ImageLayer_RefreshVisibility()
        {
            try
            {
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                if (key == _lastImageVisibleKey) return; // 缓存命中：页没变，什么都不做
                _lastImageVisibleKey = key;

                foreach (var kv in _pageImages)
                {
                    var visible = kv.Key == key ? Visibility.Visible : Visibility.Collapsed;
                    foreach (var img in kv.Value)
                        img.Visibility = visible;
                }
            }
            catch { }
        }

        /// <summary>
        /// 清屏时处理图片（BtnClear 调用，用户定稿的规则）：
        /// 只清"当前视图层"——白板模式 = 当前白板页图片；注释模式 = 批注层（-1）图片
        /// （图片与墨迹同生共死，清屏就是"全部清掉"；其他页不动，翻回去还在）。
        /// 注释模式清屏走了 inkCanvas.Children.Clear()，白板抽屉里其他页的图也被摘了，
        /// 结尾 EnsureHost 挂回（不可见，翻回对应页自然显示）。
        /// </summary>
        internal void ImageLayer_OnClearScreen()
        {
            try
            {
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                Helpers.LogHelper.NewLog($"[清图诊断] OnClearScreen 进入: currentMode={currentMode}, key={key}, " +
                    $"页表=[{string.Join(",", _pageImages.Select(kv => $"{kv.Key}页×{kv.Value.Count}张"))}]");

                if (_pageImages.TryGetValue(key, out var list))
                {
                    foreach (var img in list)
                    {
                        _imagePageKey.Remove(img);
                        inkCanvas.Children.Remove(img);
                    }
                    _pageImages.Remove(key);
                    Helpers.LogHelper.NewLog($"[清图诊断] 已删除第 {key} 页的 {list.Count} 张图片（数据+视觉树）");
                }

                // 注释模式：白板抽屉（其他页）的图被 Children.Clear 摘了，挂回（不可见）
                if (currentMode == 0) ImageLayer_EnsureHost();
                _lastImageVisibleKey = -2; // 强制下次刷新重新同步
            }
            catch (Exception ex) { Helpers.LogHelper.NewLog($"[清图诊断] OnClearScreen 异常: {ex.Message}"); }
        }

        /// <summary>
        /// 图片自愈：多人书写模式切换等处会 inkCanvas.Children.Clear()
        /// 把图片一并清掉，在这些调用点之后调用本方法把所有页的图片重新挂回
        /// （位置存在图片自身属性上，不丢失）。
        /// </summary>
        internal void ImageLayer_EnsureHost()
        {
            try
            {
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                // ===== 诊断日志：若发生"挂回"，说明有调用点在清屏删图之后又把图救活 =====
                var missing = _pageImages.SelectMany(kv => kv.Value).Where(img => !inkCanvas.Children.Contains(img)).ToList();
                if (missing.Count > 0)
                    Helpers.LogHelper.NewLog($"[清图诊断] EnsureHost 挂回 {missing.Count} 张图（调用堆栈将随异常栈打印）, 堆栈: {Environment.StackTrace}");

                foreach (var kv in _pageImages)
                {
                    var visible = kv.Key == key ? Visibility.Visible : Visibility.Collapsed;
                    foreach (var img in kv.Value)
                    {
                        if (!inkCanvas.Children.Contains(img))
                            inkCanvas.Children.Add(img);
                        img.Visibility = visible;
                    }
                }
                _lastImageVisibleKey = -2; // 强制下次刷新重新同步
            }
            catch { }
        }

        #region 滚动同步（与墨迹"坐标物化"滚动配套）

        /// <summary>
        /// 滚动同步：ScrollNote 平移墨迹的同帧调用，把当前可见页的图片同步平移
        /// （actual&gt;0 = 内容上移，图片 Top 减小；与墨迹矩阵完全同向同量）。
        /// </summary>
        internal void ImageLayer_OnScrolled(double actual)
        {
            try
            {
                if (Math.Abs(actual) < 1) return;
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                if (!_pageImages.TryGetValue(key, out var list)) return; // -1 批注层同样同步
                foreach (var img in list)
                {
                    double top = InkCanvas.GetTop(img);
                    InkCanvas.SetTop(img, top - actual);
                }
            }
            catch { }
        }

        /// <summary>
        /// 当前可见页图片的最底部 Y 坐标（滚动深度计算用）；无图片返回 double.NaN。
        /// 供 GetMaxScroll 把图片底边计入可滚深度（纯图片页也能滚）。
        /// </summary>
        internal double ImageLayer_GetContentBottom()
        {
            try
            {
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
                if (!_pageImages.TryGetValue(key, out var list) || list.Count == 0)
                    return double.NaN; // -1 批注层的图片同样计入
                double bottom = double.MinValue;
                foreach (var img in list)
                {
                    double b = InkCanvas.GetTop(img) + img.Height;
                    if (b > bottom) bottom = b;
                }
                return bottom;
            }
            catch { return double.NaN; }
        }

        #endregion

        #endregion
    }
}
