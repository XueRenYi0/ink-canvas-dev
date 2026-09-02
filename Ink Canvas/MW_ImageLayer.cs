using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
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
        /// 无论当前是注释模式还是白板模式，图片都挂在 CurrentWhiteboardIndex 页
        /// （"上次看的那一页"），等切到白板就能看到。
        /// cascade=true 时应用级联偏移（连续插入多张不完全叠死）。
        /// </summary>
        private Image ImageLayer_AddImage(BitmapSource source, bool cascade = false)
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

                // 落画布左上角（留 20px 边距，保证图片边缘可被抓取拖动）；
                // 级联时每张右下错开 24px（一轮 5 张，超出画布自动收边）
                double left = 20, top = 20;
                if (cascade)
                {
                    double off = (_imageCascadeCounter++ % 5) * 24;
                    left = Math.Min(20 + off, Math.Max(0, hostW - img.Width));
                    top = Math.Min(20 + off, Math.Max(0, hostH - img.Height));
                }
                InkCanvas.SetLeft(img, left);
                InkCanvas.SetTop(img, top);

                // 登记到页表（InkCanvas 继承自 Canvas，子元素用 Left/Top 绝对定位）
                int key = CurrentWhiteboardIndex;
                if (!_pageImages.TryGetValue(key, out var list))
                {
                    list = new List<Image>();
                    _pageImages[key] = list;
                }
                list.Add(img);
                _imagePageKey[img] = key;

                inkCanvas.Children.Add(img);
                // 白板模式下立即可见；注释模式下等切到白板再显示
                img.Visibility = currentMode != 0 && key == CurrentWhiteboardIndex
                    ? Visibility.Visible : Visibility.Collapsed;

                _lastImageVisibleKey = -2; // 强制下次 RefreshVisibility 重新同步
                return img;
            }
            catch (Exception ex)
            {
                ShowNotification("插入图片失败：" + ex.Message);
                return null;
            }
        }

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
        /// - 白板模式：清屏 = 清墨迹 + 清本页图片（图片与墨迹同生共死，清屏就是"全部清掉"）；
        /// - 注释模式：图片不在桌面层显示（都存在白板抽屉），只需自愈挂回，不删任何图。
        /// </summary>
        internal void ImageLayer_OnClearScreen()
        {
            try
            {
                // ===== 诊断日志：定位"只有图片无墨迹时清屏不清图"bug =====
                // 打印当前模式/页号/字典内全部页号和每页图片数，以及视觉树上还挂着几张图
                Helpers.LogHelper.NewLog($"[清图诊断] OnClearScreen 进入: currentMode={currentMode}, " +
                    $"CurrentWhiteboardIndex={CurrentWhiteboardIndex}, " +
                    $"页表=[{string.Join(",", _pageImages.Select(kv => $"{kv.Key}页×{kv.Value.Count}张"))}], " +
                    $"视觉树图片数={inkCanvas.Children.OfType<Image>().Count()}");

                if (currentMode != 0)
                {
                    int key = CurrentWhiteboardIndex;
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
                    else
                    {
                        Helpers.LogHelper.NewLog($"[清图诊断] 第 {key} 页在页表中无记录！（图片可能挂在别的页号上）");
                    }
                }
                else
                {
                    Helpers.LogHelper.NewLog("[清图诊断] 注释模式：不删图，走自愈挂回");
                    // 注释模式：Children.Clear 后自愈挂回
                    ImageLayer_EnsureHost();
                    return;
                }
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
                if (key < 1 || !_pageImages.TryGetValue(key, out var list)) return;
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
                if (key < 1 || !_pageImages.TryGetValue(key, out var list) || list.Count == 0)
                    return double.NaN;
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
