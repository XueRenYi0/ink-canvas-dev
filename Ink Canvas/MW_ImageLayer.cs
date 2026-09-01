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
        /// 初始化图片层（MainWindow 构造流程调用）：目前无需事件挂接，
        /// 保留方法便于后期扩展（如选择工具命中切换）。
        /// </summary>
        private void InitImageLayer()
        {
            try { }
            catch { }
        }

        /// <summary>
        /// 插入一张图片到白板当前页（自动按工作区适配缩小、居中放置）。
        /// 无论当前是注释模式还是白板模式，图片都挂在 CurrentWhiteboardIndex 页
        /// （"上次看的那一页"），等切到白板就能看到。
        /// </summary>
        private void ImageLayer_AddImage(BitmapSource source)
        {
            try
            {
                if (source == null || source.PixelWidth < 1) return;

                // 适配缩放：显示尺寸不超过工作区 90% 宽 / 85% 高；小图不放大（保持原尺寸清晰）
                double maxW = SystemParameters.WorkArea.Width * 0.9;
                double maxH = SystemParameters.WorkArea.Height * 0.85;
                double scale = Math.Min(maxW / source.PixelWidth, maxH / source.PixelHeight);
                if (scale > 1) scale = 1;

                var img = new Image
                {
                    Source = source,
                    Width = source.PixelWidth * scale,
                    Height = source.PixelHeight * scale,
                    Stretch = Stretch.Fill,
                    // 只读展示：不参与命中测试，书写/擦除完全不受影响
                    IsHitTestVisible = false
                };

                // 居中放置（画布尺寸未就绪时退回工作区尺寸估算）
                double hostW = inkCanvas.ActualWidth, hostH = inkCanvas.ActualHeight;
                if (hostW <= 0) hostW = SystemParameters.WorkArea.Width;
                if (hostH <= 0) hostH = SystemParameters.WorkArea.Height;
                InkCanvas.SetLeft(img, Math.Max(0, (hostW - img.Width) / 2));
                InkCanvas.SetTop(img, Math.Max(0, (hostH - img.Height) / 2));

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
            }
            catch (Exception ex)
            {
                ShowNotification("插入图片失败：" + ex.Message);
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
        /// 图片自愈：清屏（BtnClear）/ 多人书写模式切换等处会 inkCanvas.Children.Clear()
        /// 把图片一并清掉，在这些调用点之后调用本方法把所有页的图片重新挂回
        /// （位置存在图片自身属性上，不丢失）。
        /// </summary>
        internal void ImageLayer_EnsureHost()
        {
            try
            {
                int key = currentMode != 0 ? CurrentWhiteboardIndex : -1;
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
