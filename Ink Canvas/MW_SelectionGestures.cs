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
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Application = System.Windows.Application;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using Timer = System.Timers.Timer;

namespace Ink_Canvas
{
    /// <summary>MainWindow 分部类：墨迹选区与手势（含浮动控件）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Selection Gestures

        #region Floating Control

        /// <summary>
        /// 防误触守卫用的"按下源"：XAML 里成对写 <c>MouseDown="Border_MouseDown" MouseUp="Xxx_MouseUp"</c>，
        /// 处理器首行以 <c>if (lastBorderMouseDownObject != sender) return;</c> 校验，
        /// 语义 = "按下与松开必须落在同一个元素上"，避免从按钮 A 拖到按钮 B 松手时误触发 B。
        /// 注意：漏写 XAML 里的 MouseDown 会让该按钮静默失效（守卫直接 return，不报错不记日志）。
        /// 例外：被代码直调（sender 传 null）的处理器改成 <c>if (sender != null &amp;&amp; ...) return;</c>。
        /// </summary>
        object lastBorderMouseDownObject;

        private void Border_MouseDown(object sender, MouseButtonEventArgs e)
        {
            lastBorderMouseDownObject = sender;
        }

        private void BorderStrokeSelectionClone_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            //入口含图片：纯图片选中同样可进复制模式（拖出图片副本）
            if (inkCanvas.GetSelectedStrokes().Count == 0 && GetSelectedPageImages().Count == 0) return;

            // 复制拖拽模式（PPT/Visio 惯例）：点图标只进入模式（图标变色提示），
            // 之后按住选中图形拖动 = 拖出一份副本（原件不动），松手落定，可连续拖出多份。
            // 原交互"点击即克隆出副本"已废弃——副本固定偏移 24px 落点不可控。
            ToggleCopyDragMode();
        }

        /// <summary>复制拖拽模式激活中：按住选中图形拖动会拖出副本（见 GridInkCanvasSelectionCover_MouseDown）</summary>
        bool isCopyDragMode = false;

        /// <summary>开关复制拖拽模式</summary>
        private void ToggleCopyDragMode()
        {
            isCopyDragMode = !isCopyDragMode;
            UpdateCopyDragModeVisual();
            if (isCopyDragMode)
            {
                // 进入模式时顺手把选中墨迹存入剪贴板（联动 Ctrl+V / 跨页粘贴）
                CopySelectedStrokesToClipboard();
                ShowNotification("复制模式：按住图形拖出副本，可连续多份；再点图标或点空白退出");
            }
        }

        /// <summary>图标高亮反馈：激活 = 淡蓝底（与各面板选中态同款规范），关闭 = 透明</summary>
        private void UpdateCopyDragModeVisual()
        {
            BorderStrokeSelectionClone.Background = isCopyDragMode
                ? CreateHighlightBrush(0x26)   //淡蓝底（15%），与各面板选中态同款
                : Brushes.Transparent;
        }

        /// <summary>退出复制拖拽模式（选区消失/切换工具时由各路径调用）</summary>
        private void ExitCopyDragMode()
        {
            if (!isCopyDragMode) return;
            isCopyDragMode = false;
            UpdateCopyDragModeVisual();
        }

        /// <summary>
        /// 复制拖拽会话中创建的图片副本（墨迹副本走 Strokes.Add 天然进撤销栈；
        /// 图片副本列表在此跟踪——拖动 Move 时移动副本而非原件，松手清空）。
        /// 与 StrokesSelectionClone（恒空的防御字段）不同，此列表真实参与拖动分流。
        /// </summary>
        List<System.Windows.Controls.Image> _imageDragClones = new List<System.Windows.Controls.Image>();

        /// <summary>
        /// 从选中内容拖出一份副本：完全重合复制（不偏移）→ 选中副本。
        /// 后续正常拖动路径移动的是副本，原件不动——"从图上拖出来"的手感即由此而来。
        /// 墨迹：StrokeCollection.Clone；图片：新建 Image 拷贝源/尺寸/翻转/位置并登记页表。
        /// </summary>
        private void CreateCopyDragClone()
        {
            var selected = inkCanvas.GetSelectedStrokes();
            var selectedImages = GetSelectedPageImages();
            if (selected.Count == 0 && selectedImages.Count == 0) return;

            //克隆图片：与原件完全重合（Source/尺寸/翻转/位置一致），登记到原件同一页
            var imageClones = new List<System.Windows.Controls.Image>();
            foreach (var src in selectedImages)
            {
                try
                {
                    var clone = new System.Windows.Controls.Image
                    {
                        Source = src.Source,
                        Width = src.Width,
                        Height = src.Height,
                        Stretch = src.Stretch,
                        IsHitTestVisible = src.IsHitTestVisible
                    };
                    //翻转态照搬（ScaleTransform 是可变对象，必须新建避免与原件共享）
                    if (src.RenderTransform is ScaleTransform sst && (sst.ScaleX != 1 || sst.ScaleY != 1))
                        clone.RenderTransform = new ScaleTransform(sst.ScaleX, sst.ScaleY);
                    InkCanvas.SetLeft(clone, InkCanvas.GetLeft(src));
                    InkCanvas.SetTop(clone, InkCanvas.GetTop(src));

                    //登记页表：跟随原件的页键（跨页拖拽副本仍属原图所在页）
                    int key = _imagePageKey[src];
                    if (!_pageImages.TryGetValue(key, out var list))
                    {
                        list = new List<System.Windows.Controls.Image>();
                        _pageImages[key] = list;
                    }
                    list.Add(clone);
                    _imagePageKey[clone] = key;

                    //图片本体拖动三件套（与 ImageLayer_AddImage 同款，见 MW_ImageLayer.cs）
                    clone.MouseLeftButtonDown += ImageLayer_DragMouseDown;
                    clone.MouseMove += ImageLayer_DragMouseMove;
                    clone.MouseLeftButtonUp += ImageLayer_DragMouseUp;

                    inkCanvas.Children.Add(clone);
                    clone.Visibility = src.Visibility;
                    imageClones.Add(clone);
                }
                catch { }
            }

            //清选中（双参：墨迹+图片一起）→ 加入墨迹副本 → 选中全部副本
            isProgramChangeStrokeSelection = true;
            inkCanvas.Select(new StrokeCollection(), new System.Collections.Generic.List<UIElement>());
            isProgramChangeStrokeSelection = false;

            StrokeCollection cloneStrokes = null;
            if (selected.Count > 0)
            {
                cloneStrokes = selected.Clone(); // 副本与原件完全重合，拖动时再分开
                inkCanvas.Strokes.Add(cloneStrokes); // StrokesChanged 自动进 TimeMachine（可撤销）
            }

            isProgramChangeStrokeSelection = true;
            inkCanvas.Select(cloneStrokes ?? new StrokeCollection(), imageClones);
            isProgramChangeStrokeSelection = false;

            _imageDragClones = imageClones; //后续 Move 拖的是副本（原件不动）
        }

        /// <summary>
        /// 把选中内容存入剪贴板（Ctrl+C / 复制按钮联动共用）。按选中内容分两路：
        /// 1. 纯墨迹 → ISF 序列化（"InkStrokes" 格式）：粘回画板仍是墨迹（可编辑/可擦/走撤销栈），
        ///    同时放 PNG（外部应用 PPT/微信 粘贴兼容；画板内读取优先 ISF）。
        /// 2. 含图片 → 整体渲染成透明底 PNG（墨迹+图片同框所见即所得），双格式入剪贴板。
        /// 【黑背景坑】Clipboard.SetImage 的 DIB 转换会丢 alpha——墨迹图大半是透明像素，
        /// 直接贴回来是黑底图。解法：DataObject 同时放 PNG 流（带 alpha，自己读）+ 标准位图。
        /// 复制失败不影响调用方主流程。
        /// </summary>
        private void CopySelectedStrokesToClipboard()
        {
            try
            {
                var strokes = inkCanvas.GetSelectedStrokes();
                var selectedImages = inkCanvas.GetSelectedElements()
                    .OfType<System.Windows.Controls.Image>().ToList();
                if (strokes.Count == 0 && selectedImages.Count == 0) return;

                if (selectedImages.Count == 0)
                {
                    // ---- 纯墨迹：ISF 序列化（核心）+ PNG（外部兼容）----
                    var dataObj = new DataObject();
                    using (var ms = new MemoryStream())
                    {
                        strokes.Save(ms); // ISF（Ink Serialized Format）：保留笔迹全部属性
                        dataObj.SetData("InkStrokes", new MemoryStream(ms.ToArray()), false);
                    }
                    var rtb = RenderStrokesAndImages(strokes, selectedImages);
                    if (rtb != null)
                    {
                        dataObj.SetImage(rtb); // 标准位图（外部应用读）
                        using (var ms = new MemoryStream())
                        {
                            var enc = new PngBitmapEncoder();
                            enc.Frames.Add(BitmapFrame.Create(rtb));
                            enc.Save(ms);
                            dataObj.SetData("PNG", new MemoryStream(ms.ToArray()), false);
                        }
                    }
                    Clipboard.SetDataObject(dataObj, true);
                    return;
                }

                // ---- 含图片：整体渲染 PNG ----
                var bmp = RenderStrokesAndImages(strokes, selectedImages);
                if (bmp == null) return;
                var dataObj2 = new DataObject();
                dataObj2.SetImage(bmp);
                using (var ms = new MemoryStream())
                {
                    var enc = new PngBitmapEncoder();
                    enc.Frames.Add(BitmapFrame.Create(bmp));
                    enc.Save(ms);
                    dataObj2.SetData("PNG", new MemoryStream(ms.ToArray()), false);
                }
                Clipboard.SetDataObject(dataObj2, true);
            }
            catch { /* 静默：复制是附加动作，失败不干扰主流程 */ }
        }

        /// <summary>把墨迹(+图片)渲染成透明底位图（所见即所得：原尺寸、含 alpha）</summary>
        private RenderTargetBitmap RenderStrokesAndImages(StrokeCollection strokes, List<System.Windows.Controls.Image> images)
        {
            // 统一包围盒 = 墨迹 ∪ 图片
            Rect bounds = Rect.Empty;
            if (strokes.Count > 0) bounds = strokes.GetBounds();
            foreach (var img in images)
            {
                var r = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                bounds = bounds.IsEmpty ? r : Rect.Union(bounds, r);
            }
            if (bounds.IsEmpty || bounds.Width < 1 || bounds.Height < 1) return null;

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.PushTransform(new TranslateTransform(-bounds.X, -bounds.Y));
                foreach (Stroke s in strokes) s.Draw(dc);
                foreach (var img in images) // 图片按显示尺寸原样画入
                    dc.DrawImage(img.Source, new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height));
                dc.Pop();
            }
            var rtb = new RenderTargetBitmap(
                (int)Math.Ceiling(bounds.Width), (int)Math.Ceiling(bounds.Height),
                96, 96, System.Windows.Media.PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        /// <summary>
        /// 从剪贴板读图片：优先取 "PNG" 格式（CopySelectedStrokesToClipboard 写入，
        /// 带 alpha 透明通道——墨迹图不黑底）；没有再退回 GetImage（截图/外部来源）。
        /// </summary>
        private BitmapSource TryGetClipboardImage()
        {
            try
            {
                if (Clipboard.ContainsData("PNG"))
                {
                    var stream = Clipboard.GetData("PNG") as Stream;
                    if (stream != null)
                    {
                        var bmp = new BitmapImage();
                        bmp.BeginInit();
                        bmp.CacheOption = BitmapCacheOption.OnLoad;
                        bmp.StreamSource = stream;
                        bmp.EndInit();
                        bmp.Freeze();
                        if (bmp.PixelWidth > 0) return bmp;
                    }
                }
                if (Clipboard.ContainsImage()) return Clipboard.GetImage();
            }
            catch { }
            return null;
        }

        /// <summary>
        /// 从剪贴板读墨迹（ISF 格式，纯墨迹复制时写入）。
        /// 粘回画板仍是墨迹（可编辑/可擦/可撤销）——这是"纯墨迹复制"与"截图"的本质区别。
        /// </summary>
        private StrokeCollection TryGetClipboardStrokes()
        {
            try
            {
                if (!Clipboard.ContainsData("InkStrokes")) return null;
                var stream = Clipboard.GetData("InkStrokes") as Stream;
                if (stream == null) return null;
                var sc = new StrokeCollection(stream); // 反序列化 ISF
                return sc.Count > 0 ? sc : null;
            }
            catch { return null; }
        }


        private void BorderStrokeSelectionDelete_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ExitCopyDragMode(); // 复制模式随选区一起结束

            // 图片优先删（ImageLayer 内部自清页表并同步选中集合），再删墨迹——
            // 混合选中时两者一起删除
            ImageLayer_DeleteSelectedImages();

            // 只删除选中墨迹（原实现转调 SymbolIconDelete_MouseUp，而它已被改为转调
            // 滑动清屏——选中后点删除会把整页墨迹全清掉，bug 根因即在此）
            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count > 0)
            {
                // StrokesChanged 会把 Removed 记进 TimeMachine（Ctrl+Z 可撤销）
                inkCanvas.Strokes.Remove(strokes);
            }

            //清空选中（双参：元素+墨迹）。挡板拦住 SelectionChanged（不触发快照副作用），
            //但收尾必须手动收起覆盖层——操作条绑定覆盖层可见性，不收就残留在原地。
            //纯墨迹删除无此问题：Strokes.Remove 在挡板前先触发了 SelectionChanged 自动收起
            isProgramChangeStrokeSelection = true;
            inkCanvas.Select(new StrokeCollection(), new System.Collections.Generic.List<UIElement>());
            isProgramChangeStrokeSelection = false;
            GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
            TryEndOneShotSelection();
        }

        private void GridPenWidthDecrease_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ChangeStrokeThickness(0.8);
        }

        private void GridPenWidthIncrease_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            ChangeStrokeThickness(1.25);
        }

        private void ChangeStrokeThickness(double multipler)
        {
            foreach (Stroke stroke in inkCanvas.GetSelectedStrokes())
            {
                //stroke.DrawingAttributes.Width *= 1.25;
                //stroke.DrawingAttributes.Height *= 1.25;

                var newWidth = stroke.DrawingAttributes.Width * multipler;
                var newHeight = stroke.DrawingAttributes.Height * multipler;

                if (newWidth >= DrawingAttributes.MinWidth && newWidth <= DrawingAttributes.MaxWidth
                    && newHeight >= DrawingAttributes.MinHeight && newHeight <= DrawingAttributes.MaxHeight)
                {
                    stroke.DrawingAttributes.Width = newWidth;
                    stroke.DrawingAttributes.Height = newHeight;
                }
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

        private void ImageFlipHorizontal_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(GetGestureSelectionBounds().Left + GetGestureSelectionBounds().Width / 2,
                GetGestureSelectionBounds().Top + GetGestureSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.ScaleAt(-1, 1, center.X, center.Y);  // 缩放

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }

            //图片水平镜像：RenderTransform 反转 X（中心点翻转，不影响布局尺寸，
            //原生选框天然贴合）。连续点两次 = 翻回原样，符合翻转按钮的通用语义
            foreach (var img in GetSelectedPageImages()) FlipImage(img, flipHorizontal: true);

            if (DrawingAttributesHistory.Count > 0)
            {
                var collecion = new StrokeCollection();
                foreach (var item in DrawingAttributesHistory)
                {
                    collecion.Add(item.Key);
                }
                timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                foreach (var item in DrawingAttributesHistoryFlag)
                {
                    item.Value.Clear();
                }
            }
            //updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>
        /// 适配白板宽度（操作条按钮，仅选中含图片时可见）：
        /// 选中图片宽度 = 画布可视宽度，高度等比缩放，水平居中，垂直位置保持不动（板书习惯：左右铺满、上下不跳）。
        /// 多张图片时按各自中心对齐铺满。恢复走"统一还原"键（快照在选中时记录，含适配前尺寸）。
        /// </summary>
        private void GridSelectionFitWidth_MouseUp(object sender, MouseButtonEventArgs e)
        {
            try
            {
                //只认登记过的页面图片（排除其他来源元素），纯墨迹点击无效果
                foreach (var img in GetSelectedPageImages())
                {
                    if (img.Width <= 0 || img.Height <= 0) continue;

                    //目标宽度=画布可视宽度；高度按图片自身宽高比等比缩放
                    double targetW = inkCanvas.ActualWidth;
                    double ratio = img.Height / img.Width;
                    double targetH = targetW * ratio;

                    //垂直位置保持原中心（上下不跳动），水平居中
                    double cy = InkCanvas.GetTop(img) + img.Height / 2;
                    InkCanvas.SetLeft(img, (inkCanvas.ActualWidth - targetW) / 2); //宽度=画布宽时结果恒 0，显式写出语义
                    InkCanvas.SetTop(img, cy - targetH / 2);
                    img.Width = targetW;
                    img.Height = targetH;
                }
                updateBorderStrokeSelectionControlLocation(); //选区变大，操作条/旋转钮/手柄跟随
            }
            catch { }
        }

        /// <summary>
        /// 翻转图片（渲染层镜像，不动布局）：把已有的 ScaleTransform 的 X/Y 乘 -1 切换翻转态。
        /// 布局尺寸不变（选框贴合），翻转态持续到下次翻转或还原按钮恢复位置尺寸。
        /// </summary>
        private void FlipImage(System.Windows.Controls.Image img, bool flipHorizontal)
        {
            try
            {
                var st = img.RenderTransform as ScaleTransform;
                if (st == null)
                {
                    //首次翻转：中心点缩放（渲染中心=布局中心，天然镜像不位移）
                    st = new ScaleTransform(flipHorizontal ? -1 : 1, flipHorizontal ? 1 : -1);
                    img.RenderTransform = st;
                    img.RenderTransformOrigin = new Point(0.5, 0.5);
                }
                else
                {
                    if (flipHorizontal) st.ScaleX *= -1; else st.ScaleY *= -1;
                }
            }
            catch { }
        }

        private void ImageFlipVertical_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            Matrix m = new Matrix();

            // Find center of element and then transform to get current location of center
            FrameworkElement fe = e.Source as FrameworkElement;
            Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
            center = new Point(GetGestureSelectionBounds().Left + GetGestureSelectionBounds().Width / 2,
                GetGestureSelectionBounds().Top + GetGestureSelectionBounds().Height / 2);
            center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

            // Update matrix to reflect translation/rotation
            m.ScaleAt(1, -1, center.X, center.Y);  // 缩放

            StrokeCollection targetStrokes = inkCanvas.GetSelectedStrokes();
            foreach (Stroke stroke in targetStrokes)
            {
                stroke.Transform(m, false);
            }

            //图片垂直镜像（与水平翻转同一套 RenderTransform 逻辑）
            foreach (var img in GetSelectedPageImages()) FlipImage(img, flipHorizontal: false);
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

        #endregion


        bool isGridInkCanvasSelectionCoverMouseDown = false;
        StrokeCollection StrokesSelectionClone = new StrokeCollection();

        /// <summary>取当前选中元素里的页面图片（已登记页表的）——混合选中（墨迹+图片）拖动路径共用</summary>
        private List<System.Windows.Controls.Image> GetSelectedPageImages()
        {
            // 只认 ImageLayer 登记过的图片，排除其他来源的元素
            return inkCanvas.GetSelectedElements()
                .OfType<System.Windows.Controls.Image>()
                .Where(img => _imagePageKey.ContainsKey(img))
                .ToList();
        }

        /// <summary>
        /// 手势几何统一入口：选中内容包围盒（inkCanvas 坐标系）= 墨迹 bounds ∪ 选中页面图片 rect。
        /// 【为什么不用 InkCanvas.GetSelectionBounds()】它对"元素"的 bounds 由选区装饰器经
        /// LayoutUpdated 异步跟踪（WPF 源码 InkCanvasSelection：元素入选时走 UpdateSelectionAdorner），
        /// SelectionChanged 同步时刻常拿到空/旧值——纯图片选中时整个覆盖层命中区被清零、
        /// 手柄判定/缩放/操作条定位全部失效。自算版本完全同步，纯墨迹/纯图片/混合三态一致。
        /// </summary>
        private Rect GetGestureSelectionBounds()
        {
            Rect b = Rect.Empty;
            var strokes = inkCanvas.GetSelectedStrokes();
            if (strokes.Count > 0) b = strokes.GetBounds();
            foreach (var img in GetSelectedPageImages())
            {
                var r = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                b = b.IsEmpty ? r : Rect.Union(b, r);
            }
            return b;
        }

        //鼠标拖动选区状态（触摸走 Manipulation 事件，鼠标/数位笔走此路径；dec>0 表示触摸进行中，让位）
        bool isMouseSelectionDragging = false;
        bool hasMouseSelectionDragMoved = false; // 按下后是否真的移动过（区分"拖动"与"单击取消选区"）
        Point lastMousePointOnSelectionCover = new Point(0, 0);

        private void GridInkCanvasSelectionCover_MouseDown(object sender, MouseButtonEventArgs e)
        {
            isGridInkCanvasSelectionCoverMouseDown = true;
            if (e.ChangedButton != MouseButton.Left || dec.Count != 0) return;

            //缩放手柄命中：走手柄拖动路径（InkCanvas 原生把手被本覆盖层遮挡，此处实现同语义的拖动缩放）
            var handlePos = e.GetPosition(inkCanvas);
            var handle = HitTestSelectionHandle(handlePos);
            if (handle != SelectionHandleKind.None)
            {
                StartHandleDrag(handle, handlePos);
                try { GridInkCanvasSelectionCover.CaptureMouse(); } catch { }
                return;
            }

            // 复制模式的副本在框内按下时才创建（见下方 isCopyDragMode 分支），此处不再预创建

            //【框内才能拖】与 PowerPoint 等全行业一致：按下点在选中笔迹包围盒内（含 10px 容差，
            //方便抓细线）才进入拖动；框外按下不遥控选中物——想在别处落笔时不会误拽走整个图形
            var dragBounds = inkCanvas.GetSelectedStrokes().GetBounds();
            // 混合选中（墨迹+图片）：命中区并入图片包围盒——否则按在图片本体上会被判"框外"
            // 直接取消选中，图片区域就拖不动了（与"抓本体即可拖"的通用交互不符）
            foreach (var img in GetSelectedPageImages())
            {
                dragBounds.Union(new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height));
            }
            dragBounds.Inflate(10, 10);
            if (!dragBounds.Contains(handlePos))
            {
                //框外按下：结束本层拖动意图，放行事件让 MouseUp 走"单击取消选中"路径
                isMouseSelectionDragging = false;
                return;
            }

            // 复制拖拽模式：框内按下先拖出一份副本并选中它，
            // 之后正常拖动路径移动的是副本（原件不动）——"从图上拖出来"的核心
            if (isCopyDragMode) CreateCopyDragClone();

            //开始鼠标拖动（捕获鼠标保证移出窗口也能收到 Move/Up）
            lastMousePointOnSelectionCover = e.GetPosition(null);
            isMouseSelectionDragging = true;
            hasMouseSelectionDragMoved = false;
            try { GridInkCanvasSelectionCover.CaptureMouse(); } catch { }
        }

        private void GridInkCanvasSelectionCover_MouseMove(object sender, MouseEventArgs e)
        {
            if (!isMouseSelectionDragging) return;
            if (e.LeftButton != MouseButtonState.Pressed) { FinishMouseSelectionDrag(); return; }

            //手柄拖动缩放路径（优先于整体平移）
            if (_activeHandleDragKind != SelectionHandleKind.None)
            {
                UpdateHandleDrag(e.GetPosition(inkCanvas));
                TryCollapseBarForDrag(); // 缩放拖动中：操作条临时收成＋小圆钮
                return;
            }

            var pos = e.GetPosition(null);
            var dx = pos.X - lastMousePointOnSelectionCover.X;
            var dy = pos.Y - lastMousePointOnSelectionCover.Y;
            lastMousePointOnSelectionCover = pos;
            if (Math.Abs(dx) < 0.1 && Math.Abs(dy) < 0.1) return;
            hasMouseSelectionDragMoved = true; // 超过阈值的移动才算拖动
            TryCollapseBarForDrag(); // 平移拖动中：操作条临时收成＋小圆钮

            //与触摸路径一致：克隆时拖动副本，否则拖动选中墨迹
            StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
            if (StrokesSelectionClone.Count != 0) strokes = StrokesSelectionClone;

            var m = new Matrix();
            m.Translate(dx, dy);
            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false);
            }

            //图片平移：复制模式拖动时移动的是副本（原件不动），普通拖动移动选中原件
            //（不进 TimeMachine，与图片本体拖动口径一致）
            if (_imageDragClones.Count > 0)
            {
                foreach (var img in _imageDragClones)
                {
                    InkCanvas.SetLeft(img, InkCanvas.GetLeft(img) + dx);
                    InkCanvas.SetTop(img, InkCanvas.GetTop(img) + dy);
                }
            }
            else
            {
                foreach (var img in GetSelectedPageImages())
                {
                    InkCanvas.SetLeft(img, InkCanvas.GetLeft(img) + dx);
                    InkCanvas.SetTop(img, InkCanvas.GetTop(img) + dy);
                }
            }

            //克隆拖动时选区跟副本走（副本被选中），控制条需跟随；
            //普通拖动选区即选中物，同样跟随
            updateBorderStrokeSelectionControlLocation();
        }

        private void GridInkCanvasSelectionCover_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (!isGridInkCanvasSelectionCoverMouseDown) return;
            isGridInkCanvasSelectionCoverMouseDown = false;

            bool wasHandleDrag = _activeHandleDragKind != SelectionHandleKind.None;

            if (isMouseSelectionDragging && hasMouseSelectionDragMoved)
            {
                //真实拖动结束（平移或手柄缩放）：保持选区可见，提交撤销历史
                FinishMouseSelectionDrag();
            }
            else
            {
                //结束拖动状态；手柄上原地点击未拖动 → 保持选区（空白处单击才取消选区）
                if (isMouseSelectionDragging) FinishMouseSelectionDrag();
                if (wasHandleDrag) return;
                //单击（未拖动）：取消选区（原行为）。元素集合一并传空——
                //纯图片/混合选中时图片也同时取消，避免"墨迹没了图片还挂着"的错乱
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection(), new System.Collections.Generic.List<UIElement>());
                isProgramChangeStrokeSelection = false;
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                ExitCopyDragMode(); // 选中没了，复制模式一并结束
                //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                TryEndOneShotSelection();
            }
        }

        /// <summary>结束鼠标拖动：释放捕获、清空克隆引用、提交撤销历史（同触摸路径 ManipulationCompleted）</summary>
        private void FinishMouseSelectionDrag()
        {
            isMouseSelectionDragging = false;
            _activeHandleDragKind = SelectionHandleKind.None; // 结束手柄缩放会话
            TryRestoreBarAfterDrag(); // 拖动结束：操作条恢复展开态
            try { if (GridInkCanvasSelectionCover.IsMouseCaptured) GridInkCanvasSelectionCover.ReleaseMouseCapture(); } catch { }
            StrokesSelectionClone = new StrokeCollection();
            _imageDragClones = new List<System.Windows.Controls.Image>(); //图片副本会话结束（副本已落定，正常选中它）

            if (StrokeManipulationHistory?.Count > 0)
            {
                timeMachine.CommitStrokeManipulationHistory(StrokeManipulationHistory);
                foreach (var item in StrokeManipulationHistory)
                {
                    StrokeInitialHistory[item.Key] = item.Value.Item2;
                }
                StrokeManipulationHistory = null;
            }

            //手柄缩放会同步改笔画粗细：结束时一并提交（平移路径不产生该历史，无副作用）
            CommitDrawingAttributesHistoryNow();
        }

        #region 选区手柄拖动缩放 + Ctrl+滚轮缩放

        /// <summary>选区缩放手柄类型：4 角 + 4 边中点（与 InkCanvas 原生把手位置一致）+ 顶部旋转钮</summary>
        private enum SelectionHandleKind
        {
            None,
            TopLeft, Top, TopRight, Right, BottomRight, Bottom, BottomLeft, Left,
            Rotate //顶部中央伸出的旋转手柄（PowerPoint/Figma 惯例位置）
        }

        SelectionHandleKind _activeHandleDragKind = SelectionHandleKind.None;
        Rect _handleDragStartBounds = Rect.Empty;
        double _rotateDragStartAngle = 0.0; //旋转拖动开始时鼠标相对选区中心的方位角（弧度）

        /// <summary>旋转钮在选中框顶边上方伸出的距离（DIP）</summary>
        const double RotateHandleOffset = 22.0;

        /// <summary>手柄命中半径（DIP）。InkCanvas 原生把手视觉直径约 8，放宽到 14 便于鼠标/数位笔抓取</summary>
        const double SelectionHandleHitRadius = 14.0;

        /// <summary>
        /// 命中检测：判断 pos（inkCanvas 坐标）是否落在选区缩放手柄（小圆点）上。
        /// InkCanvas 原生把手被本覆盖层遮挡无法拖动，这里在覆盖层实现同语义交互。
        /// </summary>
        private SelectionHandleKind HitTestSelectionHandle(Point pos)
        {
            Rect b = GetGestureSelectionBounds();
            if (b.IsEmpty || b.Width <= 0 || b.Height <= 0) return SelectionHandleKind.None;

            //旋转钮：顶部中央向上伸出（优先于缩放手柄判定，位置在框外不冲突）
            var rotateCenter = new Point(b.Left + b.Width / 2, b.Top - RotateHandleOffset);
            if ((pos - rotateCenter).Length <= SelectionHandleHitRadius) return SelectionHandleKind.Rotate;

            var candidates = new Tuple<SelectionHandleKind, Point>[]
            {
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.TopLeft,      new Point(b.Left, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Top,         new Point(b.Left + b.Width / 2, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.TopRight,    new Point(b.Right, b.Top)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Right,       new Point(b.Right, b.Top + b.Height / 2)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.BottomRight, new Point(b.Right, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Bottom,      new Point(b.Left + b.Width / 2, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.BottomLeft,  new Point(b.Left, b.Bottom)),
                new Tuple<SelectionHandleKind, Point>(SelectionHandleKind.Left,        new Point(b.Left, b.Top + b.Height / 2)),
            };

            SelectionHandleKind best = SelectionHandleKind.None;
            double bestDist = SelectionHandleHitRadius;
            foreach (var c in candidates)
            {
                double d = (pos - c.Item2).Length;
                if (d <= bestDist) { best = c.Item1; bestDist = d; }
            }
            return best;
        }

        /// <summary>
        /// 开始手柄拖动：记录起始选区 bounds。复用 isMouseSelectionDragging 状态，
        /// 使 Stroke_StylusPointsChanged 在拖动期间只累计撤销历史、松开时一次提交
        /// （否则一次拖动会碎片化为多个撤销步骤）。
        /// </summary>
        private void StartHandleDrag(SelectionHandleKind kind, Point pos)
        {
            _activeHandleDragKind = kind;
            _handleDragStartBounds = GetGestureSelectionBounds();
            isMouseSelectionDragging = true;
            hasMouseSelectionDragMoved = false;
            lastMousePointOnSelectionCover = pos;

            //旋转分支：记录初始方位角（鼠标 → 选区中心的方向），后续按角度增量旋转
            if (kind == SelectionHandleKind.Rotate)
            {
                var b = _handleDragStartBounds;
                var center = new Point(b.Left + b.Width / 2, b.Top + b.Height / 2);
                _rotateDragStartAngle = Math.Atan2(pos.Y - center.Y, pos.X - center.X);
            }
        }

        /// <summary>
        /// 手柄拖动中：以起始 bounds 的对角/对边为固定锚点，把选区缩放到鼠标当前位置。
        /// 角点手柄 = 等比缩放（保持宽高比，手写批注多为文字/图形，自由拉伸易变形），
        /// 边中点手柄 = 单轴缩放（需要"只拉宽/压扁"时用，与 PPT 语义一致）。
        /// 每次移动按"当前 bounds → 目标尺寸"计算增量比例，围绕同一锚点复合即得总变换。
        /// </summary>
        private void UpdateHandleDrag(Point pos)
        {
            if (_activeHandleDragKind == SelectionHandleKind.None) return;
            Rect start = _handleDragStartBounds;
            if (start.IsEmpty || start.Width < 0.5 || start.Height < 0.5) return;

            //旋转手柄分支：按"鼠标绕选区中心的方位角增量"整体旋转选中笔迹。
            //旋转中心固定取拖动起始 bounds 的中心（整次拖动不漂移）；
            //角度跨过 ±180° 射线时按最短方向修正，避免图形瞬间反向猛转。
            if (_activeHandleDragKind == SelectionHandleKind.Rotate)
            {
                StrokeCollection rotStrokes = inkCanvas.GetSelectedStrokes();
                if (rotStrokes.Count == 0) return;

                var center = new Point(start.Left + start.Width / 2, start.Top + start.Height / 2);
                double currentAngle = Math.Atan2(pos.Y - center.Y, pos.X - center.X); //鼠标当前方位角（弧度）
                double delta = currentAngle - _rotateDragStartAngle;                    //相对上次的角度增量
                _rotateDragStartAngle = currentAngle;                                  //滚动累计，供下一帧
                if (delta > Math.PI) delta -= 2 * Math.PI;   //跨线修正：+180° 方向跳变 → 取最短转向
                if (delta < -Math.PI) delta += 2 * Math.PI;
                if (Math.Abs(delta) < 0.001) return;          //忽略亚像素级微动，省一次矩阵变换

                var rm = new Matrix();
                rm.RotateAt(delta * 180.0 / Math.PI, center.X, center.Y); //弧度 → 角度，绕中心旋转
                foreach (Stroke stroke in rotStrokes)
                {
                    stroke.Transform(rm, false); //StylusPointsChanged 自动累计撤销历史（拖动期间不提交）
                }

                hasMouseSelectionDragMoved = true;
                updateBorderStrokeSelectionControlLocation(); //旋转钮/工具条跟随新选区
                return;
            }

            //固定锚点 = 被拖动手柄的对角（角点）或对边（边中点），整次拖动保持不动
            double anchorX = 0, anchorY = 0;
            bool isCorner = false, dragX = false, dragY = false;
            switch (_activeHandleDragKind)
            {
                case SelectionHandleKind.TopLeft:      anchorX = start.Right; anchorY = start.Bottom; isCorner = true; break;
                case SelectionHandleKind.TopRight:     anchorX = start.Left;  anchorY = start.Bottom; isCorner = true; break;
                case SelectionHandleKind.BottomRight:  anchorX = start.Left;  anchorY = start.Top;    isCorner = true; break;
                case SelectionHandleKind.BottomLeft:   anchorX = start.Right; anchorY = start.Top;    isCorner = true; break;
                case SelectionHandleKind.Top:          anchorY = start.Bottom; dragY = true; break;
                case SelectionHandleKind.Bottom:       anchorY = start.Top;    dragY = true; break;
                case SelectionHandleKind.Left:         anchorX = start.Right;  dragX = true; break;
                case SelectionHandleKind.Right:        anchorX = start.Left;   dragX = true; break;
                default: return;
            }

            //目标尺寸 = 鼠标到锚点距离（下限 2 DIP，防翻转/退化），按当前 bounds 换算增量比例
            const double MinSize = 2.0;
            Rect cur = GetGestureSelectionBounds();
            if (cur.IsEmpty || cur.Width < 0.5 || cur.Height < 0.5) return;

            double fx = 1.0, fy = 1.0;
            if (isCorner)
            {
                //等比：统一因子 = 鼠标到锚点距离 / 当前被拖角到锚点距离（沿对角线跟手）
                Point curCorner;
                switch (_activeHandleDragKind)
                {
                    case SelectionHandleKind.TopLeft:     curCorner = new Point(cur.Left, cur.Top); break;
                    case SelectionHandleKind.TopRight:    curCorner = new Point(cur.Right, cur.Top); break;
                    case SelectionHandleKind.BottomRight: curCorner = new Point(cur.Right, cur.Bottom); break;
                    default:                              curCorner = new Point(cur.Left, cur.Bottom); break;
                }
                double dCur = (curCorner - new Point(anchorX, anchorY)).Length;
                if (dCur < 0.5) return;
                double f = (pos - new Point(anchorX, anchorY)).Length / dCur;
                if (f < 0.02) f = 0.02; //防缩至消失
                fx = fy = f;
            }
            else
            {
                if (dragX) fx = Math.Max(MinSize, Math.Abs(pos.X - anchorX)) / cur.Width;
                if (dragY) fy = Math.Max(MinSize, Math.Abs(pos.Y - anchorY)) / cur.Height;
            }
            if (Math.Abs(fx - 1) < 0.001 && Math.Abs(fy - 1) < 0.001) return;

            StrokeCollection strokes = inkCanvas.GetSelectedStrokes();

            var m = new Matrix();
            m.ScaleAt(fx, fy, anchorX, anchorY);
            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false); //StylusPointsChanged 自动累计撤销历史（拖动期间不提交）
                try
                {
                    //笔画粗细随选区同步缩放（与触摸双指缩放一致），夹取到 WPF 允许范围
                    double w = Math.Max(DrawingAttributes.MinWidth, Math.Min(DrawingAttributes.MaxWidth, stroke.DrawingAttributes.Width * fx));
                    double h = Math.Max(DrawingAttributes.MinHeight, Math.Min(DrawingAttributes.MaxHeight, stroke.DrawingAttributes.Height * fy));
                    stroke.DrawingAttributes.Width = w;
                    stroke.DrawingAttributes.Height = h;
                }
                catch { }
            }

            //图片同步缩放（与墨迹同一锚点同一因子；翻转态 RenderTransform 不动，
            //布局宽高缩放即可）。图片操作不进撤销栈，还原按钮负责恢复
            foreach (var img in GetSelectedPageImages())
                ScaleImage(img, fx, fy, anchorX, anchorY);

            hasMouseSelectionDragMoved = true;
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>
        /// 围绕锚点缩放单张图片：位置与尺寸一起变换（Left/Top 绕锚点缩放 + 宽高乘因子），
        /// 与墨迹 stroke.Transform(ScaleAt) 同几何，混合选中时两者保持同步形变。
        /// </summary>
        private void ScaleImage(System.Windows.Controls.Image img, double fx, double fy, double anchorX, double anchorY)
        {
            try
            {
                double left = InkCanvas.GetLeft(img), top = InkCanvas.GetTop(img);
                InkCanvas.SetLeft(img, anchorX + (left - anchorX) * fx);
                InkCanvas.SetTop(img, anchorY + (top - anchorY) * fy);
                img.Width = Math.Max(2.0, img.Width * fx);   //下限 2 DIP，与手柄拖动的 MinSize 一致
                img.Height = Math.Max(2.0, img.Height * fy);
            }
            catch { }
        }

        /// <summary>悬停反馈：手柄上显示缩放箭头；框内显示移动光标（四向箭头）；框外恢复默认</summary>
        private void GridInkCanvasSelectionCover_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            if (isMouseSelectionDragging || dec.Count != 0) return; // 拖动/触摸进行中保持当前光标
            //判定含图片：纯图片选中同样有手柄缩放/移动的光标反馈
            if (inkCanvas.GetSelectedStrokes().Count == 0 && GetSelectedPageImages().Count == 0) return;

            var pos = e.GetPosition(inkCanvas);

            //优先手柄（缩放光标）
            var handleCursor = SelectionHandleCursor(HitTestSelectionHandle(pos));
            if (handleCursor != null)
            {
                GridInkCanvasSelectionCover.Cursor = handleCursor;
                return;
            }

            //框内（与 MouseDown 的拖动判定同一套几何：包围盒 + 10px 容差）→ 四向移动光标
            //光标可供性：告诉用户"这里按住可以拖走"，和实际能拖的范围严格一致
            var bounds = inkCanvas.GetSelectedStrokes().GetBounds();
            // 混合选中：图片包围盒并入（与 MouseDown 的拖动命中区同几何，悬停在哪就能拖哪）
            foreach (var img in GetSelectedPageImages())
            {
                bounds.Union(new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height));
            }
            bounds.Inflate(10, 10);
            GridInkCanvasSelectionCover.Cursor = bounds.Contains(pos) ? Cursors.SizeAll : null;
        }

        private Cursor SelectionHandleCursor(SelectionHandleKind kind)
        {
            switch (kind)
            {
                case SelectionHandleKind.TopLeft:
                case SelectionHandleKind.BottomRight:
                    return Cursors.SizeNWSE;
                case SelectionHandleKind.TopRight:
                case SelectionHandleKind.BottomLeft:
                    return Cursors.SizeNESW;
                case SelectionHandleKind.Top:
                case SelectionHandleKind.Bottom:
                    return Cursors.SizeNS;
                case SelectionHandleKind.Left:
                case SelectionHandleKind.Right:
                    return Cursors.SizeWE;
                case SelectionHandleKind.Rotate:
                    return Cursors.ScrollAll; //旋转手柄（圆圈箭头，最接近旋转语义的系统光标）
                default:
                    return null; // 非手柄区域恢复默认光标
            }
        }

        /// <summary>Ctrl+滚轮：缩放批注（有选区缩放选区，无选区缩放整屏；与 Ctrl+加减号同语义，可撤销）</summary>
        private void GridInkCanvasSelectionCover_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            // Ctrl+滚轮 = 缩放当前选区（原有行为）
            if ((Keyboard.Modifiers & ModifierKeys.Control) == ModifierKeys.Control)
            {
                if (Math.Abs(e.Delta) >= 1 && ScaleAllOrSelection(e.Delta > 0 ? 1.1 : 0.9)) e.Handled = true;
                return;
            }

            // 非 Ctrl：滚轮要滚动笔记——必须在这里转发。
            // 原因：选区命中区（BorderSelectionHitArea，覆盖选中框附近）盖在 inkCanvas 之上，
            // 且两者是"兄弟节点"（都在同一个外层 Grid 里）→ 滚轮事件目标是命中区时，
            // inkCanvas 不在冒泡路径上，InkCanvas_PreviewMouseWheel 永远收不到。
            // 表现就是"鼠标落在选区内滚轮没反应、落在选区外又正常"这种时灵时不灵的怪象
            // （滚动胶囊的滑块也因此停着不动，看着像"找不到"）。
            if (!IsNoteScrollActive) return;
            var notches = e.Delta / 120.0;
            if (Math.Abs(notches) < 0.01) return;
            ScrollNote(-notches * NoteScrollWheelStep);
            e.Handled = true;
        }

        #endregion 选区手柄拖动缩放 + Ctrl+滚轮缩放

        #region 选中缩放/还原

        //选中快照：SelectionChanged 捕获（还原 = 恢复到本次选中时的状态）
        Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>> SelectionSnapshot;

        /// <summary>图片选中快照：还原按钮用（恢复到选中时的位置/尺寸/翻转态）。结构简单直接存字段</summary>
        private class ImageSnapshot
        {
            public double Left, Top, Width, Height;
            public double ScaleX, ScaleY; //渲染层翻转态
        }
        Dictionary<System.Windows.Controls.Image, ImageSnapshot> _imageSelectionSnapshot =
            new Dictionary<System.Windows.Controls.Image, ImageSnapshot>();

        /// <summary>
        /// 统一缩放入口（Ctrl+滚轮 / Ctrl+加减号）：
        /// 有选区 → 缩放选中批注（绕选区中心）；无选区 → 缩放当前屏幕全部批注（绕屏幕中心）。
        /// 返回是否执行了缩放（无目标时 false，调用方决定是否放行事件）。
        /// </summary>
        private bool ScaleAllOrSelection(double factor)
        {
            if (inkCanvas.Visibility != Visibility.Visible) return false;

            var strokes = inkCanvas.GetSelectedStrokes();
            var selectedImages = GetSelectedPageImages();
            if (strokes.Count > 0 || selectedImages.Count > 0)
            {
                Rect bounds = GetGestureSelectionBounds();
                var center = new Point(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2);
                if (strokes.Count > 0 && ScaleStrokes(strokes, center, factor)) { }
                //图片绕选区中心同步缩放（混合选中时与墨迹同中心同因子）
                foreach (var img in selectedImages) ScaleImage(img, factor, factor, center.X, center.Y);
                updateBorderStrokeSelectionControlLocation();
                return true;
            }

            //无选区：整屏批注绕屏幕中心缩放
            if (inkCanvas.Strokes.Count == 0) return false;
            return ScaleStrokes(inkCanvas.Strokes,
                new Point(inkCanvas.ActualWidth / 2, inkCanvas.ActualHeight / 2), factor);
        }

        /// <summary>缩放一批笔画：点坐标绕 center 缩放、笔画粗细同步（夹取 WPF 范围）、提交撤销历史</summary>
        private bool ScaleStrokes(StrokeCollection strokes, Point center, double factor)
        {
            if (strokes == null || strokes.Count == 0) return false;

            //防误触缩没：已经小到 1 DIP 以内不再继续缩小
            if (factor < 1)
            {
                var b = strokes.GetBounds();
                if (b.Width < 1 && b.Height < 1) return false;
            }

            var m = new Matrix();
            m.ScaleAt(factor, factor, center.X, center.Y);

            foreach (Stroke stroke in strokes)
            {
                stroke.Transform(m, false); //触发 StylusPointsChanged，自动进撤销历史
                try
                {
                    double w = stroke.DrawingAttributes.Width * factor;
                    double h = stroke.DrawingAttributes.Height * factor;
                    //超出 WPF 允许范围则夹取（点坐标照常缩放，仅笔宽受限）
                    w = Math.Max(DrawingAttributes.MinWidth, Math.Min(DrawingAttributes.MaxWidth, w));
                    h = Math.Max(DrawingAttributes.MinHeight, Math.Min(DrawingAttributes.MaxHeight, h));
                    stroke.DrawingAttributes.Width = w;
                    stroke.DrawingAttributes.Height = h;
                }
                catch { }
            }

            CommitDrawingAttributesHistoryNow();
            return true;
        }

        private void GridSelectionScaleRestore_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            var strokes = inkCanvas.GetSelectedStrokes();
            if ((strokes.Count == 0 && _imageSelectionSnapshot.Count == 0)
                || (strokes.Count > 0 && SelectionSnapshot == null && _imageSelectionSnapshot.Count == 0)) return;

            var history = new Dictionary<Stroke, Tuple<StylusPointCollection, StylusPointCollection>>();
            if (SelectionSnapshot != null)
            {
                foreach (Stroke s in strokes)
                {
                    if (!SelectionSnapshot.TryGetValue(s, out var snap)) continue;
                    try
                    {
                        var oldPts = s.StylusPoints.Clone();
                        s.StylusPoints = snap.Item1; //赋值触发 StylusPointsReplaced，不自动提交历史，下面手动提交
                        s.DrawingAttributes.Width = snap.Item2.Width;
                        s.DrawingAttributes.Height = snap.Item2.Height;
                        history[s] = new Tuple<StylusPointCollection, StylusPointCollection>(oldPts, s.StylusPoints.Clone());
                    }
                    catch { }
                }
            }

            if (history.Count > 0)
            {
                timeMachine.CommitStrokeManipulationHistory(history);
                foreach (var item in history)
                {
                    StrokeInitialHistory[item.Key] = item.Value.Item2;
                }
            }

            //图片还原：恢复到选中时的位置/尺寸/翻转态（与墨迹同语义；
            //图片操作不进撤销栈，与拖动/删除口径一致）
            foreach (var img in GetSelectedPageImages())
            {
                if (!_imageSelectionSnapshot.TryGetValue(img, out var snap)) continue;
                try
                {
                    InkCanvas.SetLeft(img, snap.Left);
                    InkCanvas.SetTop(img, snap.Top);
                    img.Width = snap.Width;
                    img.Height = snap.Height;
                    //翻转态一并还原（快照时没翻 = 恢复为无翻转）
                    var st = img.RenderTransform as ScaleTransform;
                    if (st == null && (snap.ScaleX != 1 || snap.ScaleY != 1))
                    {
                        st = new ScaleTransform(snap.ScaleX, snap.ScaleY);
                        img.RenderTransform = st;
                        img.RenderTransformOrigin = new Point(0.5, 0.5);
                    }
                    else if (st != null)
                    {
                        st.ScaleX = snap.ScaleX;
                        st.ScaleY = snap.ScaleY;
                    }
                }
                catch { }
            }

            CommitDrawingAttributesHistoryNow();
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>立即提交笔画粗细变更历史（避免滞留到下次拖动/缩放时才一并提交，导致撤销步骤混乱）</summary>
        private void CommitDrawingAttributesHistoryNow()
        {
            try
            {
                if (DrawingAttributesHistory.Count > 0)
                {
                    timeMachine.CommitStrokeDrawingAttributesHistory(DrawingAttributesHistory);
                    DrawingAttributesHistory = new Dictionary<Stroke, Tuple<DrawingAttributes, DrawingAttributes>>();
                    foreach (var item in DrawingAttributesHistoryFlag) item.Value.Clear();
                }
            }
            catch { }
        }

        #endregion 选中缩放/还原


        private void BtnSelect_Click(object sender, RoutedEventArgs e)
        {
            forceEraser = true;

            //用户主动切换到选择工具：清掉"一次性选中"标志，
            //之后取消选中时不再自动恢复笔模式（详见 MW_GraphStrokes.cs）
            _isOneShotGraphSelection = false;

            #region 选中快照（用于"还原"按钮：恢复到本次选中时的大小/位置/粗细）

            try
            {
                SelectionSnapshot = new Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>>();
                foreach (Stroke s in inkCanvas.GetSelectedStrokes())
                {
                    SelectionSnapshot[s] = new Tuple<StylusPointCollection, DrawingAttributes>(
                        s.StylusPoints.Clone(), s.DrawingAttributes.Clone());
                }
            }
            catch { }

            #endregion
            drawingShapeMode = 0;
            UpdateShapeIconHighlight(); //切到选择工具时熄灭图形图标高亮
            inkCanvas.IsManipulationEnabled = false;

            // 只负责切入选择模式（幂等：已在选择模式则什么都不做）。
            // 旧版"第二次点击全选、第三次点击回笔"的双击行为已废弃——
            // 全选移入选择方式面板（MW_SelectionMode.SelectAllStrokes），退出走笔图标/直接书写
            if (inkCanvas.EditingMode != InkCanvasEditingMode.Select)
            {
                inkCanvas.EditingMode = InkCanvasEditingMode.Select;
            }
        }

        double BorderStrokeSelectionControlWidth = 338.0; // 10 键×30 + 4 分隔线×7 + 面板边距 8 + 边框 2（定位兜底值，实际优先用 ActualWidth）
        double BorderStrokeSelectionControlHeight = 46.0; // 紧凑单行操作条高度（原双行 80）
        bool isProgramChangeStrokeSelection = false;

        // ===== 操作条收起/展开状态 =====

        /// <summary>本次选中是否手动收起（新选中默认展开，由用户决定收不收）</summary>
        private bool _selectionBarCollapsed = false;

        /// <summary>拖动墨迹期间的临时收起（松手自动恢复展开态，不改变用户的手动选择）</summary>
        private bool _selectionBarDragHidden = false;

        /// <summary>触摸手势累计位移（逐帧 delta 常小于 1px，累计超阈值才算拖动，轻点不收操作条）</summary>
        private double _touchDragTotalMove = 0;

        /// <summary>
        /// 操作条显隐总开关：展开态显示整条，收起态显示＋小圆钮（宿主 Grid 已绑定选区可见性，
        /// 这里只管两者互斥切换）。选区消失时宿主整体隐藏，两个标志在 SelectionChanged 里复位。
        /// </summary>
        private void UpdateSelectionBarVisibility()
        {
            bool showBar = !_selectionBarCollapsed && !_selectionBarDragHidden;
            BorderStrokeSelectionControl.Visibility = showBar ? Visibility.Visible : Visibility.Collapsed;
            BorderStrokeSelectionCollapseBubble.Visibility = showBar ? Visibility.Collapsed : Visibility.Visible;
        }

        /// <summary>拖动开始（鼠标平移/手柄缩放/触摸手势）：操作条临时收成＋小圆钮</summary>
        private void TryCollapseBarForDrag()
        {
            if (_selectionBarDragHidden) return;
            _selectionBarDragHidden = true;
            UpdateSelectionBarVisibility();
            //宽度已变（条→圆钮），强制布局后按新宽度重新居中，否则下帧定位按旧宽度算会偏
            try { GridSelectionBarHost.UpdateLayout(); updateBorderStrokeSelectionControlLocation(); } catch { }
        }

        /// <summary>拖动结束：恢复展开态（用户手动收起过则保持收起）</summary>
        private void TryRestoreBarAfterDrag()
        {
            if (!_selectionBarDragHidden) return;
            _selectionBarDragHidden = false;
            UpdateSelectionBarVisibility();
            //宽度已变（圆钮→条），强制布局让 ActualWidth 就绪，再按选中框重新定位
            try { GridSelectionBarHost.UpdateLayout(); updateBorderStrokeSelectionControlLocation(); } catch { }
        }

        /// <summary>收起钮：操作条 → ＋小圆钮</summary>
        private void BorderStrokeSelectionCollapse_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            _selectionBarCollapsed = true;
            UpdateSelectionBarVisibility();
            //宽度已变（条→圆钮），强制布局让 ActualWidth 就绪，再按新宽度重新居中——
            //否则定位仍按旧条宽计算，28px 圆钮会偏在旧条左端，视觉上严重不居中
            try { GridSelectionBarHost.UpdateLayout(); } catch { }
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>＋小圆钮：重新展开操作条</summary>
        private void BorderStrokeSelectionCollapseBubble_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            _selectionBarCollapsed = false;
            _selectionBarDragHidden = false;
            UpdateSelectionBarVisibility();
            //同上：宽度已变（圆钮→条），先强制布局再定位
            try { GridSelectionBarHost.UpdateLayout(); } catch { }
            updateBorderStrokeSelectionControlLocation();
        }

        /// <summary>
        /// 保存到本地：选中内容导出为透明背景 PNG。纯墨迹与含图片共用 RenderStrokesAndImages
        /// （墨迹+图片整体所见即所得渲染）；四周留 12px 边距，原尺寸 1:1。
        /// </summary>
        private void BorderStrokeSelectionSaveToFile_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;

            try
            {
                var strokes = inkCanvas.GetSelectedStrokes();
                var selectedImages = GetSelectedPageImages();
                if (strokes.Count == 0 && selectedImages.Count == 0) return;

                //统一渲染（纯墨迹/纯图片/混合都支持，含翻转态的渲染层变换）
                var bmp = RenderStrokesAndImagesWithFlip(strokes, selectedImages);
                if (bmp == null) return;

                //四周加 12px 边距再导出（DrawingVisual 平移即可，重新渲染一次）
                const double margin = 12;
                int w = (int)Math.Ceiling(bmp.Width + margin * 2);
                int h = (int)Math.Ceiling(bmp.Height + margin * 2);
                DrawingVisual dv = new DrawingVisual();
                using (DrawingContext dc = dv.RenderOpen())
                {
                    dc.DrawImage(bmp, new Rect(margin, margin, bmp.Width, bmp.Height));
                }
                var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
                rtb.Render(dv);

                var dlg = new Microsoft.Win32.SaveFileDialog
                {
                    Title = selectedImages.Count > 0 ? "保存选中内容为图片" : "保存墨迹为图片",
                    Filter = "PNG 图片|*.png",
                    FileName = (selectedImages.Count > 0 ? "选中内容_" : "墨迹_")
                        + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".png"
                };
                if (dlg.ShowDialog() == true)
                {
                    var encoder = new PngBitmapEncoder();
                    encoder.Frames.Add(BitmapFrame.Create(rtb));
                    using (var fs = File.Create(dlg.FileName))
                        encoder.Save(fs);
                    ShowToastNotification("已保存到本地：" + System.IO.Path.GetFileName(dlg.FileName));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 渲染选中内容（含图片翻转态）：在 RenderStrokesAndImages 基础上，
        /// 把图片的 RenderTransform（翻转）也应用进渲染——否则翻转过的图片存出来是原图方向。
        /// </summary>
        private RenderTargetBitmap RenderStrokesAndImagesWithFlip(StrokeCollection strokes, List<System.Windows.Controls.Image> images)
        {
            try
            {
                //先把翻转态应用到 DrawImage 的目标矩形镜像（用 PushTransform 包住图片绘制）
                Rect bounds = Rect.Empty;
                if (strokes.Count > 0) bounds = strokes.GetBounds();
                foreach (var img in images)
                {
                    var r = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                    bounds = bounds.IsEmpty ? r : Rect.Union(bounds, r);
                }
                if (bounds.IsEmpty || bounds.Width < 1 || bounds.Height < 1) return null;

                var dv = new DrawingVisual();
                using (DrawingContext dc = dv.RenderOpen())
                {
                    dc.PushTransform(new TranslateTransform(-bounds.X, -bounds.Y));
                    foreach (Stroke s in strokes) s.Draw(dc);
                    foreach (var img in images)
                    {
                        var st = img.RenderTransform as ScaleTransform;
                        var rect = new Rect(InkCanvas.GetLeft(img), InkCanvas.GetTop(img), img.Width, img.Height);
                        if (st != null && (st.ScaleX < 0 || st.ScaleY < 0))
                        {
                            //翻转：平移到图片中心、按翻转符号缩放、再平移回来（等价于中心点镜像）
                            var cx = rect.X + rect.Width / 2;
                            var cy = rect.Y + rect.Height / 2;
                            var m = new Matrix();
                            m.Translate(-cx, -cy);
                            m.Scale(st.ScaleX, st.ScaleY);
                            m.Translate(cx, cy);
                            dc.PushTransform(new MatrixTransform(m));
                            dc.DrawImage(img.Source, rect);
                            dc.Pop();
                        }
                        else
                        {
                            dc.DrawImage(img.Source, rect);
                        }
                    }
                    dc.Pop();
                }
                var rtb = new RenderTargetBitmap(
                    (int)Math.Ceiling(bounds.Width), (int)Math.Ceiling(bounds.Height),
                    96, 96, System.Windows.Media.PixelFormats.Pbgra32);
                rtb.Render(dv);
                rtb.Freeze();
                return rtb;
            }
            catch { return null; }
        }

        /// <summary>
        /// 按当前选中内容动态显隐操作条按钮（统一操作条原则：不适用的键直接隐藏不占位）。
        /// 纯图片选中：隐藏"变细/变粗/存图库"（墨迹专属）；含墨迹（纯墨迹或混合）：显示。
        /// 显隐后操作条宽度变化，需重新布局并按选中框重新居中。
        /// </summary>
        private void UpdateSelectionToolbarButtons()
        {
            try
            {
                bool hasInk = inkCanvas.GetSelectedStrokes().Count > 0;
                var vis = hasInk ? Visibility.Visible : Visibility.Collapsed;
                GridSelectionPenThin.Visibility = vis;
                GridSelectionPenThick.Visibility = vis;
                GridSelectionSaveShape.Visibility = vis;
                //适配白板宽度键反向显隐：选中含图片才显示，纯墨迹隐藏（图片专属操作）
                GridSelectionFitWidth.Visibility = hasInk ? Visibility.Collapsed : Visibility.Visible;
                //宽度变化后重新定位（UpdateLayout 让 ActualWidth 先就绪）
                GridSelectionBarHost.UpdateLayout();
                updateBorderStrokeSelectionControlLocation();
            }
            catch { }
        }

        private void inkCanvas_SelectionChanged(object sender, EventArgs e)
        {
            if (isProgramChangeStrokeSelection) return;

            // 非程序化的选区变化（换选别的墨迹/清空/框选/全选）→ 复制模式一律结束：
            // 用户做了新的选择动作，之前"按住图形拖出副本"的意图不再有效。
            // 程序化选区变化（CreateCopyDragClone 等）已被上方挡板拦下，不受影响
            ExitCopyDragMode();

            //选中判定含图片：纯图片选中同样出覆盖层+操作条（模仿墨迹逻辑）
            bool hasSelection = inkCanvas.GetSelectedStrokes().Count > 0 || GetSelectedPageImages().Count > 0;
            if (!hasSelection)
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                //复位收起状态：新选中一律从展开态开始（默认展开，用户可随时再收起）
                _selectionBarCollapsed = false;
                _selectionBarDragHidden = false;
                //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                TryEndOneShotSelection();
            }
            else
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                //新选中默认展开操作条
                _selectionBarCollapsed = false;
                _selectionBarDragHidden = false;
                UpdateSelectionBarVisibility();
                //按选中内容（纯图/纯墨迹/混合）显隐专属按钮
                UpdateSelectionToolbarButtons();

                //捕获选中快照（还原按钮用：恢复到选中时的大小/位置/粗细）
                try
                {
                    SelectionSnapshot = new Dictionary<Stroke, Tuple<StylusPointCollection, DrawingAttributes>>();
                    foreach (Stroke s in inkCanvas.GetSelectedStrokes())
                    {
                        SelectionSnapshot[s] = new Tuple<StylusPointCollection, DrawingAttributes>(
                            s.StylusPoints.Clone(), s.DrawingAttributes.Clone());
                    }

                    //图片快照（还原按钮用：位置/尺寸/翻转态）
                    _imageSelectionSnapshot = new Dictionary<System.Windows.Controls.Image, ImageSnapshot>();
                    foreach (var img in GetSelectedPageImages())
                    {
                        var st = img.RenderTransform as ScaleTransform;
                        _imageSelectionSnapshot[img] = new ImageSnapshot
                        {
                            Left = InkCanvas.GetLeft(img),
                            Top = InkCanvas.GetTop(img),
                            Width = img.Width,
                            Height = img.Height,
                            ScaleX = st?.ScaleX ?? 1,
                            ScaleY = st?.ScaleY ?? 1
                        };
                    }
                }
                catch { }

                updateBorderStrokeSelectionControlLocation();
            }
        }

        /// <summary>
        /// 选区可命中区域几何更新：跟随选中框外扩——
        /// 四边按手柄命中半径（14）外扩，顶部为旋转钮多留（钮心在框顶上方 22 + 命中半径 14 = 36）。
        /// 之外的区域不可命中 → 输入直接落到 inkCanvas（选中状态下也能直接书写）。
        /// 由 updateBorderStrokeSelectionControlLocation 搭车调用（选区变化/拖动/旋转/缩放后都会走它）。
        /// </summary>
        private void UpdateSelectionHitArea()
        {
            try
            {
                var b = GetGestureSelectionBounds();
                if (b.IsEmpty || b.Width <= 0 || b.Height <= 0)
                {
                    BorderSelectionHitArea.Width = 0;
                    BorderSelectionHitArea.Height = 0;
                    return;
                }
                //inkCanvas 坐标系 → 覆盖层 Grid 坐标系（覆盖层无背景可命中，
                //Border 用 Margin+尺寸绝对定位，所以必须自己换算坐标）
                var t = inkCanvas.TransformToVisual(GridInkCanvasSelectionCover);
                var tl = t.Transform(b.TopLeft);

                const double side = SelectionHandleHitRadius;      //四边：手柄命中半径
                const double top = RotateHandleOffset + SelectionHandleHitRadius; //顶边：再加旋转钮伸出距离

                BorderSelectionHitArea.Margin = new Thickness(tl.X - side, tl.Y - top, 0, 0);
                BorderSelectionHitArea.Width = b.Width + side * 2;
                BorderSelectionHitArea.Height = b.Height + top + side;
            }
            catch { }
        }

        /// <summary>
        /// 框外落笔时取消选中（配合覆盖层只覆盖选区附近的新结构）：
        /// 选中不再拦截画布输入，笔/鼠标落到 inkCanvas 的瞬间调用本方法，
        /// 清掉选中让这一笔直接写出（含一次性选中残留的 Select 模式恢复为笔模式）。
        /// 用户主动用"选择工具"选中的场景（_isOneShotGraphSelection=false）保持 Select 模式——
        /// 用户接下来通常是继续框选新内容。
        /// </summary>
        private void DeselectStrokesForCanvasInput()
        {
            try
            {
                //判定含图片：纯图片选中时框外落笔同样要取消（否则图片选区残留挡书写）
                if (inkCanvas.GetSelectedStrokes().Count == 0 && GetSelectedPageImages().Count == 0) return;

                //先读一次性选中标志：true = 选中来自图形插入（模式被 Select() 偷偷切成了 Select），
                //取消后要恢复笔模式，当前这一笔才画得出来
                bool restoreInk = _isOneShotGraphSelection;
                bool eraserShapeBefore = forcePointEraser;

                //屏蔽 SelectionChanged 的快照副作用（取消不需要快照）；元素集合一并清空（含图片选中）
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection(), new System.Collections.Generic.List<UIElement>());
                isProgramChangeStrokeSelection = false;

                if (restoreInk)
                {
                    inkCanvas.EditingMode = InkCanvasEditingMode.Ink;
                    forcePointEraser = eraserShapeBefore;
                }

                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                TryEndOneShotSelection(); //一次性选中收尾（内部幂等，延迟恢复笔模式与本处不冲突）
            }
            catch { }
        }

        private void updateBorderStrokeSelectionControlLocation()
        {
            //定位宿主 Grid（内含展开条/收起圆钮两个互斥子元素，宽度即当前可见者）；
            //宽度优先用 ActualWidth（布局完成后 > 10），常量仅作初值兜底
            double controlWidth = GridSelectionBarHost.ActualWidth > 10 ? GridSelectionBarHost.ActualWidth : BorderStrokeSelectionControlWidth;
            double borderLeft = (GetGestureSelectionBounds().Left + GetGestureSelectionBounds().Right - controlWidth) / 2;
            double borderTop = GetGestureSelectionBounds().Bottom + 15;
            if (borderLeft < 0) borderLeft = 0;
            if (borderTop < 0) borderTop = 0;
            if (Width - borderLeft < controlWidth || double.IsNaN(borderLeft)) borderLeft = Width - controlWidth;
            if (Height - borderTop < BorderStrokeSelectionControlHeight || double.IsNaN(borderTop)) borderTop = Height - BorderStrokeSelectionControlHeight;
            GridSelectionBarHost.Margin = new Thickness(borderLeft, borderTop, 0, 0);

            //旋转钮跟随选中框顶部中央（与 HitTestSelectionHandle 的命中点同一位置：
            //框顶边中点向上伸出 RotateHandleOffset 像素处，再减去钮自身半径居中）
            try
            {
                var b = GetGestureSelectionBounds();
                GridRotateHandle.Margin = new Thickness(
                    b.Left + b.Width / 2 - GridRotateHandle.Width / 2,
                    Math.Max(0, b.Top - RotateHandleOffset - GridRotateHandle.Height / 2),
                    0, 0);
            }
            catch { }

            //选区可命中区域同步跟随（覆盖层新结构：只覆盖选中框附近，框外输入直达画布）
            UpdateSelectionHitArea();
        }

        private void GridInkCanvasSelectionCover_ManipulationStarting(object sender, ManipulationStartingEventArgs e)
        {
            e.Mode = ManipulationModes.All;
            _touchDragTotalMove = 0; // 新手势开始：累计位移清零
        }

        private void GridInkCanvasSelectionCover_ManipulationCompleted(object sender, ManipulationCompletedEventArgs e)
        {
            TryRestoreBarAfterDrag(); // 触摸拖动结束：操作条恢复展开态
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

        private void GridInkCanvasSelectionCover_ManipulationDelta(object sender, ManipulationDeltaEventArgs e)
        {
            try
            {
                if (dec.Count >= 1)
                {
                    ManipulationDelta md = e.DeltaManipulation;
                    Vector trans = md.Translation;  // 获得位移矢量
                    double rotate = md.Rotation;  // 获得旋转角度
                    Vector scale = md.Scale;  // 获得缩放倍数

                    //累计位移超过阈值才算拖动（轻点的微小抖动不收操作条，避免闪烁）
                    _touchDragTotalMove += Math.Abs(trans.X) + Math.Abs(trans.Y);
                    if (_touchDragTotalMove > 4)
                    {
                        TryCollapseBarForDrag(); // 触摸拖动中：操作条临时收成＋小圆钮
                    }

                    Matrix m = new Matrix();

                    // Find center of element and then transform to get current location of center
                    FrameworkElement fe = e.Source as FrameworkElement;
                    Point center = new Point(fe.ActualWidth / 2, fe.ActualHeight / 2);
                    center = new Point(GetGestureSelectionBounds().Left + GetGestureSelectionBounds().Width / 2,
                        GetGestureSelectionBounds().Top + GetGestureSelectionBounds().Height / 2);
                    center = m.Transform(center);  // 转换为矩阵缩放和旋转的中心点

                    // Update matrix to reflect translation/rotation
                    m.Translate(trans.X, trans.Y);  // 移动
                    m.ScaleAt(scale.X, scale.Y, center.X, center.Y);  // 缩放

                    StrokeCollection strokes = inkCanvas.GetSelectedStrokes();
                    if (StrokesSelectionClone.Count != 0)
                    {
                        strokes = StrokesSelectionClone;
                    }
                    else if (Settings.Gesture.IsEnableTwoFingerRotationOnSelection)
                    {
                        m.RotateAt(rotate, center.X, center.Y);  // 旋转
                    }
                    foreach (Stroke stroke in strokes)
                    {
                        stroke.Transform(m, false);

                        try
                        {
                            stroke.DrawingAttributes.Width *= md.Scale.X;
                            stroke.DrawingAttributes.Height *= md.Scale.Y;
                        }
                        catch { }
                    }

                    //图片：复制模式拖动时移动副本（原件不动），普通拖动移动选中原件。
                    //平移+绕中心缩放（与墨迹矩阵同序，保证混合选中时两者几何同步；旋转不支持图片）
                    var touchImages = _imageDragClones.Count > 0 ? _imageDragClones : GetSelectedPageImages();
                    foreach (var img in touchImages)
                    {
                        //先平移（矩阵第一段）
                        InkCanvas.SetLeft(img, InkCanvas.GetLeft(img) + trans.X);
                        InkCanvas.SetTop(img, InkCanvas.GetTop(img) + trans.Y);
                        //再绕选区中心缩放（矩阵第二段，锚点用变换前 center）
                        ScaleImage(img, scale.X, scale.Y, center.X, center.Y);
                    }
                    updateBorderStrokeSelectionControlLocation();
                }
            }
            catch { }
        }

        private void GridInkCanvasSelectionCover_TouchDown(object sender, TouchEventArgs e)
        {
        }

        private void GridInkCanvasSelectionCover_TouchUp(object sender, TouchEventArgs e)
        {
        }

        Point lastTouchPointOnGridInkCanvasCover = new Point(0, 0);
        private void GridInkCanvasSelectionCover_PreviewTouchDown(object sender, TouchEventArgs e)
        {
            dec.Add(e.TouchDevice.Id);
            //设备1个的时候，记录中心点
            if (dec.Count == 1)
            {
                TouchPoint touchPoint = e.GetTouchPoint(null);
                centerPoint = touchPoint.Position;
                lastTouchPointOnGridInkCanvasCover = touchPoint.Position;

                // 复制拖拽模式（触摸版）：手指按在选中框内 = 拖出一份副本，
                // 之后正常手势路径移动的是副本（原件不动）——与鼠标路径同语义
                if (isCopyDragMode)
                {
                    var b = GetGestureSelectionBounds();
                    b.Inflate(10, 10);
                    if (b.Contains(touchPoint.Position)) CreateCopyDragClone();
                }
            }
        }

        private void GridInkCanvasSelectionCover_PreviewTouchUp(object sender, TouchEventArgs e)
        {
            dec.Remove(e.TouchDevice.Id);
            if (dec.Count >= 1) return;
            isProgramChangeStrokeSelection = false;
            if (lastTouchPointOnGridInkCanvasCover == e.GetTouchPoint(null).Position)
            {
                if (lastTouchPointOnGridInkCanvasCover.X < GetGestureSelectionBounds().Left ||
                    lastTouchPointOnGridInkCanvasCover.Y < GetGestureSelectionBounds().Top ||
                    lastTouchPointOnGridInkCanvasCover.X > GetGestureSelectionBounds().Right ||
                    lastTouchPointOnGridInkCanvasCover.Y > GetGestureSelectionBounds().Bottom)
                {
                    //双参清空：图片选中一并取消（与鼠标路径 MouseUp 分支同一口径）
                    inkCanvas.Select(new StrokeCollection(), new System.Collections.Generic.List<UIElement>());
                    StrokesSelectionClone = new StrokeCollection();
                    //与鼠标路径（MouseUp 分支）保持一致：取消选中的同时收起选区遮罩，
                    //否则触摸点掉选中后拖动控制条仍残留悬空。
                    GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                    //一次性选中收尾：图形插入产生的选中被取消 → 恢复笔模式（见 MW_GraphStrokes.cs）
                    TryEndOneShotSelection();
                }
            }
            else if (inkCanvas.GetSelectedStrokes().Count == 0)
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
                StrokesSelectionClone = new StrokeCollection();
                _imageDragClones = new List<System.Windows.Controls.Image>(); //触摸克隆会话收尾
            }
            else
            {
                GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                StrokesSelectionClone = new StrokeCollection();
                _imageDragClones = new List<System.Windows.Controls.Image>(); //触摸克隆会话收尾
            }
        }

        #endregion Selection Gestures
    }
}
