using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using Microsoft.Win32;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：截图与图片功能（自绘遮罩框选版）。
    ///
    /// 【为什么弃用系统截图（ms-screenclip:）】
    /// 系统截图是"黑盒"：没有完成/取消回调，只能靠监听剪贴板+超时去猜结果——
    /// 用户半途取消、期间复制了别的图，都会造成误判（莫名插入旧图/界面卡在隐藏状态）。
    ///
    /// 【本方案（方案 B）】
    /// 自己弹全屏半透明遮罩窗口（ScreenshotMaskWindow），用户在上面拖框选区：
    /// - 完成 = 松手（选区够大）→ 先关遮罩再 CopyFromScreen 抓屏 → 图落画布左上角 + 进剪贴板
    /// - 取消 = Esc / 右键 / 单击没拖开 / 窗口失焦 → 关窗即恢复原状，瞬时、无副作用、无等待
    /// 整个状态机闭环在遮罩窗口里，不存在"等外部程序结果"的悬空状态。
    ///
    /// 【两个截图入口的差异 = 一个布尔开关】
    /// - 快速截图：进遮罩前不藏 UI（用户在遮罩下看到屏幕原样，可框选桌面/PPT 等任何内容）；
    ///   选区确定的一瞬间（遮罩还盖着屏幕时）才同步藏 UI，保证最终截图里没有悬浮条。
    /// - 隐藏界面截图：进遮罩前先藏 UI 并等渲染完成，遮罩下只剩板书内容，
    ///   用于截取被悬浮条/面板挡住的画面。
    ///
    /// 模块自包含：所有入口 try/catch，异常只通知不抛出；UI 隐藏必配恢复（try/finally）。
    /// </summary>
    public partial class MainWindow
    {
        #region 截图按钮弹出菜单（悬浮条相机图标）

        /// <summary>相机图标：开/关截图功能弹出菜单（3 项，风格同清屏确认气泡）</summary>
        private void SymbolIconScreenshot_MenuToggle(object sender, MouseButtonEventArgs e)
        {
            ImageLayer_MenuSetVisible(BorderImageMenu.Visibility != Visibility.Visible);
        }

        /// <summary>菜单开关统一入口（true=显示，false=隐藏）</summary>
        private void ImageLayer_MenuSetVisible(bool visible)
        {
            BorderImageMenu.Visibility = visible ? Visibility.Visible : Visibility.Collapsed;
        }

        /// <summary>
        /// 点菜单外任意位置自动关菜单（触摸屏友好，不用专门找关闭按钮）。
        /// 挂在窗口 Preview 阶段只做关菜单这一件事，不设 e.Handled，不影响书写/擦除等正常操作。
        /// </summary>
        internal void ImageLayer_MenuCloseOnOutsideClick(object sender, MouseButtonEventArgs e)
        {
            if (BorderImageMenu.Visibility != Visibility.Visible) return;
            if (!(e.OriginalSource is DependencyObject d)) return;
            // 点在菜单内部（含子按钮）不处理，交由各菜单项自己的 Click 逻辑
            if (IsDescendantOf(d, BorderImageMenu)) return;
            // 点在相机图标本身也不处理（图标自己的开关逻辑生效）
            if (IsDescendantOf(d, GridImageMenuEntry)) return;
            ImageLayer_MenuSetVisible(false);
        }

        /// <summary>判断元素是否是 target 的子孙（沿视觉树向上找）</summary>
        private static bool IsDescendantOf(DependencyObject node, DependencyObject target)
        {
            while (node != null)
            {
                if (ReferenceEquals(node, target)) return true;
                node = System.Windows.Media.VisualTreeHelper.GetParent(node);
            }
            return false;
        }

        /// <summary>菜单项：快速截图（遮罩框选屏幕任意内容，悬浮条不会入镜）</summary>
        private void BtnMenuShotRegion_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            Screenshot_Start(hideUiFirst: false);
        }

        /// <summary>菜单项：隐藏界面截图（先藏起本软件 UI 再框选，截被悬浮条/面板挡住的内容）</summary>
        private void BtnMenuShotHidden_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            Screenshot_Start(hideUiFirst: true);
        }

        /// <summary>菜单项：上传本地图片（文件选择框 → 落画布左上角，多选时梯级错位）</summary>
        private void BtnMenuImportImage_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false); // 先关菜单再弹文件对话框，避免两个浮层叠着
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "选择图片（将插入白板当前页）",
                    Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*",
                    Multiselect = true // 支持一次选多张，逐张插入
                };
                if (dlg.ShowDialog() != true) return;

                int count = 0;
                foreach (var file in dlg.FileNames)
                {
                    try
                    {
                        var source = new BitmapImage();
                        source.BeginInit();
                        source.CacheOption = BitmapCacheOption.OnLoad; // 立即读入内存，文件可立即被覆盖/删除
                        source.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                        source.UriSource = new Uri(file, UriKind.Absolute);
                        source.EndInit();

                        // cascade=true：连续插入的多张图右下错开摆放，避免完全叠死
                        if (ImageLayer_AddImage(source, cascade: true) != null) count++;
                    }
                    catch { /* 单张失败跳过，继续插后面的 */ }
                }
                if (count > 0)
                    ShowNotification("已插入 " + count + " 张图片到白板第 " + CurrentWhiteboardIndex + " 页");
            }
            catch (Exception ex)
            {
                ShowNotification("打开图片失败：" + ex.Message);
            }
        }

        #endregion

        #region 自绘遮罩框选截图（核心流程）

        /// <summary>重入锁：截图流程进行中忽略再次触发（防连点/快捷键连按）</summary>
        bool _isScreenshotBusy = false;

        /// <summary>截图期间被隐藏的 UI 元素表（元素 → 隐藏前的可见性），恢复时逐个还原</summary>
        System.Collections.Generic.Dictionary<UIElement, Visibility> _screenshotHiddenUi;

        /// <summary>截图期间是否隐藏了数学输入面板（独立 COM 窗口，恢复要单独 Show）</summary>
        bool _mathPanelHiddenForScreenshot = false;

        /// <summary>
        /// 截图入口（两个菜单项共用，只差"进遮罩前藏不藏 UI"）。
        /// hideUiFirst=false：快速截图——直接上遮罩，屏幕原样供框选；
        /// hideUiFirst=true：隐藏界面截图——先藏 UI、等渲染完成再上遮罩，遮罩下只剩板书。
        /// </summary>
        private void Screenshot_Start(bool hideUiFirst)
        {
            if (_isScreenshotBusy) return; // 上一次截图还没收尾，直接忽略
            _isScreenshotBusy = true;
            try
            {
                if (hideUiFirst)
                {
                    Screenshot_HideChrome(hideBoard: true);
                    // 等 UI 消失真正画完再上遮罩（两级 Background 回调 = 至少跑完一轮渲染），
                    // 否则遮罩下还能看到悬浮条的最后一帧
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                        Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                            Screenshot_ShowMask(hideUiFirst)))));
                }
                else
                {
                    Screenshot_ShowMask(hideUiFirst);
                }
            }
            catch (Exception ex)
            {
                // 任何异常都必须恢复 UI 并解锁，不能出现"界面消失回不来"的死局
                Screenshot_RestoreChrome();
                _isScreenshotBusy = false;
                ShowNotification("截图失败：" + ex.Message);
            }
        }

        /// <summary>
        /// 弹出遮罩窗口收集选区（模态）。返回后按"有选区/取消"两条路收尾。
        /// </summary>
        private void Screenshot_ShowMask(bool hideUiFirst)
        {
            var mask = new ScreenshotMaskWindow();
            if (!hideUiFirst)
            {
                // 快速截图：选区确定的一瞬间（遮罩还盖着屏幕时）才藏软件 UI——
                // 这样遮罩关闭后屏幕直接就是干净画面，不会有悬浮条闪现帧。
                // 注意 hideBoard=false：板书/白板保留入镜（快速截图截"眼前的一切"）
                mask.SelectionMade += () => Screenshot_HideChrome(hideBoard: false);
            }

            bool? ok = false;
            Rect rect = Rect.Empty;
            double scaleX = 1, scaleY = 1;
            try
            {
                ok = mask.ShowDialog(); // 模态：窗口关闭（完成或取消）才返回
                if (mask.SelectedRect.HasValue)
                {
                    rect = mask.SelectedRect.Value;
                    scaleX = mask.DpiScaleX;
                    scaleY = mask.DpiScaleY;
                }
            }
            catch { ok = false; }

            if (ok == true && !rect.IsEmpty)
            {
                // ---- 完成路径：等遮罩退场渲染完 → 抓屏 → 落图 → 恢复 ----
                // 逻辑坐标 × DPI 缩放 = 物理像素（CopyFromScreen 用物理像素；
                // 150% 缩放的教学一体机不做这步截图区域会偏移错位）
                int px = (int)Math.Round(rect.X * scaleX);
                int py = (int)Math.Round(rect.Y * scaleY);
                int pw = (int)Math.Round(rect.Width * scaleX);
                int ph = (int)Math.Round(rect.Height * scaleY);

                // 两级 Background 回调：确保"遮罩已关 + UI 已藏"渲染完毕，屏幕稳定后再抓
                Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                    Dispatcher.BeginInvoke(DispatcherPriority.Background, (Action)(() =>
                    {
                        try
                        {
                            var bmp = Screenshot_CaptureRegion(px, py, pw, ph);
                            if (bmp == null)
                            {
                                ShowNotification("截图失败：无法捕获屏幕区域");
                                return;
                            }
                            // 截图同步进剪贴板（可直接 Ctrl+V 到 PPT/微信）
                            try { Clipboard.SetImage(bmp); } catch { }
                            if (ImageLayer_AddImage(bmp) != null)
                                ShowNotification("截图已插入，并复制到剪贴板");
                        }
                        catch (Exception ex)
                        {
                            ShowNotification("截图失败：" + ex.Message);
                        }
                        finally
                        {
                            Screenshot_RestoreChrome();
                            _isScreenshotBusy = false;
                        }
                    }))));
            }
            else
            {
                // ---- 取消路径：瞬时恢复原状，无任何副作用 ----
                Screenshot_RestoreChrome();
                _isScreenshotBusy = false;
            }
        }

        /// <summary>
        /// 抓取屏幕指定区域（物理像素坐标）为 WPF 位图。
        /// 注意：必须在遮罩窗口关闭之后调用，否则会把遮罩自己截进去。
        /// </summary>
        private static BitmapSource Screenshot_CaptureRegion(int x, int y, int width, int height)
        {
            if (width < 1 || height < 1) return null;
            try
            {
                using (var bmp = new System.Drawing.Bitmap(width, height,
                    System.Drawing.Imaging.PixelFormat.Format32bppArgb))
                {
                    using (var g = System.Drawing.Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height),
                            System.Drawing.CopyPixelOperation.SourceCopy);
                    }
                    // GDI 位图 → WPF 位图源（复制像素，之后 GDI 资源可安全释放）
                    IntPtr hBitmap = bmp.GetHbitmap();
                    try
                    {
                        return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                            hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                    }
                    finally
                    {
                        DeleteObject(hBitmap); // GetHbitmap 创建的非托管资源必须手动释放（防内存泄漏）
                    }
                }
            }
            catch { return null; }
        }

        /// <summary>
        /// 隐藏本软件所有会入镜的 UI。
        /// hideBoard=true（隐藏界面截图）：掀板去墨——把白板板面、墨迹/图片（inkCanvas）、
        /// 悬浮条、各弹层面板全部藏掉，回到"没开白板前"的裸桌面供框选。
        /// hideBoard=false（快速截图收尾时）：只藏悬浮条/面板——板书内容保留入镜
        /// （快速截图截的就是"眼前看到的一切"，选区确定瞬间只把软件 UI 撤走）。
        /// 细节：
        /// - GridBackgroundCoverHolder = 板面容器（白板模式 Visible / 批注模式本来 Collapsed），
        ///   藏它 = 掀板；批注模式（第 0 页透明白板）本来就无板面，记原值机制自动跳过。
        /// - inkCanvas = 墨迹 + 图片层；选区控制层父 Grid 绑定了 inkCanvas.Visibility，
        ///   藏它墨迹/图片/选区控件一层全隐（不动 Strokes 集合，零副作用，恢复一个属性搞定）。
        /// - 用 Visibility.Hidden（保持布局占位）而不是 Collapsed，恢复零风险；
        ///   记录隐藏前的可见性，恢复时逐个还原（本来就不显示的不动它）。
        /// </summary>
        private void Screenshot_HideChrome(bool hideBoard)
        {
            try
            {
                if (_screenshotHiddenUi != null) return; // 已在隐藏状态，不重复记录

                // 悬浮条（含笑脸收起态）+ 工具/图形面板 + 各弹出浮层（两种截图都要藏）
                var targets = new System.Collections.Generic.List<UIElement>
                {
                    ViewboxFloatingBar, BorderTools, BorderDrawShape,
                    BorderImageMenu, BorderClearInDelete, BorderPenWidth
                };
                if (hideBoard)
                {
                    targets.Add(GridBackgroundCoverHolder); // 板面（掀板；批注模式已 Collapsed 自动跳过）
                    targets.Add(inkCanvas);                 // 墨迹 + 图片（选区控制层绑定联动隐藏）
                }

                _screenshotHiddenUi = new System.Collections.Generic.Dictionary<UIElement, Visibility>();
                foreach (var t in targets)
                {
                    if (t == null) continue;
                    if (t.Visibility == Visibility.Hidden) continue; // 本来就藏的不记录，避免恢复时误显示
                    _screenshotHiddenUi[t] = t.Visibility;
                    t.Visibility = Visibility.Hidden;
                }

                // 数学输入面板是独立 COM 窗口（micaut），不属 WPF 视觉树，单独处理
                try
                {
                    bool visible = false;
                    _mathInputPanel?.IsVisible(out visible);
                    if (visible)
                    {
                        _mathPanelHiddenForScreenshot = true;
                        _mathInputPanel?.Hide();
                    }
                }
                catch { }
            }
            catch { }
        }

        /// <summary>恢复被隐藏的 UI（与 HideChrome 严格配对，可安全重复调用）</summary>
        private void Screenshot_RestoreChrome()
        {
            try
            {
                if (_screenshotHiddenUi != null)
                {
                    foreach (var kv in _screenshotHiddenUi)
                        kv.Key.Visibility = kv.Value; // 还原到隐藏前的原值
                    _screenshotHiddenUi = null;
                }
                if (_mathPanelHiddenForScreenshot)
                {
                    _mathPanelHiddenForScreenshot = false;
                    try { _mathInputPanel?.Show(); } catch { }
                }
            }
            catch { }
        }

        // Win32：释放 GDI 位图句柄（Screenshot_CaptureRegion 用）
        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        #endregion
    }
}
