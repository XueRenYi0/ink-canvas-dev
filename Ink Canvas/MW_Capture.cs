using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using Microsoft.Win32;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：截图与图片功能（简单版，全部对接系统能力，不自造轮子）。
    /// - 系统截图：调起 Windows 系统截图（ms-screenclip: 协议 = Win+Shift+S 的截图工具），
    ///   用户框选完成后系统把图写入剪贴板；本程序监听剪贴板变化（AddClipboardFormatListener）
    ///   收图 → 拉回窗口焦点 → 存入白板抽屉（CurrentWhiteboardIndex 页）。
    ///   多显示器/高分屏由系统兜底，零适配成本。
    /// - 隐藏界面截图：先隐藏本软件（含悬浮条/墨迹）→ 稍候调系统截图 → 收图或超时后自动恢复。
    /// - 打开本地图片：系统文件选择框 → 解码 → 存入白板抽屉。
    /// 防误插：60 秒超时 + 落笔解除等待（用户取消系统截图后，之后复制的图不会被误插）。
    /// 模块自包含：所有入口 try/catch，异常只通知不抛出。
    /// </summary>
    public partial class MainWindow
    {
        #region 相机图标弹出菜单（悬浮条）

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

        /// <summary>菜单项：系统截图（调起系统截图，框选完自动存入白板）</summary>
        private void BtnMenuShotRegion_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            LaunchSystemRegionScreenshot();
        }

        /// <summary>菜单项：隐藏界面截图（先藏起本软件再系统截图，截被挡住的内容）</summary>
        private void BtnMenuShotHidden_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false);
            try
            {
                _hiddenForScreenshot = true;
                Hide();
                // 延时等窗口真正退场再调系统截图，否则截到的还是本软件界面
                var t = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
                t.Tick += (s, _) =>
                {
                    t.Stop();
                    LaunchSystemRegionScreenshot();
                };
                t.Start();
            }
            catch (Exception ex)
            {
                _hiddenForScreenshot = false;
                Show();
                ShowNotification("隐藏界面截图失败：" + ex.Message);
            }
        }

        /// <summary>菜单项：打开本地图片（文件选择框 → 存入白板抽屉）</summary>
        private void BtnMenuImportImage_Click(object sender, RoutedEventArgs e)
        {
            ImageLayer_MenuSetVisible(false); // 先关菜单再弹文件对话框，避免两个浮层叠着
            try
            {
                var dlg = new OpenFileDialog
                {
                    Title = "选择图片（将插入白板当前页）",
                    Filter = "图片文件|*.png;*.jpg;*.jpeg;*.bmp;*.gif|所有文件|*.*"
                };
                if (dlg.ShowDialog() != true) return;

                var source = new BitmapImage();
                source.BeginInit();
                source.CacheOption = BitmapCacheOption.OnLoad; // 立即读入内存，文件可立即被覆盖/删除
                source.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                source.UriSource = new Uri(dlg.FileName, UriKind.Absolute);
                source.EndInit();

                ImageLayer_AddImage(source);
                if (currentMode != 0)
                    ShowNotification("图片已插入白板第 " + CurrentWhiteboardIndex + " 页");
                else
                    ShowNotification("图片已存入白板（上次看的第 " + CurrentWhiteboardIndex + " 页）");
            }
            catch (Exception ex)
            {
                ShowNotification("打开图片失败：" + ex.Message);
            }
        }

        #endregion

        #region 系统截图对接（剪贴板监听）

        // ===== Win32 互操作：剪贴板监听 + 窗口置前 =====
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>剪贴板内容变化消息（AddClipboardFormatListener 注册后收到）</summary>
        private const int WM_CLIPBOARDUPDATE = 0x031D;

        /// <summary>是否正在等待系统截图结果（true=已调起系统截图，等剪贴板回图）</summary>
        bool _awaitingSystemScreenshot = false;

        /// <summary>调起系统截图的时刻（Environment.TickCount）；超期的剪贴板变化不再当作截图结果</summary>
        int _systemScreenshotRequestTick = 0;

        /// <summary>等待有效期：调起后 60 秒内剪贴板出现的图片视为截图结果（用户取消时无图入板，过期防误插）</summary>
        const int SystemScreenshotTimeoutMs = 60_000;

        /// <summary>超时看门狗：到期仍未收到截图（用户取消/忘记）→ 解除等待；隐藏界面截图时还要恢复窗口</summary>
        System.Windows.Threading.DispatcherTimer _screenshotTimeoutTimer;

        /// <summary>为截图隐藏了主窗口（true=收图或超时后要 Show 恢复）</summary>
        bool _hiddenForScreenshot = false;

        /// <summary>
        /// 挂剪贴板监听（MainWindow 构造流程调用）。
        /// 用消息钩子（HwndSource.AddHook）而不是轮询，开销为零、响应即时。
        /// SourceInitialized 之后窗口句柄才可用。
        /// </summary>
        internal void InitSystemScreenshotHook()
        {
            try
            {
                SourceInitialized += (s, e) =>
                {
                    try
                    {
                        var source = (HwndSource)PresentationSource.FromVisual(this);
                        source?.AddHook(WndProc);
                        AddClipboardFormatListener(new WindowInteropHelper(this).Handle);
                    }
                    catch { }
                };
            }
            catch { }
        }

        /// <summary>窗口消息钩子：只关心剪贴板变化消息，其余放行</summary>
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg == WM_CLIPBOARDUPDATE && _awaitingSystemScreenshot)
            {
                OnSystemScreenshotClipboard();
            }
            return IntPtr.Zero;
        }

        /// <summary>
        /// 系统截图结果到达剪贴板：稍等一拍取图（截图工具刚写完剪贴板，立即读偶发取空），
        /// 拉回本窗口焦点（系统截图后焦点不归还）→ 存入白板抽屉。
        /// </summary>
        private void OnSystemScreenshotClipboard()
        {
            _awaitingSystemScreenshot = false;
            StopScreenshotTimeoutTimer();
            // 过期保护：太久之前的调起请求不再消费（用户可能中途自己复制了别的东西）
            if (Environment.TickCount - _systemScreenshotRequestTick > SystemScreenshotTimeoutMs) return;

            var timer = new System.Threading.Timer(_ => Dispatcher.Invoke(() =>
            {
                try
                {
                    if (!Clipboard.ContainsImage()) return;
                    var src = Clipboard.GetImage();
                    if (src == null) return;
                    // 隐藏界面截图：先把主窗口恢复回来
                    if (_hiddenForScreenshot)
                    {
                        _hiddenForScreenshot = false;
                        Show();
                    }
                    // 截图后系统不还焦点：主动把窗口拉回前台，用户不用再点一下窗口
                    try { SetForegroundWindow(new WindowInteropHelper(this).Handle); } catch { }
                    // 存入白板抽屉（不管当前模式，都挂在上次看的那一页）
                    ImageLayer_AddImage(src);
                    if (currentMode != 0)
                        ShowNotification("截图已插入白板第 " + CurrentWhiteboardIndex + " 页");
                    else
                        ShowNotification("截图已存入白板（上次看的第 " + CurrentWhiteboardIndex + " 页）");
                }
                catch { }
            }), null, 250, System.Threading.Timeout.Infinite);
        }

        /// <summary>启动截图超时看门狗（60 秒没等到图 = 用户取消，解除等待并恢复隐藏的窗口）</summary>
        private void StartScreenshotTimeoutTimer()
        {
            StopScreenshotTimeoutTimer();
            _screenshotTimeoutTimer = new System.Windows.Threading.DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(SystemScreenshotTimeoutMs)
            };
            _screenshotTimeoutTimer.Tick += (s, e) =>
            {
                StopScreenshotTimeoutTimer();
                _awaitingSystemScreenshot = false;
                if (_hiddenForScreenshot)
                {
                    _hiddenForScreenshot = false;
                    Show();
                }
            };
            _screenshotTimeoutTimer.Start();
        }

        private void StopScreenshotTimeoutTimer()
        {
            try { _screenshotTimeoutTimer?.Stop(); } catch { }
            _screenshotTimeoutTimer = null;
        }

        /// <summary>
        /// 用户已回到画布操作（落笔）→ 视为放弃截图等待，防止之后自己复制的图被误插。
        /// 由 MainWindow_StylusDown 调用（窗口隐藏期间收不到本事件，只在可见时生效，恰好正确）。
        /// </summary>
        internal void CancelAwaitingSystemScreenshot()
        {
            if (!_awaitingSystemScreenshot) return;
            _awaitingSystemScreenshot = false;
            StopScreenshotTimeoutTimer();
        }

        /// <summary>
        /// 调起 Windows 系统截图（ms-screenclip: 协议 = Win+Shift+S 的截图工具）。
        /// 用户框选完成后图进剪贴板 → OnSystemScreenshotClipboard 收图存入白板。
        /// 系统 UI（顶部模式条/右下通知）无法隐藏，但误选其他模式也只是换了框选方式，结果同样是图。
        /// </summary>
        private void LaunchSystemRegionScreenshot()
        {
            try
            {
                _awaitingSystemScreenshot = true;
                _systemScreenshotRequestTick = Environment.TickCount;
                StartScreenshotTimeoutTimer();
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo("ms-screenclip:")
                {
                    UseShellExecute = true
                });
                ShowNotification("已调出系统截图：框选完成后自动存入白板");
            }
            catch (Exception ex)
            {
                _awaitingSystemScreenshot = false;
                StopScreenshotTimeoutTimer();
                ShowNotification("无法调起系统截图（需 Win10 1809+）：" + ex.Message);
            }
        }

        #endregion
    }
}
