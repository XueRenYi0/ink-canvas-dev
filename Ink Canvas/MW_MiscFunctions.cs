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
using System.Windows.Threading;
using Application = System.Windows.Application;
using File = System.IO.File;
using MessageBox = System.Windows.MessageBox;
using Path = System.IO.Path;
using Point = System.Windows.Point;
using Timer = System.Timers.Timer;

namespace Ink_Canvas
{
    /// <summary>MainWindow 分部类：杂项（自启/主题/截图/通知/工具）（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Functions

        /// <summary>
        /// 传入域名返回对应的IP
        /// </summary>
        /// <param name="domainName">域名</param>
        /// <returns></returns>
        public static string GetIp(string domainName)
        {
            domainName = domainName.Replace("http://", "").Replace("https://", "");
            IPHostEntry hostEntry = Dns.GetHostEntry(domainName);
            IPEndPoint ipEndPoint = new IPEndPoint(hostEntry.AddressList[0], 0);
            return ipEndPoint.Address.ToString();
        }

        public static string GetWebClient(string url)
        {
            HttpWebRequest myrq = (HttpWebRequest)WebRequest.Create(url);

            myrq.Proxy = null;
            myrq.KeepAlive = false;
            myrq.Timeout = 30 * 1000;
            myrq.Method = "Get";
            myrq.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            myrq.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 UBrowser/6.2.4098.3 Safari/537.36";

            HttpWebResponse myrp;
            try
            {
                myrp = (HttpWebResponse)myrq.GetResponse();
            }
            catch (WebException ex)
            {
                myrp = (HttpWebResponse)ex.Response;
            }

            if (myrp?.StatusCode != HttpStatusCode.OK)
            {
                return "null";
            }

            using (StreamReader sr = new StreamReader(myrp.GetResponseStream()))
            {
                return sr.ReadToEnd();
            }
        }

        #region 开机自启
        /// <summary>
        /// 开机自启创建
        /// </summary>
        /// <param name="exeName">程序名称</param>
        /// <returns></returns>
        public static bool StartAutomaticallyCreate(string exeName)
        {
            try
            {
                WshShell shell = new WshShell();
                IWshShortcut shortcut = (IWshShortcut)shell.CreateShortcut(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + exeName + ".lnk");
                //设置快捷方式的目标所在的位置(源程序完整路径)
                shortcut.TargetPath = System.Windows.Forms.Application.ExecutablePath;
                //应用程序的工作目录
                //当用户没有指定一个具体的目录时，快捷方式的目标应用程序将使用该属性所指定的目录来装载或保存文件。
                shortcut.WorkingDirectory = System.Environment.CurrentDirectory;
                //目标应用程序窗口类型(1.Normal window普通窗口,3.Maximized最大化窗口,7.Minimized最小化)
                shortcut.WindowStyle = 1;
                //快捷方式的描述
                shortcut.Description = exeName + "_Ink";
                //设置快捷键(如果有必要的话.)
                //shortcut.Hotkey = "CTRL+ALT+D";
                shortcut.Save();
                return true;
            }
            catch (Exception) { }
            return false;
        }

        /// <summary>
        /// 开机自启删除
        /// </summary>
        /// <param name="exeName">程序名称</param>
        /// <returns></returns>
        public static bool StartAutomaticallyDel(string exeName)
        {
            try
            {
                System.IO.File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\" + exeName + ".lnk");
                return true;
            }
            catch (Exception) { }
            return false;
        }
        #endregion

        #region Auto Theme

        Color FloatBarForegroundColor = Color.FromRgb(102, 102, 102);
        /// <summary>
        /// 应用界面主题。
        /// ★ 幂等 + 就地替换：App.xaml 已内置 Light/Dark 与三个图片字典，
        /// 原实现每次都无条件再 Add 一份 —— 启动即多挂 4 个重复字典，而本方法在启动与
        /// 每次系统主题变化时都会被调用，链只会越来越长（DynamicResource 是线性扫描：
        /// 启动 12 个 → 每切一次主题再 +4，无上限）。现在已存在就不重复加载，恒为 8 个。
        /// 注：实测重复 Add **并不会**重新解析 XAML（WPF 对同 Source 的字典有内部缓存、
        /// 复用同一实例），所以本改动修的是"查找链长度 + 无限增长"，
        /// **不改变启动耗时**（AB 各 3 次实测无显著差异，勿误记为提速项）。
        /// </summary>
        private void SetTheme(string theme)
        {
            if (theme != "Light" && theme != "Dark") return;

            var merged = Application.Current.Resources.MergedDictionaries;

            // 主题字典：Light / Dark 同时只应存在一个。
            // 用"就地替换"而不是"删了再加"——字典在链中的位置决定 DynamicResource 的查找成本，
            // 删了再加会把它挪到链尾，反而变慢。
            string target = theme == "Light" ? LightSheet : DarkSheet;
            string other = theme == "Light" ? DarkSheet : LightSheet;
            if (!ReplaceSheet(merged, other, target)) EnsureSheet(merged, target);

            // 三个图片字典与主题无关：App.xaml 已加载则一次都不碰（缺失才补，保证健壮）
            EnsureSheet(merged, DrawShapeSheet);
            EnsureSheet(merged, SeewoSheet);
            EnsureSheet(merged, IconSheet);

            ThemeManager.SetRequestedTheme(window, theme == "Light" ? ElementTheme.Light : ElementTheme.Dark);

            FloatBarForegroundColor = (Color)Application.Current.FindResource("FloatBarForegroundColor");
            SetSelectToolColor(FloatBarForegroundColor);

            // 字典总数与来源留痕（低频事件，不刷屏）：去重生效后应为固定值，不随切换次数增长
            string sheetList = string.Join(" | ", merged.Select(d => d.Source == null ? "(inline)" : d.Source.OriginalString));
            LogHelper.WriteLogToFile($"[Theme] 应用 {theme}，MergedDictionaries={merged.Count}：[{sheetList}]", LogHelper.LogType.Event);
        }

        #region 主题资源字典清单（集中在此，避免路径字符串散落）

        const string LightSheet = "Resources/Styles/Light.xaml";
        const string DarkSheet = "Resources/Styles/Dark.xaml";
        const string DrawShapeSheet = "Resources/DrawShapeImageDictionary.xaml";
        const string SeewoSheet = "Resources/SeewoImageDictionary.xaml";
        const string IconSheet = "Resources/IconImageDictionary.xaml";

        /// <summary>
        /// 按 Source 找字典下标（-1 = 不存在）。
        /// 用尾部比对而非全等：App.xaml 里写的相对路径与代码里 new Uri 的写法
        /// 规范化程度可能不同；内联写法（ui:ThemeResources 等）Source 为 null，一律跳过。
        /// </summary>
        private static int IndexOfSheet(Collection<ResourceDictionary> merged, string sheet)
        {
            for (int i = 0; i < merged.Count; i++)
            {
                var source = merged[i]?.Source;
                if (source == null) continue;
                string s = source.OriginalString;
                if (s.Equals(sheet, StringComparison.OrdinalIgnoreCase) ||
                    s.EndsWith("/" + sheet, StringComparison.OrdinalIgnoreCase))
                    return i;
            }
            return -1;
        }

        /// <summary>把 other 那份原地替换成 target（保持链中位置，不改变查找成本）；找不到 other 返回 false</summary>
        private static bool ReplaceSheet(Collection<ResourceDictionary> merged, string other, string target)
        {
            int i = IndexOfSheet(merged, other);
            if (i < 0) return false;
            merged[i] = new ResourceDictionary { Source = new Uri(target, UriKind.Relative) };
            return true;
        }

        /// <summary>缺失才补（幂等：已存在则完全不碰，避免重复解析）</summary>
        private static void EnsureSheet(Collection<ResourceDictionary> merged, string sheet)
        {
            if (IndexOfSheet(merged, sheet) >= 0) return;
            merged.Add(new ResourceDictionary { Source = new Uri(sheet, UriKind.Relative) });
        }

        #endregion

        private void SystemEvents_UserPreferenceChanged(object sender, Microsoft.Win32.UserPreferenceChangedEventArgs e)
        {
            switch (Settings.Appearance.Theme)
            {
                case 0:
                    SetTheme("Light");
                    break;
                case 1:
                    SetTheme("Dark");
                    break;
                case 2:
                    if (IsSystemThemeLight()) SetTheme("Light");
                    else SetTheme("Dark");
                    break;
            }
        }

        private bool IsSystemThemeLight()
        {
            bool light = false;
            try
            {
                RegistryKey registryKey = Registry.CurrentUser;
                RegistryKey themeKey = registryKey.OpenSubKey("software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize");
                int keyValue = 0;
                if (themeKey != null)
                {
                    keyValue = (int)themeKey.GetValue("SystemUsesLightTheme");
                }
                if (keyValue == 1) light = true;
            }
            catch { }
            return light;
        }
        #endregion

        #endregion Functions

        #region Screenshot

        private void BtnScreenshot_Click(object sender, RoutedEventArgs e)
        {
            bool isHideNotification = false;
            if (sender is bool) isHideNotification = (bool)sender;

            GridNotifications.Visibility = Visibility.Collapsed;

            new Thread(new ThreadStart(() =>
            {
                Thread.Sleep(20);
                try
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                            SaveScreenShot(isHideNotification, $"{pptName}/{previousSlideID}_{DateTime.Now:HH-mm-ss}");
                        else
                            SaveScreenShot(isHideNotification);
                    });
                }
                catch
                {
                    if (!isHideNotification)
                    {
                        ShowNotification("截图保存失败");
                    }
                }

                try
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        if (inkCanvas.Visibility != Visibility.Visible || inkCanvas.Strokes.Count == 0 || !Settings.Automation.IsAutoSaveStrokesAtScreenshot) return;
                        SaveInkCanvasStrokes(false);
                    });
                }
                catch { }

                if (isHideNotification)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        BtnClear_Click(BtnClear, null);
                    });
                }
            })).Start();
        }

        private void SaveScreenShot(bool isHideNotification, string fileName = null)
        {
            var size = System.Windows.Forms.SystemInformation.PrimaryMonitorSize;
            var rc = new System.Drawing.Rectangle(new System.Drawing.Point(0, 0), new System.Drawing.Size(size.Width, size.Height));
            var bitmap = new System.Drawing.Bitmap(rc.Width, rc.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);

            //截图前取消选中并刷一帧渲染（Dispatcher 上下文内同步执行）：
            //否则选框虚线/缩放手柄/操作条会被 CopyFromScreen 一并截进保存的 PNG（脏图）
            try
            {
                CancelActiveSelection();
                //让"收起覆盖层"的布局变化真正渲染到屏幕再截图。
                //Render 优先级排在正常布局之后，相当于插队等一帧渲染完成
                System.Windows.Threading.DispatcherPriority renderPriority = System.Windows.Threading.DispatcherPriority.Render;
                Dispatcher.Invoke(new Action(() => { }), renderPriority);
            }
            catch { }

            using (System.Drawing.Graphics memoryGrahics = System.Drawing.Graphics.FromImage(bitmap))
            {
                memoryGrahics.CopyFromScreen(rc.X, rc.Y, 0, 0, rc.Size, System.Drawing.CopyPixelOperation.SourceCopy);
            }

            if (Settings.Automation.IsSaveScreenshotsInDateFolders)
            {
                if (string.IsNullOrWhiteSpace(fileName))
                    fileName = DateTime.Now.ToString("HH-mm-ss");
                var savePath =
                    $@"{Environment.GetFolderPath(Environment.SpecialFolder.MyPictures)}\Ink Canvas Screenshots\{DateTime.Now.Date:yyyyMMdd}\{fileName}.png";


                if (!Directory.Exists(Path.GetDirectoryName(savePath)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(savePath));
                }

                bitmap.Save(savePath, ImageFormat.Png);

                if (!isHideNotification)
                {
                    ShowNotification("截图成功保存至 " + savePath);
                }
            }
            else
            {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) + @"\Ink Canvas Screenshots"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                                              @"\Ink Canvas Screenshots");
                }

                bitmap.Save(Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                            @"\Ink Canvas Screenshots\" + DateTime.Now.ToString("u").Replace(':', '-') + ".png", ImageFormat.Png);

                if (!isHideNotification)
                {
                    ShowNotification("截图成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyPictures) +
                                     @"\Ink Canvas Screenshots\" + DateTime.Now.ToString("u").Replace(':', '-') + ".png");
                }
            }
        }

        #endregion

        #region Notification

        int lastNotificationShowTime = 0;
        int notificationShowTime = 2500;

        public static void ShowNewMessage(string notice, bool isShowImmediately = true)
        {
            (Application.Current?.Windows.Cast<Window>().FirstOrDefault(window => window is MainWindow) as MainWindow)?.ShowNotification(notice, isShowImmediately);
        }

        public void ShowNotification(string notice, bool isShowImmediately = true)
        {
            lastNotificationShowTime = Environment.TickCount;

            GridNotifications.Visibility = Visibility.Visible;
            //GridNotifications.Opacity = 1;
            TextBlockNotice.Text = notice;

            new Thread(new ThreadStart(() =>
            {
                Thread.Sleep(notificationShowTime + 200);
                if (Environment.TickCount - lastNotificationShowTime >= notificationShowTime)
                {
                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        GridNotifications.Visibility = Visibility.Collapsed;
                        //DoubleAnimation daV = new DoubleAnimation(1, 0, new Duration(TimeSpan.FromSeconds(0.15)));
                        //GridNotifications.BeginAnimation(UIElement.OpacityProperty, daV);

                        //new Thread(new ThreadStart(() => {
                        //    Thread.Sleep(200);
                        //    Application.Current.Dispatcher.Invoke(() =>
                        //    {
                        //        if (GridNotifications.Opacity == 0)
                        //        {
                        //            GridNotifications.Visibility = Visibility.Collapsed;
                        //            GridNotifications.Opacity = 1;
                        //        }
                        //    });
                        //})).Start();
                    });
                }
            })).Start();
        }

        private void AppendNotification(string notice)
        {
            TextBlockNotice.Text = TextBlockNotice.Text + Environment.NewLine + notice;
        }

        #endregion

        #region Tools

        private void BtnTools_Click(object sender, RoutedEventArgs e)
        {
            if (StackPanelToolButtons.Visibility == Visibility.Visible)
            {
                StackPanelToolButtons.Visibility = Visibility.Collapsed;
            }
            else
            {
                StackPanelToolButtons.Visibility = Visibility.Visible;
            }
        }

        private void BtnCountdownTimer_Click(object sender, RoutedEventArgs e)
        {
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            new CountdownTimerWindow().Show();
        }

        private void BtnRand_Click(object sender, RoutedEventArgs e)
        {
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            new RandWindow().Show();
        }

        #endregion Tools
    }
}
