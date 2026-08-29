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
        /// 悬浮条右键菜单：打开"手写公式识别测试"窗口（可行性验证用，验证完可删）
        /// </summary>
        private void MenuItemHandwritingTest_Click(object sender, RoutedEventArgs e)
        {
            var window = new HandwritingTestWindow();
            window.Show();
        }

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
        private void SetTheme(string theme)
        {
            if (theme == "Light")
            {
                ResourceDictionary rd1 = new ResourceDictionary() { Source = new Uri("Resources/Styles/Light.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd1);

                ResourceDictionary rd2 = new ResourceDictionary() { Source = new Uri("Resources/DrawShapeImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd2);

                ResourceDictionary rd3 = new ResourceDictionary() { Source = new Uri("Resources/SeewoImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd3);

                ResourceDictionary rd4 = new ResourceDictionary() { Source = new Uri("Resources/IconImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd4);

                ThemeManager.SetRequestedTheme(window, ElementTheme.Light);

                FloatBarForegroundColor = (Color)Application.Current.FindResource("FloatBarForegroundColor");
            }
            else if (theme == "Dark")
            {
                ResourceDictionary rd1 = new ResourceDictionary() { Source = new Uri("Resources/Styles/Dark.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd1);

                ResourceDictionary rd2 = new ResourceDictionary() { Source = new Uri("Resources/DrawShapeImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd2);

                ResourceDictionary rd3 = new ResourceDictionary() { Source = new Uri("Resources/SeewoImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd3);

                ResourceDictionary rd4 = new ResourceDictionary() { Source = new Uri("Resources/IconImageDictionary.xaml", UriKind.Relative) };
                Application.Current.Resources.MergedDictionaries.Add(rd4);

                ThemeManager.SetRequestedTheme(window, ElementTheme.Dark);

                FloatBarForegroundColor = (Color)Application.Current.FindResource("FloatBarForegroundColor");
            }

            SetSelectToolColor(FloatBarForegroundColor);
            SymbolIconDelete.Foreground = new SolidColorBrush(FloatBarForegroundColor);
        }

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
