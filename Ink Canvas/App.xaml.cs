using AutoUpdaterDotNET;
using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern.Controls;
using System;
using System.Linq;
using System.Reflection;
using System.Windows;
using MessageBox = System.Windows.MessageBox;

namespace Ink_Canvas
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        System.Threading.Mutex mutex;

        public static string[] StartArgs = null;
        public static string RootPath = Environment.GetEnvironmentVariable("APPDATA") + "\\Ink Canvas\\";

        public App()
        {
            this.Startup += new StartupEventHandler(App_Startup);
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Ink_Canvas.MainWindow.ShowNewMessage("抱歉，出现未预期的异常，可能导致 Inkboard 画板运行不稳定。\n建议保存墨迹后重启应用。", true);
            LogHelper.NewLog(e.Exception.ToString());
            e.Handled = true;
        }

        void App_Startup(object sender, StartupEventArgs e)
        {
            if (!StoreHelper.IsStoreApp) RootPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            LogHelper.NewLog(string.Format("Ink Canvas Starting (Version: {0})", Assembly.GetExecutingAssembly().GetName().Version.ToString()));

            bool ret;
            mutex = new System.Threading.Mutex(true, "Ink_Canvas", out ret);

            if (!ret && !e.Args.Contains("-m")) //-m multiple
            {
                LogHelper.NewLog("Detected existing instance");
                MessageBox.Show("已有一个程序实例正在运行");
                LogHelper.NewLog("Ink Canvas automatically closed");
                Environment.Exit(0);
            }

            StartArgs = e.Args;

            if (!StoreHelper.IsStoreApp)
            {
                StartUpdateWatcher();
            }
        }

        #region 自动更新（AutoUpdater.NET + 国内镜像回退链）

        /// <summary>update.xml 所在的 GitHub raw 地址（走国内镜像加速，失败自动回退直连）</summary>
        // 镜像现状（2026-09 核实）：gh-proxy.com 日请求 3000w+ 持续活跃，主站无缓存无速度限制；
        // gh-proxy.at9.net 为备胎。镜像是公益服务有跑路风险，所以末位保留 GitHub 直连兜底。
        static readonly string[] UpdateXmlUrls =
        {
            "https://gh-proxy.com/https://raw.githubusercontent.com/XueRenYi0/ink-canvas-dev/master/update.xml",
            "https://gh-proxy.at9.net/https://raw.githubusercontent.com/XueRenYi0/ink-canvas-dev/master/update.xml",
            "https://raw.githubusercontent.com/XueRenYi0/ink-canvas-dev/master/update.xml"
        };

        /// <summary>
        /// 启动更新检查（延迟 5 秒静默进行，不阻塞主窗口出现）。
        /// update.xml 里的安装包 url 本身也走镜像，且发版时随仓库更新——
        /// 即使某镜像失效，下一版换前缀即可，用户不会卡死在老版本上。
        /// </summary>
        void StartUpdateWatcher()
        {
            try
            {
                if (!(Ink_Canvas.MainWindow.Settings?.Startup?.IsAutoCheckUpdate ?? true)) return; //用户关了自动检查

                var timer = new System.Windows.Threading.DispatcherTimer
                {
                    Interval = TimeSpan.FromSeconds(5)
                };
                timer.Tick += (s, _) =>
                {
                    timer.Stop();
                    CheckForUpdate(false);
                };
                timer.Start();
            }
            catch { }
        }

        /// <summary>
        /// 检查更新：依次尝试镜像 URL，第一个成功的生效（静默失败——网络问题不打扰用户）。
        /// 手动检查（showNoUpdateTip=true）时无新版会弹提示，自动检查不弹。
        /// 供 MainWindow 右键菜单"检查更新"复用。
        /// </summary>
        public void CheckForUpdate(bool showNoUpdateTip)
        {
            foreach (var url in UpdateXmlUrls)
            {
                try
                {
                    AutoUpdater.Start(url);
                    AutoUpdater.ApplicationExitEvent += () => Environment.Exit(0);
                    return; //Start 不抛错即视为受理（后续弹窗由库接管）
                }
                catch (Exception ex)
                {
                    LogHelper.NewLog($"Update check failed via {url}: {ex.Message}");
                }
            }
            if (showNoUpdateTip)
                MessageBox.Show("检查更新失败：网络不可用或更新服务器暂不可达。\n请稍后重试，或到 GitHub Releases 页面手动下载。",
                    "检查更新", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        #endregion

        private void ScrollViewer_PreviewMouseWheel(object sender, System.Windows.Input.MouseWheelEventArgs e)
        {
            try
            {
                if (System.Windows.Forms.SystemInformation.MouseWheelScrollLines == -1)
                    e.Handled = false;
                else
                    try
                    {
                        ScrollViewerEx SenderScrollViewer = (ScrollViewerEx)sender;
                        SenderScrollViewer.ScrollToVerticalOffset(SenderScrollViewer.VerticalOffset - e.Delta * 10 * System.Windows.Forms.SystemInformation.MouseWheelScrollLines / (double)120);
                        e.Handled = true;
                    }
                    catch
                    {
                    }
            }
            catch
            {
            }
        }
    }
}
