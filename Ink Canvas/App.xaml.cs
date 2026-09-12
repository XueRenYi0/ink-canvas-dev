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
            // 启动耗时诊断零点：对齐"进程启动"时刻（见 Helpers/StartupProfiler.cs）
            StartupProfiler.Init();

            this.Startup += new StartupEventHandler(App_Startup);
            this.DispatcherUnhandledException += App_DispatcherUnhandledException;
        }

        private void App_DispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            Ink_Canvas.MainWindow.ShowNewMessage("抱歉，出现未预期的异常，可能导致 InkClass 画板运行不稳定。\n建议保存墨迹后重启应用。", true);
            LogHelper.NewLog(e.Exception.ToString());
            e.Handled = true;
        }

        void App_Startup(object sender, StartupEventArgs e)
        {
            if (!StoreHelper.IsStoreApp) RootPath = AppDomain.CurrentDomain.SetupInformation.ApplicationBase;

            LogHelper.NewLog(string.Format("InkClass Starting (Version: {0})", Assembly.GetExecutingAssembly().GetName().Version.ToString()));

            bool ret;
            mutex = new System.Threading.Mutex(true, "Ink_Canvas", out ret);

            if (!ret && !e.Args.Contains("-m")) //-m multiple
            {
                LogHelper.NewLog("Detected existing instance");
                MessageBox.Show("已有一个程序实例正在运行");
                LogHelper.NewLog("InkClass automatically closed");
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
            "https://gh-proxy.com/https://raw.githubusercontent.com/XueRenYi0/InkClass/master/update.xml",
            "https://gh-proxy.at9.net/https://raw.githubusercontent.com/XueRenYi0/InkClass/master/update.xml",
            "https://raw.githubusercontent.com/XueRenYi0/InkClass/master/update.xml"
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
        /// <summary>本次检查是否由用户手动触发（决定"已是最新"要不要提示；自动检查保持静默不打扰）</summary>
        static bool _isManualUpdateCheck;

        public void CheckForUpdate(bool showNoUpdateTip)
        {
            _isManualUpdateCheck = showNoUpdateTip;

            // 手动检查：先清掉上一次"稍后提醒 / 跳过此版本"留下的拦截状态。
            // 不清的话库会在连请求都不发的情况下静默返回 —— 表现就是"点检查更新一点反应没有"，
            // 根因见 ClearUpdateBlockers 的注释（2026-09-12 修）。
            if (showNoUpdateTip) ClearUpdateBlockers();

            // 手动检查：立刻给个回应（否则要等几秒才有窗口，用户以为点了没反应）
            if (showNoUpdateTip) Ink_Canvas.MainWindow.ShowUpdateTip("正在检查更新…");

            // 订阅结果回调。注意：一旦订阅，AutoUpdater.NET 就**不再自行弹窗**，
            // 三种结果（有新版本 / 已是最新 / 失败）全部由 OnAutoUpdaterCheckForUpdateEvent 接管
            // —— 这是库的既定行为（官方文档：事件代码 "instead of showing the update dialog"）。
            AutoUpdater.CheckForUpdateEvent -= OnAutoUpdaterCheckForUpdateEvent;
            AutoUpdater.CheckForUpdateEvent += OnAutoUpdaterCheckForUpdateEvent;
            AutoUpdater.ApplicationExitEvent += () => Environment.Exit(0);

            foreach (var url in UpdateXmlUrls)
            {
                try
                {
                    AutoUpdater.Start(url);
                    return; // 请求已受理，真正的结果由上面的回调给出
                }
                catch (Exception ex)
                {
                    LogHelper.NewLog($"Update check failed via {url}: {ex.Message}");
                }
            }

            // 所有镜像连请求都没发出去
            if (showNoUpdateTip)
                Ink_Canvas.MainWindow.ShowUpdateTip("检查更新失败：网络不可用或更新服务器暂不可达");
        }

        /// <summary>
        /// 清掉 AutoUpdater.NET 在"稍后提醒 / 跳过此版本"之后留下的拦截状态。
        ///
        /// 症状（2026-09-12 用户反馈）：更新弹窗里点了「延迟 30 分钟」后，再点「检查更新」完全没反应，
        /// 连 toast 都不再变化 —— 看起来像按钮坏了。
        ///
        /// 根因（AutoUpdater.NET 1.8.1，两层拦截叠加，缺一不可）：
        ///   ① 持久层：延迟时间点写在注册表
        ///      HKCU\Software\InkClass Team\InkClass · 板书白板\AutoUpdater\RemindLaterAt。
        ///      CheckUpdate() 若读到"还没到点"就 `return remindLaterAt`（返回的不是 args），
        ///      于是 StartUpdate() 只挂一个定时器、**完全不触发 CheckForUpdateEvent**；
        ///      而我们的全部 UI 反馈都在这个事件里 → 静默。
        ///      同理 SkippedVersion >= 上线版本时 CheckUpdate() 直接 `return null`，也是静默。
        ///   ② 内存定时器：UpdateForm 点"稍后提醒"会调 AutoUpdater.SetTimer()，
        ///      而 Start() 开头是 `if (Running || _remindLaterTimer != null) return;`
        ///      —— 定时器没到期之前，后续每次 Start() 都是进门就被弹回去，请求都不发。
        ///
        /// 手动检查 = 用户明确要求"现在就看有没有新版本"，两处都清；
        /// 自动检查（启动 5 秒静默检查）保持不动 —— "稍后提醒"本来就该对它生效。
        /// </summary>
        static void ClearUpdateBlockers()
        {
            try
            {
                // ① 清持久层。必须先保证 PersistenceProvider 非空：进程刚起来、还没跑过任何检查时
                //    它是 null，直接 ?. 会漏掉"上一轮运行存下来的"延迟状态（这正是用户遇到的情形——
                //    延迟是昨天点的，今天启动后手动检查依然被拦）。
                if (AutoUpdater.PersistenceProvider == null)
                    AutoUpdater.PersistenceProvider = new RegistryPersistenceProvider(GetAutoUpdaterRegistryLocation());

                // 先读出来再清：清掉的那一刻记一行日志，否则将来"到底是被什么拦住的"又要靠猜
                var blockedUntil = AutoUpdater.PersistenceProvider.GetRemindLater();
                var skippedVersion = AutoUpdater.PersistenceProvider.GetSkippedVersion();
                if (blockedUntil != null || skippedVersion != null)
                {
                    LogHelper.NewLog($"手动检查更新：清除拦截状态（延迟至 {blockedUntil}、跳过版本 {skippedVersion}）");
                }

                AutoUpdater.PersistenceProvider.SetRemindLater(null);
                AutoUpdater.PersistenceProvider.SetSkippedVersion(null);

                // ② 清内存定时器。库里没有公开 API 能取消它，只能反射拿私有静态字段
                //    （字段名 _remindLaterTimer 在 1.8.1 固定；包版本已在 csproj 里锁死）
                var timerField = typeof(AutoUpdater).GetField("_remindLaterTimer",
                    BindingFlags.NonPublic | BindingFlags.Static);
                var pendingTimer = timerField?.GetValue(null) as System.Threading.Timer;
                if (pendingTimer != null)
                {
                    pendingTimer.Dispose();
                    timerField.SetValue(null, null);
                    LogHelper.NewLog("手动检查更新：已取消等待中的延迟提醒定时器");
                }
            }
            catch (Exception ex)
            {
                // 清不掉也不能让"检查更新"本身崩掉；退化成"这次可能仍无反应"，但至少留了可查的日志
                LogHelper.NewLog($"清除更新拦截状态失败：{ex.GetType().Name}：{ex.Message}");
            }
        }

        /// <summary>
        /// 复刻 AutoUpdater.NET 内部计算注册表位置的逻辑（Software\{公司}\{标题}\AutoUpdater）。
        /// 目的是在库自己懒初始化 PersistenceProvider 之前，就能读写它那份"稍后提醒"状态。
        /// </summary>
        static string GetAutoUpdaterRegistryLocation()
        {
            var asm = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();

            var companyAttr = (AssemblyCompanyAttribute)Attribute.GetCustomAttribute(asm, typeof(AssemblyCompanyAttribute));
            string appCompany = companyAttr != null ? companyAttr.Company : "";

            string appTitle = AutoUpdater.AppTitle;
            if (string.IsNullOrEmpty(appTitle))
            {
                var titleAttr = (AssemblyTitleAttribute)Attribute.GetCustomAttribute(asm, typeof(AssemblyTitleAttribute));
                appTitle = titleAttr != null ? titleAttr.Title : asm.GetName().Name;
            }

            return !string.IsNullOrEmpty(appCompany)
                ? $@"Software\{appCompany}\{appTitle}\AutoUpdater"
                : $@"Software\{appTitle}\AutoUpdater";
        }

        /// <summary>
        /// 更新检查结果回调（订阅 CheckForUpdateEvent 后，UI 由本方法全权接管）。
        /// </summary>
        static void OnAutoUpdaterCheckForUpdateEvent(UpdateInfoEventArgs args)
        {
            try
            {
                if (args == null || args.Error != null)
                {
                    // 别只说"网络不可用"——把库返回的具体原因一并显示与记录，
                    // 否则失败原因永远查不出来（上一版就吃了这个亏：只提示笼统失败、日志里又没记）
                    string detail = args?.Error != null
                        ? $"{args.Error.GetType().Name}：{args.Error.Message}"
                        : "无法访问更新服务器（未取得更新信息）";
                    LogHelper.NewLog($"Update check failed: {detail}");
                    Ink_Canvas.MainWindow.ShowUpdateTip($"检查更新失败：{detail}");
                    return;
                }

                if (args.IsUpdateAvailable)
                {
                    // 有新版本：沿用库自带的标准更新窗口（原由库自动弹，订阅事件后需自己调）
                    AutoUpdater.ShowUpdateForm(args);
                }
                else if (_isManualUpdateCheck)
                {
                    // 手动检查且已是最新：给明确反馈（自动检查保持静默，不打扰上课）
                    Ink_Canvas.MainWindow.ShowUpdateTip($"当前已是最新版本（v{args.InstalledVersion}）");
                }
            }
            catch (Exception ex)
            {
                LogHelper.NewLog($"Update check event failed: {ex.Message}");
            }
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
