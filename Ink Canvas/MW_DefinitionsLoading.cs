using Ink_Canvas.Helpers;
using iNKORE.UI.WPF.Modern;
using iNKORE.UI.WPF.Modern.Helpers;
using IWshRuntimeLibrary;
using Microsoft.Office.Interop.PowerPoint;
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
    /// <summary>MainWindow 分部类：字段定义与初始化加载（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Definations and Loading

        /// <summary>
        /// 自定义图标的统一 StrokeThickness 基准。
        /// 来源：同排 SeewoImageSource.PPTExitNormal，viewBox≈30 宽，矩形框 Pen Thickness=2 → 相对 2/30 ≈ 6.67%
        /// 自定义图标 Grid Width=20 → 20 × 6.67% = 1.22（已与 Undo/Redo Symbol 原生笔画粗细对比校准通过）
        /// </summary>
        internal const double IconStrokeThickness = 1.22;

        public static Settings Settings = new Settings();
        public static string settingsFileName = "Settings.json";
        bool isLoaded = false;
        //bool isAutoUpdateEnabled = false;

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            // ★ 2026-09-11 移除：原此处有一个后台线程，向上游作者服务器 http://ink.wxriw.cn:1957
            // 上报本机版本号，并根据返回内容决定是否弹「专版」欢迎窗 / 通知框、维护 Versions.ini。
            // 实测：该服务器虽返回 200，但根路径与 /?verinfo= 都只回字面量 "null"，
            // 判断条件 Contains("Special Version") 恒为 false → 所有分支永不触发，每次启动纯空转；
            // 且会向第三方泄露本机版本信息，并存在「二中专版」品牌污染风险（服务器一旦返回
            // Special Version，界面会突然出现“二中专版”字样与 logo2.png）。故整块移除。
            // 连带移除：MainWindow.xaml 的 GroupBoxMASEZVersion 及 logo/logo2/text/textCN 四张专版图片。
            // 连带移除：WelcomeWindow.xaml / .xaml.cs（首次运行设置向导，唯一调用点就是上面这块）。
            //      向导里的功能并未丢失：推荐设置仍在设置页的「恢复推荐设置」按钮
            //      （MW_Settings.cs 的 BtnResetToSuggestion_Click → SetSettingsToRecommendation），
            //      开机自启仍在 MW_Settings.cs:51 的 StartAutomaticallyCreate("InkCanvas")。
            StartupProfiler.Mark("Window_Loaded 开始");

            using (StartupProfiler.Measure("loadPenCanvas()")) loadPenCanvas();

            //加载设置
            using (StartupProfiler.Measure("LoadSettings()")) LoadSettings();

            // 笔设置面板：订阅事件 + 注入颜色，再按配置恢复笔种类/笔宽（含荧光笔倍率）
            using (StartupProfiler.Measure("笔设置面板初始化")) { InitPenSettingsPanel(); ApplyLoadedPenSettings(); }

            // 橡皮设置面板：订阅事件（擦除方式/大小/滑动清屏）
            using (StartupProfiler.Measure("橡皮设置面板初始化")) { InitEraserSettingsPanel(); ApplyLoadedEraserSettings(); }

            // 选择方式：订阅面板事件 + 拖选拦截，再按配置恢复（矩形框选/自由选择）
            using (StartupProfiler.Measure("选择方式初始化")) { InitSelectionMode(); ApplyLoadedSelectionMode(); }

            // 初始化动态快捷键（此时窗口句柄已就绪，避免上次 bc673dd 在构造函数注册全局热键崩溃的坑）
            using (StartupProfiler.Measure("InitDynamicShortcuts()")) InitDynamicShortcuts();

            // 注册全局逃生热键 Ctrl+Alt+Shift+R（手写板卡死时鼠标/触摸全灭，键盘是唯一活通道）
            using (StartupProfiler.Measure("InitGlobalEscapeHotkey()")) InitGlobalEscapeHotkey();

            // 安全重启的墨迹恢复（-restore 参数 + recovery.icstk 存在时自动加载）
            using (StartupProfiler.Measure("TryRestoreStrokesOnStartup()")) TryRestoreStrokesOnStartup();

            if (Environment.Is64BitProcess)
            {
                GroupBoxInkRecognition.Visibility = Visibility.Collapsed;
            }

            using (StartupProfiler.Measure("主题初始化（SetTheme + 资源字典）"))
            {
                ThemeManager.Current.ApplicationTheme = ApplicationTheme.Light;
                SystemEvents_UserPreferenceChanged(null, null);
            }

            TextBlockVersion.Text = Assembly.GetExecutingAssembly().GetName().Version.ToString();
            LogHelper.WriteLogToFile("Ink Canvas Loaded", LogHelper.LogType.Event);

            isLoaded = true;

            // ★ 启动提速（2026-09-11）：PreloadIALibrary() 原本在这里**同步**执行，
            // 实测 270 ms / 占启动耗时 12% —— 它只是给墨迹分析器做一次空跑预热
            // （避免首次墨迹识别卡顿），与"界面能不能用"毫无关系，却让用户白等。
            // 改为界面就绪后、消息队列空闲时再跑：仍是 UI 线程（不引入任何跨线程 /
            // COM 单元风险），但已经不计入"等待界面出现"的时间。
            // 加 try-catch：预热失败不影响任何功能（首次真正识别时库会自己初始化）。
            Dispatcher.BeginInvoke(DispatcherPriority.ApplicationIdle, (Action)(() =>
            {
                try
                {
                    var sw = Stopwatch.StartNew();
                    PreloadIALibrary();
                    sw.Stop();
                    LogHelper.WriteLogToFile($"[Startup] IA 预热已在界面就绪后完成（{sw.ElapsedMilliseconds} ms，原先同步阻塞启动）", LogHelper.LogType.Event);
                }
                catch { }
            }));

            StartupProfiler.Dump("启动完成（界面已就绪）");
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            LogHelper.WriteLogToFile("Ink Canvas closing", LogHelper.LogType.Event);
            if (!CloseIsFromButton)
            {
                e.Cancel = true;
                if (MessageBox.Show("是否继续关闭 Inkboard 画板，这将丢失当前未保存的工作。", "Inkboard 画板", MessageBoxButton.OKCancel, MessageBoxImage.Warning) == MessageBoxResult.OK)
                {
                    if (MessageBox.Show("真的狠心关闭 Inkboard 画板吗？", "Inkboard 画板", MessageBoxButton.OKCancel, MessageBoxImage.Error) == MessageBoxResult.OK)
                    {
                        if (MessageBox.Show("是否取消关闭 Inkboard 画板？", "Inkboard 画板", MessageBoxButton.OKCancel, MessageBoxImage.Error) != MessageBoxResult.OK)
                        {
                            e.Cancel = false;
                        }
                    }
                }
            }
            if (e.Cancel)
            {
                LogHelper.WriteLogToFile("Ink Canvas closing cancelled", LogHelper.LogType.Event);
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            LogHelper.WriteLogToFile("Ink Canvas closed", LogHelper.LogType.Event);
            //释放函数识别的数学面板 COM 资源（MW_MathGraph.cs）
            ReleaseMathInputPanel();
            //主动退出（退出/重启按钮路径）：立即终止进程释放单实例互斥锁。
            //否则后台线程收尾期间进程残留持锁，用户紧接着双击 exe 会误报"已有一个程序实例正在运行"（日志 14:13:08 实证，等17秒才恢复）
            if (CloseIsFromButton)
            {
                Environment.Exit(0);
            }
        }

        private static void PreloadIALibrary()
        {
            GC.KeepAlive(typeof(InkAnalyzer));
            GC.KeepAlive(typeof(AnalysisAlternate));
            GC.KeepAlive(typeof(InkDrawingNode));
            var analyzer = new InkAnalyzer();
            analyzer.AddStrokes(new StrokeCollection() {
                new Stroke(new StylusPointCollection() {
                    new StylusPoint(114,514),
                    new StylusPoint(191,9810),
                    new StylusPoint(7,21),
                    new StylusPoint(123,789),
                })
            });
            analyzer.Analyze();
        }

        private void LoadSettings(bool isStartup = true)
        {
            if (File.Exists(App.RootPath + settingsFileName))
            {
                try
                {
                    string text = File.ReadAllText(App.RootPath + settingsFileName);
                    Settings = JsonConvert.DeserializeObject<Settings>(text);
                }
                catch { }
            }

            if (Settings.Startup.IsAutoEnterModeFinger)
            {
                ToggleSwitchModeFinger.IsOn = true;
                ToggleSwitchAutoEnterModeFinger.IsOn = true;
            }
            else
            {
                ToggleSwitchAutoEnterModeFinger.IsOn = false;
            }
            // 自动更新开关回显（新字段有默认值 true，老配置文件缺字段也安全）
            ToggleSwitchAutoCheckUpdate.IsOn = Settings.Startup.IsAutoCheckUpdate;
            if (Settings.Startup.IsAutoHideCanvas)
            {
                if (isStartup)
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }
                ToggleSwitchAutoHideCanvas.IsOn = true;
            }
            else
            {
                if (isStartup)
                {
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                    BtnHideInkCanvas_Click(BtnHideInkCanvas, null);
                }
                ToggleSwitchAutoHideCanvas.IsOn = false;
            }

            if (Settings.Appearance.IsShowEraserButton)
            {
                BtnErase.Visibility = Visibility.Visible;
                ToggleSwitchShowButtonEraser.IsOn = true;
            }
            else
            {
                BtnErase.Visibility = Visibility.Collapsed;
                ToggleSwitchShowButtonEraser.IsOn = false;
            }
            if (Settings.Appearance.IsShowExitButton)
            {
                BtnExit.Visibility = Visibility.Visible;
                ToggleSwitchShowButtonExit.IsOn = true;
            }
            else
            {
                BtnExit.Visibility = Visibility.Collapsed;
                ToggleSwitchShowButtonExit.IsOn = false;
            }

            PptNavigationBtn.Visibility =
                Settings.PowerPointSettings.IsShowPPTNavigation ? Visibility.Visible : Visibility.Collapsed;
            ToggleSwitchShowButtonPPTNavigation.IsOn = Settings.PowerPointSettings.IsShowPPTNavigation;

            ComboBoxTheme.SelectedIndex = Settings.Appearance.Theme;
            if (Settings.Appearance.IsShowHideControlButton)
            {
                BtnHideControl.Visibility = Visibility.Visible;
                ToggleSwitchShowButtonHideControl.IsOn = true;
            }
            else
            {
                BtnHideControl.Visibility = Visibility.Collapsed;
                ToggleSwitchShowButtonHideControl.IsOn = false;
            }
            if (Settings.Appearance.IsShowLRSwitchButton)
            {
                BtnSwitchSide.Visibility = Visibility.Visible;
                ToggleSwitchShowButtonLRSwitch.IsOn = true;
            }
            else
            {
                BtnSwitchSide.Visibility = Visibility.Collapsed;
                ToggleSwitchShowButtonLRSwitch.IsOn = false;
            }
            if (Settings.Appearance.IsShowModeFingerToggleSwitch)
            {
                StackPanelModeFinger.Visibility = Visibility.Visible;
                ToggleSwitchShowButtonModeFinger.IsOn = true;
            }
            else
            {
                StackPanelModeFinger.Visibility = Visibility.Collapsed;
                ToggleSwitchShowButtonModeFinger.IsOn = false;
            }
            if (Settings.Appearance.IsTransparentButtonBackground)
            {
                BtnExit.Background = new SolidColorBrush(StringToColor("#7F909090"));
            }
            else
            {
                if (BtnSwitchTheme.Content.ToString() == "深色")
                {
                    //Light
                    BtnExit.Background = new SolidColorBrush(StringToColor("#FFCCCCCC"));
                }
                else
                {
                    //Dark
                    BtnExit.Background = new SolidColorBrush(StringToColor("#FF555555"));
                }
            }

            if (Settings.PowerPointSettings.PowerPointSupport)
            {
                ToggleSwitchSupportPowerPoint.IsOn = true;
                timerCheckPPT.Start();
            }
            else
            {
                ToggleSwitchSupportPowerPoint.IsOn = false;
                timerCheckPPT.Stop();
            }
            if (Settings.PowerPointSettings.IsShowCanvasAtNewSlideShow)
            {
                ToggleSwitchShowCanvasAtNewSlideShow.IsOn = true;
            }
            else
            {
                ToggleSwitchShowCanvasAtNewSlideShow.IsOn = false;
            }

            if (Settings.Gesture == null)
            {
                Settings.Gesture = new Gesture();
            }
            if (Settings.Gesture.IsDisableLockSmithByDefault)
            {
                ToggleSwitchDisableLockSmithByDefault.IsOn = true;
                _lockSmith = false;
            }
            else
            {
                ToggleSwitchDisableLockSmithByDefault.IsOn = false;
                _lockSmith = true;
            }
            UpdateGestureLockIcon();
            if (Settings.Gesture.IsEnableTwoFingerZoom)
            {
                ToggleSwitchEnableTwoFingerZoom.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableTwoFingerZoom.IsOn = false;
            }
            if (Settings.Gesture.IsEnableTwoFingerTranslate)
            {
                ToggleSwitchEnableTwoFingerTranslate.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableTwoFingerTranslate.IsOn = false;
            }
            if (Settings.Gesture.IsEnableTwoFingerRotation)
            {
                ToggleSwitchEnableTwoFingerRotation.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableTwoFingerRotation.IsOn = false;
            }
            if (Settings.Gesture.IsEnableTwoFingerRotationOnSelection)
            {
                ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableTwoFingerRotationOnSelection.IsOn = false;
            }
            if (Settings.PowerPointSettings.IsEnableTwoFingerGestureInPresentationMode)
            {
                ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableTwoFingerGestureInPresentationMode.IsOn = false;
            }
            if (Settings.PowerPointSettings.IsEnableFingerGestureSlideShowControl)
            {
                ToggleSwitchEnableFingerGestureSlideShowControl.IsOn = true;
            }
            else
            {
                ToggleSwitchEnableFingerGestureSlideShowControl.IsOn = false;
            }

            if (File.Exists(Environment.GetFolderPath(Environment.SpecialFolder.Startup) + "\\InkCanvas" + ".lnk"))
            {
                ToggleSwitchRunAtStartup.IsOn = true;
            }

            if (Settings.Canvas != null)
            {
                drawingAttributes.Height = Settings.Canvas.InkWidth;
                drawingAttributes.Width = Settings.Canvas.InkWidth;

                InkWidthSlider.Value = Settings.Canvas.InkWidth * 2;

                if (Settings.Canvas.IsShowCursor)
                {
                    ToggleSwitchShowCursor.IsOn = true;
                    inkCanvas.ForceCursor = true;
                }
                else
                {
                    ToggleSwitchShowCursor.IsOn = false;
                    inkCanvas.ForceCursor = false;
                }

                ComboBoxPenStyle.SelectedIndex = Settings.Canvas.InkStyle;

                ComboBoxEraserSize.SelectedIndex = Settings.Canvas.EraserSize;

                ComboBoxHyperbolaAsymptoteOption.SelectedIndex = (int)Settings.Canvas.HyperbolaAsymptoteOption;
            }
            else
            {
                Settings.Canvas = new Canvas();
            }

            if (Settings.Automation != null)
            {
                if (Settings.Automation.IsAutoKillEasiNote || Settings.Automation.IsAutoKillPptService)
                {
                    timerKillProcess.Start();
                }
                else
                {
                    timerKillProcess.Stop();
                }

                if (Settings.Automation.IsAutoKillEasiNote)
                {
                    ToggleSwitchAutoKillEasiNote.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoKillEasiNote.IsOn = false;
                }

                if (Settings.Automation.IsAutoClearWhenExitingWritingMode)
                {
                    ToggleSwitchClearExitingWritingMode.IsOn = true;
                }
                else
                {
                    ToggleSwitchClearExitingWritingMode.IsOn = false;
                }


                if (Settings.Automation.IsAutoSaveStrokesAtClear)
                {
                    ToggleSwitchAutoSaveStrokesAtClear.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoSaveStrokesAtClear.IsOn = false;
                }



                if (Settings.Automation.IsAutoKillPptService)
                {
                    ToggleSwitchAutoKillPptService.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoKillPptService.IsOn = false;
                }

                if (Settings.Automation.IsSaveScreenshotsInDateFolders)
                {
                    ToggleSwitchSaveScreenshotsInDateFolders.IsOn = true;
                }
                else
                {
                    ToggleSwitchSaveScreenshotsInDateFolders.IsOn = false;
                }

                if (Settings.Automation.IsAutoSaveStrokesAtScreenshot)
                {
                    ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoSaveStrokesAtScreenshot.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsAutoSaveStrokesInPowerPoint)
                {
                    ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoSaveStrokesInPowerPoint.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsNotifyPreviousPage)
                {
                    ToggleSwitchNotifyPreviousPage.IsOn = true;
                }
                else
                {
                    ToggleSwitchNotifyPreviousPage.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsNotifyHiddenPage)
                {
                    ToggleSwitchNotifyHiddenPage.IsOn = true;
                }
                else
                {
                    ToggleSwitchNotifyHiddenPage.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsNoClearStrokeOnSelectWhenInPowerPoint)
                {
                    ToggleSwitchNoStrokeClearInPowerPoint.IsOn = true;
                }
                else
                {
                    ToggleSwitchNoStrokeClearInPowerPoint.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsShowStrokeOnSelectInPowerPoint)
                {
                    ToggleSwitchShowStrokeOnSelectInPowerPoint.IsOn = true;
                }
                else
                {
                    ToggleSwitchShowStrokeOnSelectInPowerPoint.IsOn = false;
                }

                if (Settings.PowerPointSettings.IsSupportWPS)
                {
                    ToggleSwitchSupportWPS.IsOn = true;
                }
                else
                {
                    ToggleSwitchSupportWPS.IsOn = false;
                }

                SideControlMinimumAutomationSlider.Value = Settings.Automation.MinimumAutomationStrokeNumber;

                if (Settings.Canvas.HideStrokeWhenSelecting)
                {
                    ToggleSwitchHideStrokeWhenSelecting.IsOn = true;
                }
                else
                {
                    ToggleSwitchHideStrokeWhenSelecting.IsOn = false;
                }

                if (Settings.Canvas.UsingWhiteboard)
                {
                    ToggleSwitchUsingWhiteboard.IsOn = true;
                }
                else
                {
                    ToggleSwitchUsingWhiteboard.IsOn = false;
                }
                if (Settings.Canvas.UsingWhiteboard)
                {
                    BtnSwitchTheme.Content = "深色";
                    BtnSwitchTheme_Click(null, null);
                }

                // 白板底纹：恢复上次使用的类型与间距（MW_WhiteboardPattern.cs）
                InitWhiteboardPatternUI();

                switch (Settings.Canvas.EraserType)
                {
                    case 1:
                        forcePointEraser = true;
                        break;
                    case 2:
                        forcePointEraser = false;
                        break;
                }
                // 初始化橡皮图标（匹配当前模式）
                UpdateEraserIcon();

                ComboBoxEraserType.SelectedIndex = Settings.Canvas.EraserType;

                if (Settings.PowerPointSettings.IsAutoSaveScreenShotInPowerPoint)
                {
                    ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn = true;
                }
                else
                {
                    ToggleSwitchAutoSaveScreenShotInPowerPoint.IsOn = false;
                }
            }
            else
            {
                Settings.Automation = new Automation();
            }

            if (Settings.Advanced != null)
            {
                TouchMultiplierSlider.Value = Settings.Advanced.TouchMultiplier;
                if (Settings.Advanced.IsLogEnabled)
                {
                    ToggleSwitchIsLogEnabled.IsOn = true;
                }
                else
                {
                    ToggleSwitchIsLogEnabled.IsOn = false;
                }
                if (Settings.Advanced.EraserBindTouchMultiplier)
                {
                    ToggleSwitchEraserBindTouchMultiplier.IsOn = true;
                }
                else
                {
                    ToggleSwitchEraserBindTouchMultiplier.IsOn = false;
                }

                if (Settings.Advanced.IsSpecialScreen)
                {
                    ToggleSwitchIsSpecialScreen.IsOn = true;
                }
                else
                {
                    ToggleSwitchIsSpecialScreen.IsOn = false;
                }
                TouchMultiplierSlider.Visibility = ToggleSwitchIsSpecialScreen.IsOn ? Visibility.Visible : Visibility.Collapsed;

                ToggleSwitchIsQuadIR.IsOn = Settings.Advanced.IsQuadIR;
            }
            else
            {
                Settings.Advanced = new Advanced();
            }

            if (Settings.InkToShape != null)
            {
                if (Settings.InkToShape.IsInkToShapeEnabled)
                {
                    ToggleSwitchEnableInkToShape.IsOn = true;
                }
                else
                {
                    ToggleSwitchEnableInkToShape.IsOn = false;
                }
            }
            else
            {
                Settings.InkToShape = new InkToShape();
            }
        }

        #endregion Definations and Loading
    }
}
