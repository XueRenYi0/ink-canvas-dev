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
    /// <summary>MainWindow 分部类：窗口初始化与定时器（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Window Initialization

        public MainWindow()
        {
            InitializeComponent();

            BorderSettings.Opacity = 0;
            BorderSettings.Visibility = Visibility.Collapsed;
            StackPanelToolButtons.Visibility = Visibility.Collapsed;
            BorderDrawShape.Visibility = Visibility.Collapsed;
            GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;

            //图形面板"解挂"到主窗口根层：从悬浮条的吊挂改为自由拖动小窗口（方案B，
            //详见 MW_ShapeDrawing.cs 里的原理注释；失败时自动降级回旧吊挂方式）
            DetachShapePanelToRoot();

            if (App.StartArgs.Contains("-b")) //-b border
            {
                AllowsTransparency = false;
                WindowStyle = WindowStyle.SingleBorderWindow;
                ResizeMode = ResizeMode.CanResize;
                Background = new SolidColorBrush(StringToColor("#FFF2F2F2"));
                Topmost = false;
            }

            if (!App.StartArgs.Contains("-o")) //-old ui
            {
                GroupBoxAppearance.Visibility = Visibility.Collapsed;
                ViewBoxStackPanelMain.Visibility = Visibility.Collapsed;
                ViewBoxStackPanelShapes.Visibility = Visibility.Collapsed;
                HideSubPanels();

                ViewboxFloatingBar.Margin = new Thickness((SystemParameters.WorkArea.Width - 284) / 2, SystemParameters.WorkArea.Height - 80, -2000, -200);
            }
            else
            {
                GroupBoxAppearanceNewUI.Visibility = Visibility.Collapsed;
                ViewboxFloatingBar.Visibility = Visibility.Collapsed;
                GridForRecoverOldUI.Visibility = Visibility.Collapsed;
            }

            if (File.Exists("debug.ini")) Label.Visibility = Visibility.Visible;

            InitTimers();
            InitFloatingBarWatchdog(); // 悬浮条存活性看门狗：拖动卡死自愈 + 运行中强制可见（除非程序关闭）
            timeMachine.OnRedoStateChanged += TimeMachine_OnRedoStateChanged;
            timeMachine.OnUndoStateChanged += TimeMachine_OnUndoStateChanged;
            inkCanvas.Strokes.StrokesChanged += StrokesOnStrokesChanged;
            InitNoteScroll();
            InitImageLayer(); //页面图片层：截图/本地图片按白板页存储（见 MW_ImageLayer.cs）
            InitSystemScreenshotHook(); //系统截图对接：剪贴板监听（ms-screenclip: 框选完自动存入白板，见 MW_Capture.cs）
            PreviewMouseDown += ImageLayer_MenuCloseOnOutsideClick; //截图菜单：点窗口任意位置自动关闭
            InitCustomShapes();
            InitGraphStrokeGroupErasing(); //图形笔迹整组擦除：橡皮碰到图形任意部分即整组消失（见 MW_GraphStrokes.cs）
            InitShapeIconStyles(); //图形图标三态视觉：悬停浅灰 / 激活蓝色高亮（见 MW_ShapeDrawing.cs）

            //启动即收起悬浮工具栏，只留笑脸把手（XAML 初始 ScaleX=0；单击把手展开；PPT 放映仍会自动展开）
            //-o（old ui）保持老行为：启动即展开
            if (App.StartArgs.Contains("-o"))
            {
                SetBorderFloatingBarMainControlsVisibility(true, false);
            }

            Microsoft.Win32.SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;

            //窗口句柄就绪后挂 Win32 消息钩子：监听设备插拔（WM_DEVICECHANGE），
            //手写板/触摸屏被移除时刷新 WPF 手写笔设备表，防止输入通道悬死（详见下方 DeviceChangeWndProc）
            SourceInitialized += MainWindow_SourceInitialized;
        }

        #endregion

        #region Timer

        Timer timerCheckPPT = new Timer();
        Timer timerKillProcess = new Timer();

        private void InitTimers()
        {
            timerCheckPPT.Elapsed += TimerCheckPPT_Elapsed;
            timerCheckPPT.Interval = 1000;

            timerKillProcess.Elapsed += TimerKillProcess_Elapsed;
            timerKillProcess.Interval = 5000;
        }

        private void TimerKillProcess_Elapsed(object sender, ElapsedEventArgs e)
        {
            try
            {
                // 希沃相关： easinote swenserver RemoteProcess EasiNote.MediaHttpService smartnote.cloud EasiUpdate smartnote EasiUpdate3 EasiUpdate3Protect SeewoP2P CefSharp.BrowserSubprocess SeewoUploadService
                string arg = "/F";
                if (Settings.Automation.IsAutoKillPptService)
                {
                    Process[] processes = Process.GetProcessesByName("PPTService");
                    if (processes.Length > 0)
                    {
                        arg += " /IM PPTService.exe";
                    }
                    processes = Process.GetProcessesByName("SeewoIwbAssistant");
                    if (processes.Length > 0)
                    {
                        arg += " /IM SeewoIwbAssistant.exe" +
                            " /IM Sia.Guard.exe";
                    }
                }
                if (Settings.Automation.IsAutoKillEasiNote)
                {
                    Process[] processes = Process.GetProcessesByName("EasiNote");
                    if (processes.Length > 0)
                    {
                        arg += " /IM EasiNote.exe";

                    }
                }
                if (arg != "/F")
                {
                    Process p = new Process();
                    p.StartInfo = new ProcessStartInfo("taskkill", arg);
                    p.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    p.Start();

                    if (arg.Contains("EasiNote"))
                    {

                        BtnSwitch_Click(BtnSwitch, null);
                        MessageBox.Show("“希沃白板 5”已自动关闭");
                    }
                }
            }
            catch { }
        }

        #endregion Timer

        #region 设备插拔监听（手写板热移除卡死的对症修复）

        // ===== 背景 =====
        //批注中物理关闭手写板 → WPF 触摸/手写笔输入栈（wisp）收不到收尾事件，
        //该窗口的输入通道悬死：之后手写笔/触摸点不动，但鼠标和键盘正常（实测确认）。
        //对症修复：监听系统设备插拔广播（WM_DEVICECHANGE），一旦设备增删就调用
        //WPF 内部的 StylusLogic.RefreshStylusDevices() 重建手写笔设备表——
        //这是 WPF 生态里对付"设备变更后触摸/笔失灵"的经典自救手段（反射调用内部 API）。

        private const int WM_DEVICECHANGE = 0x0219;   // 设备状态变化消息
        private const int DBT_DEVICEARRIVAL = 0x8000;          // 设备接入（卷设备类才有，如U盘）
        private const int DBT_DEVICEREMOVECOMPLETE = 0x8004;    // 设备移除完成（卷设备类才有）
        private const int DBT_DEVNODES_CHANGED = 0x0007;       // 设备树变化（笼统事件，无细节）
        private DateTime _lastStylusRefreshTime = DateTime.MinValue; // 上次刷新时刻（节流用）

        /// <summary>窗口句柄就绪：挂上 Win32 消息钩子（构造器里订阅本事件）</summary>
        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            try
            {
                var source = System.Windows.Interop.HwndSource.FromHwnd(
                    new System.Windows.Interop.WindowInteropHelper(this).Handle);
                source?.AddHook(DeviceChangeWndProc);
                LogHelper.WriteLogToFile("[DeviceChange] 消息钩子已挂载", LogHelper.LogType.Event);
            }
            catch (Exception ex)
            {
                LogHelper.WriteLogToFile($"[DeviceChange] 钩子挂载失败: {ex.Message}", LogHelper.LogType.Event);
            }
        }

        /// <summary>Win32 消息处理：只关心设备插拔，其余消息原样放行</summary>
        private IntPtr DeviceChangeWndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            if (msg != WM_DEVICECHANGE) return IntPtr.Zero;

            int evt = wParam.ToInt32();
            //关键：手写板/HID/蓝牙这类普通 PnP 设备，系统对未注册通知的窗口只广播
            //笼统的 DBT_DEVNODES_CHANGED（"设备树变了"，无细节）；带详情的
            //ARRIVAL/REMOVECOMPLETE 是U盘那类卷设备的特权。三种都认，宁滥勿缺
            //（刷新本身无害 + 有 2 秒节流防抖）。
            if (evt != DBT_DEVICEARRIVAL && evt != DBT_DEVICEREMOVECOMPLETE && evt != DBT_DEVNODES_CHANGED)
                return IntPtr.Zero;

            //节流：一次插拔可能连发多条广播（各接口层各发一条），
            //2 秒内只刷一次，避免反复重建设备表
            if ((DateTime.Now - _lastStylusRefreshTime).TotalMilliseconds < 2000) return IntPtr.Zero;
            _lastStylusRefreshTime = DateTime.Now;

            LogHelper.WriteLogToFile($"[DeviceChange] 检测到设备变化(evt=0x{evt:X})，刷新手写笔设备表", LogHelper.LogType.Event);

            //延迟到 Background 优先级执行：让排在队列前面的残余输入消息
            //（比如设备消失瞬间产生的错误事件）先处理完，再重建设备表
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    RefreshStylusDevices();
                    ResetInputAfterDeviceChange(); //第二刀：释放悬死捕获 + 中止幽灵笔画（见方法注释）
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLogToFile($"[DeviceChange] 刷新失败: {ex.Message}", LogHelper.LogType.Event);
                }
            }), System.Windows.Threading.DispatcherPriority.Background);

            return IntPtr.Zero;
        }

        /// <summary>
        /// 设备插拔后的输入通道大扫除（第二刀）。
        /// 背景：书写中设备被移除 → StylusUp 永不送达 → WPF 认为"笔还按着"：
        /// 捕获被攥住、InkCanvas 悬着一条"幽灵笔画"。之后笔的悬停/移动正常（不需要捕获），
        /// 但所有按下都被当旧笔画吞掉——实测表现为：光标会动、能悬停变形，就是点不动。
        /// 处理：①释放各层输入捕获 ②复位停顿拉直悬死态 ③清触摸脏 ID ④强制中止幽灵笔画。
        /// （副作用可忽略：若恰好正在正常书写时插拔U盘，当前这一笔会被截断，属于罕见场景）
        /// </summary>
        private void ResetInputAfterDeviceChange()
        {
            //①释放输入捕获：InkCanvas（书写中捕获笔）、选区覆盖层（拖动中捕获鼠标）
            try { inkCanvas.ReleaseStylusCapture(); } catch { }
            try { inkCanvas.ReleaseMouseCapture(); } catch { }
            try { GridInkCanvasSelectionCover.ReleaseStylusCapture(); } catch { }
            try { GridInkCanvasSelectionCover.ReleaseMouseCapture(); } catch { }
            try { if (Mouse.Captured != null) Mouse.Captured.ReleaseMouseCapture(); } catch { }

            //②停顿拉直若处于 armed 悬死态（模式卡在 None 等抬笔），先复位（内部会恢复原模式）
            try { ResetLineAssist(); } catch { }

            //③触摸 ID 表清空：设备消失时 TouchUp 不送达，残留的脏 ID
            //会让之后每次单指触摸都被误判为"多指手势"而切 None（写不了字）
            try { dec.Clear(); } catch { }

            //④强制中止幽灵笔画：来回切一次编辑模式，InkCanvas 会丢弃
            //进行中的动态笔画并复位内部状态机（经典自救手段）
            var saved = inkCanvas.EditingMode;
            if (saved != InkCanvasEditingMode.None)
            {
                inkCanvas.EditingMode = InkCanvasEditingMode.None;
                inkCanvas.EditingMode = saved;
            }

            LogHelper.WriteLogToFile("[DeviceChange] 输入捕获已重置（第二刀）", LogHelper.LogType.Event);
        }

        /// <summary>
        /// 反射调用 WPF 内部 StylusLogic.RefreshStylusDevices()：重建手写笔/触摸设备表。
        /// 等效于让 WPF "重新扫描所有输入设备"——设备热插拔后输入通道悬死的对症药。
        /// （内部 API 无公开入口，只能反射；调用失败不影响程序其他功能）
        /// </summary>
        private static void RefreshStylusDevices()
        {
            var inputManager = System.Windows.Input.InputManager.Current;
            //第一跳：InputManager.StylusLogic（internal 属性）→ 拿到手写笔逻辑对象
            object stylusLogic = typeof(System.Windows.Input.InputManager).InvokeMember(
                "StylusLogic",
                System.Reflection.BindingFlags.GetProperty | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, inputManager, null);
            if (stylusLogic == null) return;

            //第二跳：StylusLogic.RefreshStylusDevices()（internal 方法）→ 重建设备表
            stylusLogic.GetType().InvokeMember(
                "RefreshStylusDevices",
                System.Reflection.BindingFlags.InvokeMethod | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                null, stylusLogic, null);

            LogHelper.WriteLogToFile("[DeviceChange] 手写笔设备表已刷新", LogHelper.LogType.Event);
        }

        #endregion
    }
}
