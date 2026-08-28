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
            timeMachine.OnRedoStateChanged += TimeMachine_OnRedoStateChanged;
            timeMachine.OnUndoStateChanged += TimeMachine_OnUndoStateChanged;
            inkCanvas.Strokes.StrokesChanged += StrokesOnStrokesChanged;
            InitNoteScroll();
            InitCustomShapes();
            InitShortcuts();

            //启动即收起悬浮工具栏，只留笑脸把手（XAML 初始 ScaleX=0；单击把手展开；PPT 放映仍会自动展开）
            //-o（old ui）保持老行为：启动即展开
            if (App.StartArgs.Contains("-o"))
            {
                SetBorderFloatingBarMainControlsVisibility(true, false);
            }

            Microsoft.Win32.SystemEvents.UserPreferenceChanged += SystemEvents_UserPreferenceChanged;
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
    }
}
