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
    /// <summary>MainWindow 分部类：白板页面控制（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Whiteboard Controls

        StrokeCollection[] strokeCollections = new StrokeCollection[101];
        bool[] whiteboadLastModeIsRedo = new bool[101];
        StrokeCollection lastTouchDownStrokeCollection = new StrokeCollection();

        int CurrentWhiteboardIndex = 1;
        int WhiteboardTotalCount = 1;
        TimeMachineHistory[][] TimeMachineHistories = new TimeMachineHistory[101][]; //最多99页，0用来存储非白板时的墨迹以便还原

        private void SaveStrokes(bool isBackupMain = false)
        {
            if (isBackupMain)
            {
                var timeMachineHistory = timeMachine.ExportTimeMachineHistory();
                TimeMachineHistories[0] = timeMachineHistory;
                timeMachine.ClearStrokeHistory();

            }
            else
            {
                var timeMachineHistory = timeMachine.ExportTimeMachineHistory();
                TimeMachineHistories[CurrentWhiteboardIndex] = timeMachineHistory;
                timeMachine.ClearStrokeHistory();
            }
        }

        private void ClearStrokes(bool isErasedByCode)
        {

            _currentCommitType = CommitReason.ClearingCanvas;
            if (isErasedByCode) _currentCommitType = CommitReason.CodeInput;
            inkCanvas.Strokes.Clear();
            _currentCommitType = CommitReason.UserInput;
        }

        private void RestoreStrokes(bool isBackupMain = false)
        {
            try
            {
                if (TimeMachineHistories[CurrentWhiteboardIndex] == null) return; //防止白板打开后不居中
                if (isBackupMain)
                {
                    timeMachine.ImportTimeMachineHistory(TimeMachineHistories[0]);
                    foreach (var item in TimeMachineHistories[0])
                    {
                        ApplyHistoryToCanvas(item);
                    }
                }
                else
                {
                    timeMachine.ImportTimeMachineHistory(TimeMachineHistories[CurrentWhiteboardIndex]);
                    foreach (var item in TimeMachineHistories[CurrentWhiteboardIndex])
                    {
                        ApplyHistoryToCanvas(item);
                    }
                }
            }
            catch { }
        }

        private void BtnWhiteBoardSwitchPrevious_Click(object sender, EventArgs e)
        {
            if (CurrentWhiteboardIndex <= 1) return;

            SaveStrokes();

            ClearStrokes(true);
            CurrentWhiteboardIndex--;

            RestoreStrokes();

            UpdateIndexInfoDisplay();
        }

        private void BtnWhiteBoardSwitchNext_Click(object sender, EventArgs e)
        {
            if (CurrentWhiteboardIndex >= WhiteboardTotalCount)
            {
                BtnWhiteBoardAdd_Click(sender, e);
                return;
            }
            if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
            {
                SaveScreenShot(true);
                if (Settings.Automation.IsAutoSaveStrokesAtScreenshot) SaveInkCanvasStrokes(false);
            }
            SaveStrokes();


            ClearStrokes(true);
            CurrentWhiteboardIndex++;

            RestoreStrokes();

            UpdateIndexInfoDisplay();
        }

        private void BtnWhiteBoardAdd_Click(object sender, EventArgs e)
        {
            if (WhiteboardTotalCount >= 99) return;
            if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
            {
                SaveScreenShot(true);
                if (Settings.Automation.IsAutoSaveStrokesAtScreenshot) SaveInkCanvasStrokes(false);
            }
            SaveStrokes();
            ClearStrokes(true);

            WhiteboardTotalCount++;
            CurrentWhiteboardIndex++;

            if (CurrentWhiteboardIndex != WhiteboardTotalCount)
            {
                for (int i = WhiteboardTotalCount; i > CurrentWhiteboardIndex; i--)
                {
                    TimeMachineHistories[i] = TimeMachineHistories[i - 1];
                }
            }

            UpdateIndexInfoDisplay();

            if (WhiteboardTotalCount >= 99) BtnWhiteBoardAdd.IsEnabled = false;
        }

        private void BtnWhiteBoardDelete_Click(object sender, RoutedEventArgs e)
        {
            ClearStrokes(true);

            if (CurrentWhiteboardIndex != WhiteboardTotalCount)
            {
                for (int i = CurrentWhiteboardIndex; i <= WhiteboardTotalCount; i++)
                {
                    TimeMachineHistories[i] = TimeMachineHistories[i + 1];
                }
            }
            else
            {
                CurrentWhiteboardIndex--;
            }

            WhiteboardTotalCount--;

            RestoreStrokes();

            UpdateIndexInfoDisplay();

            if (WhiteboardTotalCount < 99) BtnWhiteBoardAdd.IsEnabled = true;
        }

        private void UpdateIndexInfoDisplay()
        {
            TextBlockWhiteBoardIndexInfo.Text = string.Format("{0} / {1}", CurrentWhiteboardIndex, WhiteboardTotalCount);

            if (CurrentWhiteboardIndex == 1)
            {
                BtnWhiteBoardSwitchPrevious.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardSwitchPrevious.IsEnabled = true;
            }

            if (CurrentWhiteboardIndex == WhiteboardTotalCount)
            {
                BtnWhiteBoardSwitchNext.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardSwitchNext.IsEnabled = true;
            }

            if (WhiteboardTotalCount == 1)
            {
                BtnWhiteBoardDelete.IsEnabled = false;
            }
            else
            {
                BtnWhiteBoardDelete.IsEnabled = true;
            }
        }

        private void SetColors()
        {
            if (currentMode != 0 && !Settings.Canvas.UsingWhiteboard)
            {
                if (File.Exists(App.RootPath + "Colors\\Light.ini"))
                {
                    try
                    {
                        string[] lightColors = File.ReadAllLines(App.RootPath + "Colors\\Light.ini");
                        BtnColorRed.Background = new SolidColorBrush(StringToColor(lightColors[0]));
                        BtnColorGreen.Background = new SolidColorBrush(StringToColor(lightColors[1]));
                        BtnColorBlue.Background = new SolidColorBrush(StringToColor(lightColors[2]));
                        BtnColorYellow.Background = new SolidColorBrush(StringToColor(lightColors[3]));
                    }
                    catch (Exception) { ShowNotification("读取亮色画笔颜色配置文件时遇到问题"); }
                }
                else
                {
                    BtnColorRed.Background = new SolidColorBrush(StringToColor("#FFFF3333"));
                    BtnColorGreen.Background = new SolidColorBrush(StringToColor("#FF1ED760"));
                    BtnColorBlue.Background = new SolidColorBrush(StringToColor("#FF239AD6"));
                    BtnColorYellow.Background = new SolidColorBrush(StringToColor("#FFFFC000"));
                }
            }
            else
            {
                if (File.Exists(App.RootPath + "Colors\\Dark.ini"))
                {
                    try
                    {
                        string[] darkColors = File.ReadAllLines(App.RootPath + "Colors\\Dark.ini");
                        BtnColorRed.Background = new SolidColorBrush(StringToColor(darkColors[0]));
                        BtnColorGreen.Background = new SolidColorBrush(StringToColor(darkColors[1]));
                        BtnColorBlue.Background = new SolidColorBrush(StringToColor(darkColors[2]));
                        BtnColorYellow.Background = new SolidColorBrush(StringToColor(darkColors[3]));
                    }
                    catch (Exception) { ShowNotification("读取深色画笔颜色配置文件时遇到问题"); }
                }
                else
                {
                    BtnColorRed.Background = new SolidColorBrush(Colors.Red);
                    BtnColorGreen.Background = new SolidColorBrush(StringToColor("#FF169141"));
                    BtnColorBlue.Background = new SolidColorBrush(StringToColor("#FF239AD6"));
                    BtnColorYellow.Background = new SolidColorBrush(StringToColor("#FFF38B00"));
                }
            }
        }

        #endregion Whiteboard Controls
    }
}
