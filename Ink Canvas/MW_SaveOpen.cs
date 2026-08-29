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
    /// <summary>MainWindow 分部类：墨迹保存与打开（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Save & Open

        private void SymbolIconSaveStrokes_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender || inkCanvas.Visibility != Visibility.Visible) return;

            BorderTools.Visibility = Visibility.Collapsed;

            GridNotifications.Visibility = Visibility.Collapsed;

            SaveInkCanvasStrokes();
        }

        private void SaveInkCanvasStrokes(bool newNotice = true)
        {
            try
            {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved");
                }

                FileStream fs = new FileStream(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                    @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk", FileMode.Create); //Ink Canvas STroKes
                inkCanvas.Strokes.Save(fs);

                if (newNotice)
                {
                    ShowNotification("墨迹成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                        @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk");
                }
                else
                {
                    AppendNotification("墨迹成功保存至 " + Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) +
                        @"\Ink Canvas Strokes\User Saved\" + DateTime.Now.ToString("u").Replace(':', '-') + ".icstk");
                }
            }
            catch (Exception ex)
            {
                ShowNotification($"墨迹保存失败：{ex.Message}");
            }
        }

        private void SymbolIconPin_MouseUp(object sender, MouseButtonEventArgs e)
        {
            _lockSmith = !_lockSmith;
            UpdateGestureLockIcon();
        }

        /// <summary>双指手势锁图标状态：锁定=显示斜杠+手指变淡；解锁=隐藏斜杠+手指正常</summary>
        private void UpdateGestureLockIcon()
        {
            if (_lockSmith)
            {
                PathGestureSlash.Visibility = Visibility.Visible;
                PathGestureFingers.Opacity = 0.45;
            }
            else
            {
                PathGestureSlash.Visibility = Visibility.Collapsed;
                PathGestureFingers.Opacity = 1.0;
            }
        }

        private void SymbolIconOpenStrokes_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (lastBorderMouseDownObject != sender) return;
            BorderTools.Visibility = Visibility.Collapsed;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            string defaultFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\Ink Canvas Strokes\User Saved";
            if (Directory.Exists(defaultFolderPath))
            {
                openFileDialog.InitialDirectory = defaultFolderPath;
            }
            openFileDialog.Title = "打开墨迹文件";
            openFileDialog.Filter = "Ink Canvas Strokes File (*.icstk)|*.icstk";
            if (openFileDialog.ShowDialog() == true)
            {
                LogHelper.WriteLogToFile(string.Format("Strokes Insert: Name: {0}", openFileDialog.FileName), LogHelper.LogType.Event);
                try
                {
                    var fileStreamHasNoStroke = false;
                    using (var fs = new FileStream(openFileDialog.FileName, FileMode.Open, FileAccess.Read))
                    {
                        var strokes = new StrokeCollection(fs);
                        fileStreamHasNoStroke = strokes.Count == 0;
                        if (!fileStreamHasNoStroke)
                        {
                            ClearStrokes(true);
                            timeMachine.ClearStrokeHistory();
                            inkCanvas.Strokes.Add(strokes);
                            LogHelper.NewLog(string.Format("Strokes Insert: Strokes Count: {0}", inkCanvas.Strokes.Count.ToString()));
                        }
                    }
                    if (fileStreamHasNoStroke)
                    {
                        using (var ms = new MemoryStream(File.ReadAllBytes(openFileDialog.FileName)))
                        {
                            ms.Seek(0, SeekOrigin.Begin);
                            var strokes = new StrokeCollection(ms);
                            ClearStrokes(true);
                            timeMachine.ClearStrokeHistory();
                            inkCanvas.Strokes.Add(strokes);
                            LogHelper.NewLog(string.Format("Strokes Insert (2): Strokes Count: {0}", strokes.Count.ToString()));
                        }
                    }

                    if (inkCanvas.Visibility != Visibility.Visible)
                    {
                        SymbolIconCursor_Click(sender, null);
                    }
                }
                catch
                {
                    ShowNotification("墨迹打开失败");
                }
            }
        }



        #endregion

        #region Multi-finger Inking


        #endregion
    }
}
