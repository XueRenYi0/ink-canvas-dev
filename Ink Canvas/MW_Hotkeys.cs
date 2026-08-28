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
    /// <summary>MainWindow 分部类：快捷键（自 MainWindow.xaml.cs 拆分，逻辑未改动）</summary>
    public partial class MainWindow
    {
        #region Hotkeys

        private void Window_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (StackPanelPPTControls.Visibility != Visibility.Visible || currentMode != 0) return;
            if (e.Delta >= 120)
            {
                BtnPPTSlidesUp_Click(BtnPPTSlidesUp, null);
            }
            else if (e.Delta <= -120)
            {
                BtnPPTSlidesDown_Click(BtnPPTSlidesDown, null);
            }
        }

        private void Main_Grid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (StackPanelPPTControls.Visibility != Visibility.Visible || currentMode != 0) return;

            if (e.Key == Key.Down || e.Key == Key.PageDown || e.Key == Key.Right || e.Key == Key.N || e.Key == Key.Space)
            {
                BtnPPTSlidesDown_Click(BtnPPTSlidesDown, null);
            }
            if (e.Key == Key.Up || e.Key == Key.PageUp || e.Key == Key.Left || e.Key == Key.P)
            {
                BtnPPTSlidesUp_Click(BtnPPTSlidesUp, null);
            }
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                KeyExit(null, null);
            }
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void back_HotKey(object sender, ExecutedRoutedEventArgs e)
        {
            try
            {
                inkCanvas.Strokes.Remove(inkCanvas.Strokes[inkCanvas.Strokes.Count - 1]);
            }
            catch { }
        }

        private void KeyExit(object sender, ExecutedRoutedEventArgs e)
        {
            BtnPPTSlideShowEnd_Click(BtnPPTSlideShowEnd, null);
        }

        private void KeyChangeToDrawTool(object sender, ExecutedRoutedEventArgs e)
        {
            if (inkCanvas.Visibility == Visibility.Collapsed)
            {
                BtnHideInkCanvas_Click(sender, e);
            }
        }

        private void KeyChangeToSelect(object sender, ExecutedRoutedEventArgs e)
        {
            if (inkCanvas.Visibility == Visibility.Visible)
            {
                BtnHideInkCanvas_Click(sender, e);
            }
        }

        private void KeyChangeToEraser(object sender, ExecutedRoutedEventArgs e)
        {
            if (ImageEraserMask.Visibility == Visibility.Visible)
            {
                BorderPenColorRed_MouseUp(null, null);
            }
            else
            {
                BtnErase_Click(sender, e);
            }
        }

        private void KeyCapture(object sender, ExecutedRoutedEventArgs e)
        {
            BtnScreenshot_Click(sender, e);
        }

        private void KeyDrawLine(object sender, ExecutedRoutedEventArgs e)
        {
            SetColorByIndex();
            BtnDrawLine_Click(lastMouseDownSender, e);
        }

        private void KeyHide(object sender, ExecutedRoutedEventArgs e)
        {
            SymbolIconEmoji_MouseUp(sender, null);
        }

        #endregion Hotkeys
    }
}
