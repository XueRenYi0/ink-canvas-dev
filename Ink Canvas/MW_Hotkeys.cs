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
                // 粘贴模式中：Esc = 取消粘贴（恢复原工具和光标），不触发全局退出
                if (isPasteMode)
                {
                    ExitPasteMode();
                    return;
                }
                KeyExit(null, null);
            }
            // Delete 键删除选中的图片（选择工具下点选图片后按 Delete；墨迹删除走操作条按钮）
            if (e.Key == Key.Delete)
            {
                ImageLayer_DeleteSelectedImages();
            }

            // ---- 剪贴板贯通快捷键（设置页文本框获得焦点时由文本框处理，不会进到这里） ----
            // Ctrl+C：选中墨迹 → 渲染成图片存剪贴板（无选中时不占用剪贴板——别覆盖用户刚截的图）
            if (e.KeyboardDevice.Modifiers == ModifierKeys.Control && e.Key == Key.C)
            {
                if (inkCanvas.GetSelectedStrokes().Count > 0)
                {
                    CopySelectedStrokesToClipboard();
                    ShowNotification("选中墨迹已复制到剪贴板，Ctrl+V 可粘贴");
                }
                return;
            }
            // Ctrl+V：剪贴板图片 → 插入白板当前页（级联错位，连续粘贴不叠死）
            if (e.KeyboardDevice.Modifiers == ModifierKeys.Control && e.Key == Key.V)
            {
                PasteClipboardImageToCanvas(null);
                return;
            }
        }

        private void CommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
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
            // 幂等：切到「选择墨迹」工具 —— 与点悬浮条的选择图标走同一条路径（BtnSelect_Click），
            // 而不是旧的 BtnHideInkCanvas_Click。
            // 旧实现名义上叫 ChangeToSelect，实际是"隐藏画布"（上游遗留），
            // 按 Ctrl+M 会把画布藏起来，看起来就是"选择快捷键没反应"。
            BtnSelect_Click(BtnSelect, null);
        }

        private void KeyChangeToEraser(object sender, ExecutedRoutedEventArgs e)
        {
            // 幂等：每次都进入橡皮（BtnErase_Click 内部已是"固定进入当前擦除方式，不 toggle"）。
            // 旧实现是「橡皮 ⇄ 画笔」来回切（依据旧界面控件 ImageEraserMask 的可见性判断），
            // 表现为按一下切橡皮、再按又切回画笔、再按才又切橡皮 —— 用户以为快捷键坏了。
            BtnErase_Click(BtnErase, null);
        }

        #endregion Hotkeys
    }
}
