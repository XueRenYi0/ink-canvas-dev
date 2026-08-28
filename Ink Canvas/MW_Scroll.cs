using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Threading;
using Point = System.Windows.Point;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：笔记上下滚动（虚拟无限画布）。
    /// 设计见 DEVELOPMENT.md。核心机制：
    /// 采用"坐标物化"方案——滚动时把平移矩阵直接应用到 Stroke 点数据
    /// （inkCanvas.Strokes.Transform(matrix, true)），而不是用渲染变换。
    /// 这样橡皮命中检测、墨迹保存、撤销/重做（TimeMachine 快照与画布
    /// 共享同一批 Stroke 对象引用）都天然保持坐标一致。
    /// 新墨迹直接写在当前可视区域（旧墨迹已物化移走，屏幕自然空白）。
    /// 启用范围：画布可见（正在书写）时——白板/黑板与屏幕/PPT 注释模式；
    /// 画布收起时滚轮穿透给底层应用（PPT 翻页等），互不干扰。
    /// 后续统一快捷键系统可直接注册 ScrollNote(delta)。
    /// </summary>
    public partial class MainWindow
    {
        #region Note Scroll（笔记上下滚动）

        /// <summary>当前累计滚动量：&gt;0 表示已向下滚动（历史墨迹已向上平移）。仅作记账，不参与坐标计算。</summary>
        double _noteScrollOffsetY = 0;

        /// <summary>按钮单次点击的步长：工作区高度的 35%（小半屏）</summary>
        double NoteScrollButtonStep => SystemParameters.WorkArea.Height * 0.35;

        /// <summary>滚轮每格（Delta=120）的步长：工作区高度的 10%，细粒度顺滑</summary>
        double NoteScrollWheelStep => SystemParameters.WorkArea.Height * 0.10;

        /// <summary>按钮可见性同步定时器（模式切换点分散，统一低频同步，避免逐点挂钩）</summary>
        DispatcherTimer _noteScrollVisibilityTimer;

        /// <summary>
        /// 当前是否允许笔记滚动：
        /// 画布可见（正在书写）即可滚——白板/黑板模式，或屏幕/PPT 注释模式且画布已激活。
        /// 画布收起时滚轮穿透给底层应用（如 PPT 放映翻页），互不干扰。
        /// 后续统一快捷键系统可直接注册 ScrollNote()。
        /// </summary>
        private bool IsNoteScrollActive
        {
            get
            {
                if (inkCanvas.Visibility != Visibility.Visible) return false; // 画布收起：滚轮还给系统
                return currentMode == 0 || currentMode == 1;                  // 屏幕注释 / 白板黑板
            }
        }

        /// <summary>
        /// 初始化滚动功能（由 MainWindow 构造函数调用）：
        /// 1. 墨迹清空（清屏/换页/模式切换中转）时滚动记账归零；
        /// 2. 挂滚轮事件；
        /// 3. 启动按钮可见性同步定时器。
        /// </summary>
        private void InitNoteScroll()
        {
            inkCanvas.Strokes.StrokesChanged += NoteScroll_StrokesChanged;
            inkCanvas.PreviewMouseWheel += InkCanvas_PreviewMouseWheel;

            _noteScrollVisibilityTimer = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(400) };
            _noteScrollVisibilityTimer.Tick += (s, e) =>
            {
                GridNoteScrollControls.Visibility = IsNoteScrollActive ? Visibility.Visible : Visibility.Collapsed;
            };
            _noteScrollVisibilityTimer.Start();
        }

        private void NoteScroll_StrokesChanged(object sender, StrokeCollectionChangedEventArgs e)
        {
            // 换页/清屏流程为 ClearStrokes -> RestoreStrokes：
            // 墨迹数归零的瞬间重置滚动记账，新页从顶部开始。
            // 用户擦除全部墨迹时同理（无内容则滚动位置无意义）。
            if (inkCanvas.Strokes.Count == 0) _noteScrollOffsetY = 0;
        }

        private void InkCanvas_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (!IsNoteScrollActive) return;

            // 滚轮向上（正 Delta）= 回看上方历史；向下 = 露出下方空白
            var notches = e.Delta / 120.0;
            if (Math.Abs(notches) < 0.01) return;
            ScrollNote(-notches * NoteScrollWheelStep);
            e.Handled = true;
        }

        private void BtnScrollUp_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(-NoteScrollButtonStep);
        }

        private void BtnScrollDown_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(NoteScrollButtonStep);
        }

        /// <summary>
        /// 滚动画布。delta &gt; 0 向下滚动（历史墨迹上移，露出下方空白区域继续书写）。
        /// 顶部夹取为 0（不允许滚出第一屏上方）；向下不设上限。
        /// 滚动本身不进入撤销栈（仅墨迹操作可撤销）。
        /// </summary>
        private void ScrollNote(double delta)
        {
            // 防御：仅白板/黑板与屏幕注释模式（与按钮可见性逻辑双保险）
            if (!IsNoteScrollActive) return;

            var target = Math.Max(0, _noteScrollOffsetY + delta);
            var actual = target - _noteScrollOffsetY;
            if (Math.Abs(actual) < 1) return;

            // 坐标物化：矩阵直接应用到 Stroke 点数据（applyOnStrokes: true）
            // 向下滚动 actual>0，墨迹整体上移（Y 减小）
            var matrix = new Matrix(1, 0, 0, 1, 0, -actual);
            inkCanvas.Strokes.Transform(matrix, true);
            _noteScrollOffsetY = target;
        }

        #endregion
    }
}
