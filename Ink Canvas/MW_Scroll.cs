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
    /// </summary>
    public partial class MainWindow
    {
        #region Note Scroll（笔记上下滚动）

        /// <summary>当前累计滚动量：&gt;0 表示已向下滚动（历史墨迹已向上平移）。仅作记账，不参与坐标计算。</summary>
        double _noteScrollOffsetY = 0;

        /// <summary>一次滚动的步长：工作区高度的 70%（约一屏板书，类似翻页）</summary>
        double NoteScrollStep => SystemParameters.WorkArea.Height * 0.7;

        /// <summary>
        /// 初始化滚动功能（由 MainWindow 构造函数调用）。
        /// 画布墨迹清空（清屏 / 换页 / 模式切换的中转过程）时滚动位置归零。
        /// </summary>
        private void InitNoteScroll()
        {
            inkCanvas.Strokes.StrokesChanged += NoteScroll_StrokesChanged;
        }

        private void NoteScroll_StrokesChanged(object sender, StrokeCollectionChangedEventArgs e)
        {
            // 换页/清屏流程为 ClearStrokes -> RestoreStrokes：
            // 墨迹数归零的瞬间重置滚动记账，新页从顶部开始。
            // 用户擦除全部墨迹时同理（无内容则滚动位置无意义）。
            if (inkCanvas.Strokes.Count == 0) _noteScrollOffsetY = 0;
        }

        private void BtnScrollUp_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(-NoteScrollStep);
        }

        private void BtnScrollDown_MouseUp(object sender, MouseButtonEventArgs e)
        {
            ScrollNote(NoteScrollStep);
        }

        /// <summary>
        /// 滚动画布。delta &gt; 0 向下滚动（历史墨迹上移，露出下方空白区域继续书写）。
        /// 顶部夹取为 0（不允许滚出第一屏上方）；向下不设上限。
        /// 滚动本身不进入撤销栈（仅墨迹操作可撤销）。
        /// </summary>
        private void ScrollNote(double delta)
        {
            // 防御：仅白板/黑板模式（按钮可见性已由 XAML 绑定控制，此处双保险）
            if (GridBackgroundCover.Visibility != Visibility.Visible) return;

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
