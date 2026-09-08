using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// 选择方式面板：矩形框选 / 自由选择（套索）/ 全选。
    /// 纯 UI 组件——选择动作通过事件抛给 MainWindow 订阅处理，
    /// 自身不碰 inkCanvas / Settings，保持与主窗口解耦（一面板一控件规范）。
    /// </summary>
    public partial class SelectionModePanel : UserControl
    {
        public SelectionModePanel()
        {
            InitializeComponent();
        }

        // ===== 对外事件（MainWindow 订阅） =====

        /// <summary>选择方式被点击：0=矩形框选 1=自由选择（套索）</summary>
        public event Action<int> SelectionModeSelected;

        /// <summary>全选被点击（动作项：立即全选当前页全部墨迹）</summary>
        public event Action SelectAllRequested;

        // ===== 状态更新入口（MainWindow 调用，用于同步高亮） =====

        /// <summary>刷新选中高亮（0=矩形 1=自由；全选是动作项无选中态）</summary>
        public void UpdateModeHighlight(int mode)
        {
            SetGridHighlight(ModeRectHighlight, mode == 0);
            SetGridHighlight(ModeLassoHighlight, mode == 1);
        }

        // ===== 内部实现 =====

        /// <summary>选中态视觉：淡蓝底 + 蓝边，未选中透明（与笔面板同规范）</summary>
        private void SetGridHighlight(Border border, bool selected)
        {
            border.Background = selected
                ? (Brush)TryFindResource("SelPanelSoftHighlight")
                : Brushes.Transparent;
            border.BorderBrush = selected
                ? (Brush)TryFindResource("HighlightBrush")
                : Brushes.Transparent;
        }

        // 两个方式按钮共用：报索引进事件
        private void ModeButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (sender == BorderModeRect) SelectionModeSelected?.Invoke(0);
            else if (sender == BorderModeLasso) SelectionModeSelected?.Invoke(1);
        }

        // 全选按钮：报"执行全选"事件（MainWindow 全选后关面板）
        private void SelectAllButton_MouseUp(object sender, MouseButtonEventArgs e)
        {
            SelectAllRequested?.Invoke();
        }
    }
}
