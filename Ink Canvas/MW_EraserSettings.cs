using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：橡皮设置面板（EraserSettingsPanel）的弹出/定位/关闭，
    /// 擦除方式（面积擦/笔画擦）与大小的实现，以及滑动清屏。
    ///
    /// 交互模型（与笔面板/选择面板同款）：
    /// - 点橡皮图标 → 进入上次的擦除方式 + 弹出面板（方式/大小/滑动清屏三区）
    /// - 面板内点选不关面板（可连续调整）
    /// - 点面板外任意处 → 面板消失（Window_PreviewMouseDown 统一判定）
    ///
    /// 改造要点：
    /// - 原交互"点图标在面积擦/笔画擦间 toggle"已废弃（隐晦、不可见）——
    ///   方式切换只能通过面板，点图标固定进入当前方式
    /// - 原浮动条垃圾桶清屏入口已移入面板（滑动清屏，ClassIn 式防误触）
    /// </summary>
    public partial class MainWindow
    {
        #region 面板初始化与事件接线

        /// <summary>初始化橡皮设置面板：订阅面板事件（MainWindow 构造后调用一次）</summary>
        private void InitEraserSettingsPanel()
        {
            // 面板事件 → 主窗口动作（面板自身不碰 inkCanvas/Settings，保持解耦）
            EraserSettingsPanel.EraserModeSelected += m => SetEraserMode(m == 0);
            EraserSettingsPanel.EraserSizeSelected += s => SetEraserSize(s);
            EraserSettingsPanel.ClearRequested += ExecuteSlideClearScreen;
        }

        /// <summary>设置加载完成后同步面板状态（MW_DefinitionsLoading 的 Canvas 分支末尾调用）</summary>
        private void ApplyLoadedEraserSettings()
        {
            RefreshEraserSettingsPanelState();
        }

        /// <summary>
        /// 切换擦除方式（面板区一点击触发）。
        /// 当场生效：更新橡皮形状与编辑模式、同步浮动条双层图标；EraserType 设置不改
        /// （它是"进入橡皮模式时的初值"语义：0=记住上次 1=强制面积 2=强制笔画，仍归设置页管）。
        /// </summary>
        private void SetEraserMode(bool pointEraser)
        {
            forcePointEraser = pointEraser;
            inkCanvas.EraserShape = CreateEraserShape(forcePointEraser);

            // 当前正处于橡皮模式时立即切编辑模式，否则只在下次进入时生效
            if (inkCanvas.EditingMode == InkCanvasEditingMode.EraseByPoint ||
                inkCanvas.EditingMode == InkCanvasEditingMode.EraseByStroke)
            {
                inkCanvas.EditingMode =
                    forcePointEraser ? InkCanvasEditingMode.EraseByPoint : InkCanvasEditingMode.EraseByStroke;
            }

            UpdateEraserIcon();
            RefreshEraserSettingsPanelState();
        }

        /// <summary>切换橡皮大小（面板区二点击触发）：写设置 + 面积擦时立即重建橡皮形状。
        /// 【WPF 坑】EditingMode 处于擦除模式时直接改 EraserShape 不生效（光标圈不变）——
        /// 必须先切到 None 再设形状再切回擦除模式，形状和光标才会真正刷新。</summary>
        private void SetEraserSize(int index)
        {
            if (index < 0 || index > 4) return;
            Settings.Canvas.EraserSize = index;
            if (forcePointEraser)
            {
                var mode = inkCanvas.EditingMode;
                if (mode == InkCanvasEditingMode.EraseByPoint)
                    inkCanvas.EditingMode = InkCanvasEditingMode.None; // 先离开擦除模式
                inkCanvas.EraserShape = CreateEraserShape(true);
                if (mode == InkCanvasEditingMode.EraseByPoint)
                    inkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint; // 切回（形状已更新）
            }
            RefreshEraserSettingsPanelState();
        }

        #endregion

        #region 面板弹出 / 定位 / 点击外部收起

        /// <summary>切换橡皮设置面板显隐（橡皮图标单击入口，ImageEraser_MouseUp 调用）</summary>
        private void ToggleEraserSettingsPanel()
        {
            if (EraserSettingsPanel.Visibility == Visibility.Visible)
            {
                EraserSettingsPanel.Visibility = Visibility.Collapsed;
                return;
            }

            RefreshEraserSettingsPanelState();
            EraserSettingsPanel.ResetSlideThumb(); // 每次弹出滑块归位（上次未滑完的残留清零）
            EraserSettingsPanel.Visibility = Visibility.Visible;
            PositionEraserSettingsPanel();
        }

        /// <summary>关闭橡皮设置面板（HideSubPanels / 点击外部等场景调用）</summary>
        private void CloseEraserSettingsPanel()
        {
            EraserSettingsPanel.Visibility = Visibility.Collapsed;
        }

        /// <summary>弹出前刷新面板高亮状态（方式/大小可能被设置页或热键改过）</summary>
        private void RefreshEraserSettingsPanelState()
        {
            EraserSettingsPanel.UpdateModeHighlight(forcePointEraser ? 0 : 1);
            EraserSettingsPanel.UpdateSizeHighlight(Settings.Canvas.EraserSize);
            // 笔画擦无大小概念：整区灰显禁用（不适用功能灰显规范）
            EraserSettingsPanel.SetSizeSectionEnabled(forcePointEraser);
        }

        /// <summary>
        /// 面板定位：以橡皮图标为锚，图标上方空间够就弹在上方（顶不够时弹下方），
        /// 水平居中对齐橡皮图标并夹在屏幕内（与笔面板同款算法，见 PositionPenSettingsPanel 注释）。
        /// </summary>
        private void PositionEraserSettingsPanel()
        {
            Main_Grid.UpdateLayout();
            double panelW = EraserSettingsPanel.ActualWidth;
            double panelH = EraserSettingsPanel.ActualHeight;
            if (panelW < 1 || panelH < 1) return; // 布局未就绪，放弃定位（保持上次位置）

            var transform = EraserContainer.TransformToAncestor(Main_Grid);
            Point iconTopLeft = transform.Transform(new Point(0, 0));
            Point iconBottomRight = transform.Transform(new Point(EraserContainer.ActualWidth, EraserContainer.ActualHeight));
            double iconCenterX = (iconTopLeft.X + iconBottomRight.X) / 2;
            double iconTopY = iconTopLeft.Y;
            double iconBottomY = iconBottomRight.Y;

            const double Gap = 8;
            double gridW = Main_Grid.ActualWidth;
            double gridH = Main_Grid.ActualHeight;

            double y = iconTopY - panelH - Gap;
            if (y < Gap) y = iconBottomY + Gap;
            if (y + panelH > gridH - Gap) y = Math.Max(Gap, gridH - panelH - Gap);

            double x = iconCenterX - panelW / 2;
            if (x < Gap) x = Gap;
            if (x + panelW > gridW - Gap) x = Math.Max(Gap, gridW - panelW - Gap);

            EraserSettingsPanel.Margin = new Thickness(x, y, 0, 0);
        }

        #endregion

        #region 滑动清屏（原 SymbolIconDelete_MouseUp 逻辑迁移）

        /// <summary>
        /// 滑动清屏执行体（面板滑块拖到底触发）：
        /// 1. 有墨迹或当前页有图片 → 可选自动截图 + 清屏，随后切回画笔
        ///    （清完屏最常见的下一步是写新内容；原"保持橡皮不变"会让第一笔变成擦除）
        /// 2. 空白页 → 不做任何事（原垃圾桶会把空屏滑动变成"隐藏画布回鼠标模式"，
        ///    滑动清屏语义明确指向清墨迹，不该有意外的模式切换——用户定稿废除）
        ///
        /// 注：原"有选中墨迹时改为删除选中"分支已删——弹面板前必先点橡皮图标，
        /// 而 InkCanvas 切到擦除模式时选区已被框架清掉，该分支实际永远走不到（死代码）；
        /// 删除选中的需求由选中操作条的删除按钮直接承担。
        /// </summary>
        private void ExecuteSlideClearScreen()
        {
            CloseEraserSettingsPanel(); // 清屏触发后收起面板

            // 当前视图层键：白板模式=当前白板页；注释模式（第 0 页）=批注层 -1（两层的图片都要能清屏）
            int viewKey = currentMode != 0 ? CurrentWhiteboardIndex : -1;
            if (inkCanvas.Strokes.Count > 0
                || (_pageImages.TryGetValue(viewKey, out var pageImgs) && pageImgs.Count > 0))
            {
                // 有墨迹，或当前视图层有图片（纯图片页也要能清屏）
                if (Settings.Automation.IsAutoSaveStrokesAtClear && inkCanvas.Strokes.Count > Settings.Automation.MinimumAutomationStrokeNumber)
                {
                    if (BtnPPTSlideShowEnd.Visibility == Visibility.Visible)
                        SaveScreenShot(true, $"{pptName}/{previousSlideID}_{DateTime.Now:HH-mm-ss}");
                    else
                        SaveScreenShot(true);
                }
                BtnClear_Click(BtnClear, null);

                // 清屏后切回画笔：ColorSwitchCheck 无条件恢复 Ink 模式 + 橡皮图标可见性 +
                // 熄灭图形高亮（与笔图标点选同一条路径，状态全部同步）
                ColorSwitchCheck();
            }
            // 空白页 → 什么都不做（见方法注释）
        }

        #endregion
    }
}
