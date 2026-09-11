using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：弹层（悬浮栏弹出的各种面板/菜单）的统一「点外面即收起」模型。
    ///
    /// 【为什么单独一个文件】原实现散在两处、规则各写一遍，行为还不一致：
    ///   · MW_PenSettings.cs 的 Window_PreviewMouseDown —— 笔/选择/橡皮 三段手写 if
    ///   · MW_Capture.cs 的 ImageLayer_MenuCloseOnOutsideClick —— 截图菜单，靠代码订阅
    ///   · 「更多」面板压根没人管它 → 点外面不会收（用户反馈的"有的一直存在"）
    /// 2026-09-11 收口成本文件下方这一张表：谁要"点外面即收起"，加一行即可。
    ///
    /// 【全项目 7 个弹层 —— 本表只登记"点外收起"这部分】
    ///   1 笔设置     PenSettingsPanel    ← 开关图标 PenIconHost          ✅ 登记
    ///   2 选择方式   SelectionModePanel  ← 开关图标 GridSelectTool       ✅ 登记
    ///   3 橡皮设置   EraserSettingsPanel ← 开关图标 EraserContainer      ✅ 登记
    ///   4 更多面板   BorderTools         ← 开关图标 GridToolsEntry       ✅ 登记（2026-09-11 新增）
    ///   5 截图菜单   BorderImageMenu     ← 开关图标 GridImageMenuEntry   ✅ 登记（原独立处理器，已并入本表）
    ///   6 图形绘制   BorderDrawShape     — 不进表：有图钉、可拖动、可钉住的常驻小窗，
    ///                                      点外面就关会打断连续画图（有意如此，不是疏漏）
    ///   7 数学面板   原生 COM 窗口        — 不进表：不是 WPF 元素，收不到 WPF 鼠标事件，
    ///                                      开关见 MW_MathGraph.cs
    ///
    /// 【注意区分另一个入口】"主动关闭（不依赖点击）"的统一入口是 MW_FloatBar.cs 的
    /// ClosePopupLayers(PopupScope)（切工具 / 工具条收起时用）。两者登记的面板范围有意不同，改一边时请看另一边。
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>一个「点面板外面就收起」的弹层：面板 + 它的开关图标 + 怎么关它</summary>
        private sealed class DismissibleLayer
        {
            /// <summary>要收起的面板（用它判断是否可见，也用来判断"按下点是否在面板内部"）</summary>
            public readonly FrameworkElement Panel;

            /// <summary>该面板的开关图标：点它不收起，交给图标自己的 MouseUp 做 toggle</summary>
            public readonly DependencyObject Anchor;

            /// <summary>收起动作（统一调各面板自己的 Close* 方法，便于将来往那里加清理逻辑）</summary>
            public readonly Action Close;

            public DismissibleLayer(FrameworkElement panel, DependencyObject anchor, Action close)
            {
                Panel = panel;
                Anchor = anchor;
                Close = close;
            }
        }

        // 只在首次点击时构建一次 —— PreviewMouseDown 是每次落笔都走的热路径，不能每次 new 数组
        private DismissibleLayer[] _dismissibleLayers;

        private DismissibleLayer[] GetDismissibleLayers()
        {
            if (_dismissibleLayers == null)
            {
                _dismissibleLayers = new[]
                {
                    new DismissibleLayer(PenSettingsPanel,    PenIconHost,        ClosePenSettingsPanel),
                    new DismissibleLayer(SelectionModePanel,  GridSelectTool,     CloseSelectionModePanel),
                    new DismissibleLayer(EraserSettingsPanel, EraserContainer,    CloseEraserSettingsPanel),
                    // BorderTools 是 Border 不是 UserControl，没有专门的 Close 方法（也没有额外清理）
                    new DismissibleLayer(BorderTools,         GridToolsEntry,     () => BorderTools.Visibility = Visibility.Collapsed),
                    new DismissibleLayer(BorderImageMenu,     GridImageMenuEntry, () => ImageLayer_MenuSetVisible(false)),
                };
            }
            return _dismissibleLayers;
        }

        /// <summary>
        /// 点击面板外任意处收起面板（Window.PreviewMouseDown，隧道事件最先收到）。
        /// 规则统一：表里每个"当前可见"的面板，只要按下点既不在它内部、也不在它的开关图标上，就收起它。
        /// 顺带得到"同组互斥"：点另一个图标时，先前那个面板先在这里被收掉，再由图标自己的 MouseUp toggle 打开。
        /// 点图标不处理，是因为图标自己的 MouseUp 要做 toggle，这里抢先收起会变成"点图标关不掉"（先关再开）。
        /// 只关面板、不设 e.Handled，不影响书写/擦除/双指手势。
        /// </summary>
        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            if (!(e.OriginalSource is DependencyObject dep)) return;

            // 记录按下位置：快速截图后的白板单击检测用（Window_PreviewMouseUp 比较位移区分单击/书写）
            _pasteClickDownPos = e.GetPosition(Main_Grid);

            var layers = GetDismissibleLayers();
            for (int i = 0; i < layers.Length; i++)
            {
                var layer = layers[i];
                if (layer.Panel.Visibility != Visibility.Visible) continue;
                if (IsVisualDescendantOf(dep, layer.Panel)) continue; // 点在面板内：不收（面板内可连续调整）
                if (IsVisualDescendantOf(dep, layer.Anchor)) continue; // 点在开关图标：交给它自己 toggle
                layer.Close();
            }
        }

        /// <summary>判断 visual 是否是 ancestor 的后代（自身或视觉树子孙）</summary>
        private bool IsVisualDescendantOf(DependencyObject visual, DependencyObject ancestor)
        {
            while (visual != null)
            {
                if (ReferenceEquals(visual, ancestor)) return true;
                // 沿视觉树向上（跨 ContentElement 用 VisualTreeHelper.GetParent 也兼容 null 返回自身场景）
                var parent = VisualTreeHelper.GetParent(visual);
                if (parent == null && visual is FrameworkContentElement fce)
                    parent = fce.Parent as DependencyObject;
                visual = parent;
            }
            return false;
        }
    }
}
