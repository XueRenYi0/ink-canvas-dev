using System;
using System.Windows;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using Ink_Canvas.Helpers;

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
        /// （收起时"顺手画出的那一个点"由下方 DismissPressDown/Up 状态机负责丢弃，见其注释。）
        /// </summary>
        private void Window_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Point pos = e.GetPosition(Main_Grid);
            var dep = e.OriginalSource as DependencyObject;

            // 记录按下位置：快速截图后的白板单击检测用（Window_PreviewMouseUp 比较位移区分单击/书写）
            if (dep != null) _pasteClickDownPos = pos;

            bool closedAny = CloseDismissibleLayersOutside(dep);
            DismissPressDown(pos, closedAny);
        }

        /// <summary>触笔通道的"按下"。与鼠标通道共用同一套状态机，否则笔/鼠标/手指行为会不一致。</summary>
        private void Window_PreviewStylusDown(object sender, StylusDownEventArgs e)
        {
            Point pos = e.GetPosition(Main_Grid);
            var dep = e.OriginalSource as DependencyObject;
            if (dep != null) _pasteClickDownPos = pos;

            bool closedAny = CloseDismissibleLayersOutside(dep);
            DismissPressDown(pos, closedAny);
        }

        /// <summary>触笔通道的"抬起"：交给同一条判定（重复调用安全，状态机会自己短路）。</summary>
        private void Window_PreviewStylusUp(object sender, StylusEventArgs e)
        {
            DismissPressUp(e.GetPosition(Main_Grid));
        }

        /// <summary>把"按下点不在其内部、也不在其开关图标上"的可见面板全部收起；返回是否真的收起了至少一个。</summary>
        private bool CloseDismissibleLayersOutside(DependencyObject dep)
        {
            if (dep == null) return false;
            bool closed = false;
            var layers = GetDismissibleLayers();
            for (int i = 0; i < layers.Length; i++)
            {
                var layer = layers[i];
                if (layer.Panel.Visibility != Visibility.Visible) continue;
                if (IsVisualDescendantOf(dep, layer.Panel)) continue; // 点在面板内：不收（面板内可连续调整）
                if (IsVisualDescendantOf(dep, layer.Anchor)) continue; // 点在开关图标：交给它自己 toggle
                layer.Close();
                closed = true;
            }
            // 诊断（低频、只在真的收了面板时写一行）：这个功能的判定链较长，
            // 有这行才能区分"没收到面板"（机制没启动）和"收了但点被留下"（判定出错）。
            if (closed)
                LogHelper.WriteLogToFile("[Popup] 点外收起：按下时收起了可见面板 → 本次抬起将按位移判定单击/书写",
                    LogHelper.LogType.Event);
            return closed;
        }

        // ==================== "点外收起"不要顺手画出一个点 ====================
        //
        // 现象：面板开着时点画布空白处收面板 —— 面板收起来了，但画布上多了一个点。
        // 根因：收起面板是"不拦事件"的（上面那句注释：不影响书写/擦除/双指手势），
        //       于是同一次点击继续落到 InkCanvas，抬起时提交了一根"单点笔迹"。
        //
        // 为什么不能干脆在按下时 e.Handled = true 一拦了事：
        //   ① 会把"书写"手势的按下一起吞掉 → 面板开着时想直接写（调完笔设置立刻写，很常见）
        //      的第一笔会丢；
        //   ② 触摸被 WPF 提升为 Stylus，吞掉第一根手指的按下可能破坏双指手势。
        // ⇒ 只能在**抬起时**用位移判定这次到底是"单击"还是"书写"。
        //
        // 时序（已按 WPF 源码核实 InkCanvas.cs: 先 Strokes.Add 再 OnStrokeCollected）：
        //   按下(隧道) → [InkCanvas 落笔] → 抬起(隧道) → Strokes.Add（此时才记撤销） → StrokeCollected
        //   ▲ 我们在这里判定                      ▲ 我们在这里改提交类型，来得及影响撤销记录
        // 所以：单击 → 全程保持 CodeInput → "新增"和"删除"都不进撤销栈（零污染）；
        //       书写 → 抬起时立刻恢复 UserInput → 这一次照常进撤销栈。
        //
        // 无痕移除的写法照抄既有先例 MW_PenSettings.cs 的激光笔淡出分支
        // （_currentCommitType = CodeInput; Strokes.Remove(...); 恢复 UserInput）。

        /// <summary>抬起时位移不超过这么多像素才算"单击"（与 MW_Capture 的粘贴气泡判定同一量级）。</summary>
        private const double DismissTapMaxMovePx = 6.0;

        /// <summary>
        /// "同一次物理按下"的判定窗口（ms）。
        /// ★ 实测（手写板 Ink 模式）：一次落笔会先来 Stylus 版的按下、紧接着再来一个被"提升"出来的
        /// Mouse 版按下，两者相隔只有几毫秒。若当成两次独立按下，第二次会走"没收面板"分支，
        /// 把第一次的待定状态和 CodeInput 一起冲掉 → 抬起判定失效 → 点被留下。
        /// 80ms 远大于提升间隔（几毫秒），又远小于人手连按两次的间隔，能干净区分。
        /// </summary>
        private const int DismissSamePressMs = 80;

        /// <summary>
        /// "要丢弃"标记的保鲜期（ms）。超时说明这次按下早已结束（例如那一下没产生笔迹），
        /// 不再认这个标记——否则会把之后一根真实的顿点也误删。
        /// </summary>
        private const int DismissDiscardFreshMs = 500;

        private bool _dismissPressPending;   // 本次按下收起了面板，抬起时要判定
        private bool _dismissPressDiscard;   // 判定为单击：随之而来的那根笔迹要丢弃
        private Point _dismissPressDownPos;
        private int _dismissPressDownTick;      // 按下时刻（TickCount）
        private int _dismissPressDiscardTick;   // 判定为"要丢弃"的时刻

        /// <summary>按下时调用（鼠标/触笔两条通道共用）。dismissedSomething = 本次按下是否真的收起了可见面板。</summary>
        private void DismissPressDown(Point pos, bool dismissedSomething)
        {
            int now = Environment.TickCount;

            // 同一次按下的第二条通道上报：保留已有状态（尤其不能把 CodeInput 冲掉），只更新坐标
            if (_dismissPressPending && unchecked(now - _dismissPressDownTick) < DismissSamePressMs)
            {
                _dismissPressDownPos = pos;
                return;
            }

            if (dismissedSomething)
            {
                _dismissPressPending = true;
                _dismissPressDiscard = false;
                _dismissPressDownPos = pos;
                _dismissPressDownTick = now;
                // 先假定要丢弃：这样连"新增笔迹"那一步都不会写进撤销栈。
                // 若抬起判定是"书写"，会在 DismissPressUp 里立刻恢复。
                _currentCommitType = CommitReason.CodeInput;
            }
            else if (_dismissPressPending)
            {
                // 新的一次按下、且没收面板，但上次的待定状态还挂着（例如抬起事件丢了）
                // → 防御性复位，否则 CodeInput 会一直挂着，导致后续所有输入都不进撤销栈（很严重）。
                LogHelper.WriteLogToFile(
                    "[Popup] 防御性复位：新按下时没收面板，但上次的待定状态还挂着（上一次抬起没结算）",
                    LogHelper.LogType.Event);
                ResetDismissPress();
            }
        }

        /// <summary>抬起时调用（鼠标/触笔两条通道共用；重复调用安全）。</summary>
        private void DismissPressUp(Point pos)
        {
            // 未标记 → 直接返回。但若"刚收过面板"（2 秒内），说明待定状态被冲掉了，
            // 这是"点了外面还留点"的唯一可疑路径，必须打出来才知道（否则整条链是静默的）。
            if (!_dismissPressPending)
            {
                if (DismissPressRecent(2000))
                    LogHelper.WriteLogToFile(
                        "[Popup] ⚠ 抬起时待定状态已空（按下标记被冲掉）→ 这一次不会丢弃点",
                        LogHelper.LogType.Event);
                return;
            }

            // 同一次抬起的第二条通道上报 → 不重复判定，避免刷重复日志
            if (_dismissPressDiscard) return;

            if ((pos - _dismissPressDownPos).Length <= DismissTapMaxMovePx)
            {
                // 单击：这就是"点外收起"的那一下，不该留下墨迹。
                // 保持 CodeInput，等笔迹提交后由 TryDiscardDismissTapStroke 静默移除。
                _dismissPressDiscard = true;
                _dismissPressDiscardTick = Environment.TickCount;
                LogHelper.WriteLogToFile(
                    $"[Popup] 抬起判定为单击（位移 {(pos - _dismissPressDownPos).Length:F1}px ≤ {DismissTapMaxMovePx}）→ 将丢弃随之产生的点",
                    LogHelper.LogType.Event);
            }
            else
            {
                // 是书写：这次笔迹要正常进撤销栈 → 立刻把提交类型恢复回去。
                // 隧道抬起事件早于 InkCanvas 提交笔迹，所以这里来得及。
                LogHelper.WriteLogToFile(
                    $"[Popup] 抬起判定为书写（位移 {(pos - _dismissPressDownPos).Length:F1}px）→ 笔迹照常进撤销栈",
                    LogHelper.LogType.Event);
                ResetDismissPress();
            }
        }

        /// <summary>
        /// 在 inkCanvas_StrokeCollected 最开始调用。
        /// 返回 true = 这根笔迹是"收起面板的那次单击"产生的，已被静默丢弃（调用方应直接 return，跳过后续处理）。
        /// </summary>
        private bool TryDiscardDismissTapStroke(Stroke stroke)
        {
            if (!_dismissPressDiscard)
            {
                // 诊断：刚收过面板、却有笔迹到达，而"要丢弃"标记已失效 → 这根点会被留下。
                // 这就是"点了外面还留点"的直接证据（2 秒门槛，普通书写不会刷这条）。
                if (DismissPressRecent(2000))
                    LogHelper.WriteLogToFile(
                        "[Popup] ⚠ 有笔迹到达，但\"要丢弃\"标记已失效 → 这根点会留在画布上",
                        LogHelper.LogType.Event);
                return false;
            }

            // 保鲜期：标记过期说明这次按下早就结束了（例如那一下没产生笔迹），
            // 不认它，免得把之后一根真实的顿点误删。
            if (unchecked(Environment.TickCount - _dismissPressDiscardTick) > DismissDiscardFreshMs)
            {
                ResetDismissPress();
                return false;
            }

            try
            {
                _currentCommitType = CommitReason.CodeInput;   // 删除这一步也不进撤销栈
                inkCanvas.Strokes.Remove(stroke);
                LogHelper.WriteLogToFile("[Popup] 已丢弃\"点外收起\"时顺手产生的点笔迹（新增与删除都不进撤销栈）",
                    LogHelper.LogType.Event);
            }
            finally
            {
                ResetDismissPress();                          // 无论如何都要复位，不能让 CodeInput 泄漏
            }
            return true;
        }

        private void ResetDismissPress()
        {
            _dismissPressPending = false;
            _dismissPressDiscard = false;
            _currentCommitType = CommitReason.UserInput;
        }

        /// <summary>
        /// 最近是否发生过"点外收起面板"的按下。仅用于给诊断日志设门槛——
        /// 上面两个方法在每次落笔/每次抬起都会被调用，不加门槛会把日志刷爆。
        /// </summary>
        private bool DismissPressRecent(int ms)
        {
            return _dismissPressDownTick != 0
                   && unchecked(Environment.TickCount - _dismissPressDownTick) < ms;
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
