using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using Ink_Canvas.Helpers;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 的分部类：图形笔迹统一管理（路线 A —— 轻量标签方案）
    ///
    /// 背景：画布上"一切都是笔迹"，但有三类内容其实是"图形"：
    ///   1. 函数识别生成的图像（MW_MathGraph）
    ///   2. 图形面板拖拽画的图形（MW_ShapeDrawing）
    ///   3. 手写识别成图形 / 自定义图库插入（MW_SimulatePressure / MW_CustomShapes）
    /// 它们的"插入后行为"应该一致：打标签 + 自动选中，方便直接平移/缩放。
    ///
    /// 方案：不改变存储模型（仍是普通笔迹，翻页/撤销/保存零改动），
    /// 只给图形笔迹附加一个"组 ID"属性（WPF Stroke 原生支持附加数据），
    /// 后续可基于该标签做"框选时整组吸附"等统一交互。
    ///
    /// 可弃性：本文件只被各插入点调用，删除调用即可整体移除，不影响主程序。
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>图形组标签的唯一标识（存在每条 Stroke 的附加属性里）</summary>
        private static readonly Guid GraphGroupGuid = new Guid("D4F5A6B7-C8E9-4F0A-9B1C-2D3E4F5A6B7C");

        /// <summary>组 ID 计数器：每画一个图形自增一次，同一次插入的所有笔迹共用一个 ID</summary>
        private static long _graphGroupSeq = 0;

        /// <summary>
        /// "一次性选中"标志：true 表示当前选中态来自图形插入（而非用户主动选择）。
        /// 背景实测：inkCanvas.Select() 会偷偷把 EditingMode 切成 Select，
        /// 取消选中后模式仍停留在 Select（光标一直是选择箭头、落笔变套索）。
        /// 按一次性选中交互：这种选中被取消时应恢复笔模式；
        /// 用户主动用选择工具（BtnSelect_Click）时会清掉本标志，不受影响。
        /// </summary>
        private bool _isOneShotGraphSelection = false;

        /// <summary>
        /// 立即结束"插入后一次性选中"（不等用户点空白）。
        /// 使用场景：老师激活任何工具（图形/笔/橡皮）的瞬间就收起选区遮罩，
        /// 让画布恢复干净——否则遮罩会吃掉第一笔的鼠标事件，导致"点完图形按钮画不出来"
        /// （得先在空白处点一下取消选中才能画，交互断裂）。
        /// 与 TryEndOneShotSelection 的区别：那个是"用户取消选中后"的被动收尾，
        /// 这个是"用户激活新工具"的主动清场，且不恢复笔模式（新工具马上要用）。
        /// </summary>
        private void EndOneShotSelectionNow()
        {
            if (!_isOneShotGraphSelection) return;
            _isOneShotGraphSelection = false;
            try
            {
                //关键：先记住当前模式。WPF 的 InkCanvas.Select() 有个文档行为——
                //调用它会把 EditingMode 强制切成 Select，且不恢复。
                //不记的话：本方法在"点橡皮→设置擦除模式→再取消选中"的顺序中被调用时，
                //刚设置好的擦除模式会被 Select 覆盖成 Select 模式，工具状态全乱
                //（用户反馈的"点橡皮后显示还是笔但又能擦"即此因）。
                var modeBefore = inkCanvas.EditingMode;
                //恢复模式会再触发一次 EditingModeChanged，事件里有 forcePointEraser 翻转逻辑
                //（笔/橡皮形态互切的 hack），一并记住恢复，防止橡皮形态被意外互换
                bool eraserShapeBefore = forcePointEraser;

                //与空白处单击取消选中的路径完全一致（屏蔽 SelectionChanged 快照副作用）
                isProgramChangeStrokeSelection = true;
                inkCanvas.Select(new StrokeCollection());
                isProgramChangeStrokeSelection = false;

                //取消选中后把模式恢复回去（Select 已经把它改成了 Select）
                if (modeBefore != InkCanvasEditingMode.Select)
                    inkCanvas.EditingMode = modeBefore;
                forcePointEraser = eraserShapeBefore;
                UpdateEraserIcon(); //事件里按翻转值刷过图标，这里按恢复后的正确值再刷一次

                GridInkCanvasSelectionCover.Visibility = Visibility.Collapsed;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLogToFile("[图形管理] 结束一次性选中失败 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// 一次性选中的收尾：若当前选中来自图形插入且已被取消，恢复笔模式。
        /// 各取消选中的路径（鼠标/触摸/事件驱动）统一调它，逻辑只此一份。
        /// </summary>
        private void TryEndOneShotSelection()
        {
            if (!_isOneShotGraphSelection) return;
            if (inkCanvas.GetSelectedStrokes().Count > 0) return; //选中还在（拖动/调整中），不动
            _isOneShotGraphSelection = false;
            if (drawingShapeMode != 0) return; //防御：用户已激活图形工具时不恢复笔（会杀掉刚选的图形模式）

            //延迟一拍再恢复笔模式：EditingMode 赋值会"内联"触发 SelectionChanged，
            //本方法常在事件链内被调用——若在这里同步改模式，会反过来覆盖
            //外层正在设置的橡皮/其他工具（"点橡皮显示橡皮实际是笔"的错乱根因）。
            //派发到队尾执行，等外层工具切换全部落地后再恢复。
            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    //仅在模式仍停留在 Select（=没有工具被激活过）时才恢复笔。
                    //若用户已点了橡皮/荧光笔等，模式会被同步改成对应值——
                    //此时再恢复笔会覆盖用户刚选的工具（"点橡皮后还是笔、还能写"的根因）。
                    //原则：选中框只属于选中框的操作；外部任何操作后一切照旧，如同没有选中这回事。
                    if (inkCanvas.EditingMode == InkCanvasEditingMode.Select)
                        BtnPen_Click(BtnPen, null); //恢复笔模式（内部会设 EditingMode=Ink）
                }
                catch { }
            }));
        }

        /// <summary>
        /// 给一组笔迹打上"同一个图形"的标签（供将来整组吸附使用）。
        /// 同一次插入的坐标系 + 曲线 = 一个组，框选碰到任何一条就能拉出整组。
        /// </summary>
        private void TagAsGraph(StrokeCollection strokes)
        {
            _graphGroupSeq++;
            foreach (Stroke s in strokes)
            {
                try { s.AddPropertyData(GraphGroupGuid, _graphGroupSeq); } catch { }
            }
        }

        /// <summary>
        /// 找到画布上最近一次插入的函数图像组的"原点圆点"位置。
        /// 原理：函数图像组的原点是一个孤立的单点 Stroke（GraphBuilder 生成，
        /// 只含 1 个采样点且组标签值最大 = 最近插入）。返回 null = 画布上没有函数图像。
        /// 用途：多函数共系——新插入的函数把原点对齐到这里，y=x 和 y=x² 同一坐标系。
        /// </summary>
        internal Point? FindLastGraphOrigin()
        {
            try
            {
                Point? best = null;
                int bestSeq = int.MinValue;
                foreach (Stroke s in inkCanvas.Strokes)
                {
                    //有组标签才可能是函数/图形笔迹
                    if (!s.ContainsPropertyData(GraphGroupGuid)) continue;
                    //原点圆点的特征：单点笔迹（一条 Stroke 只有一个 StylusPoint）
                    if (s.StylusPoints.Count != 1) continue;
                    int seq = (int)s.GetPropertyData(GraphGroupGuid);
                    if (seq <= bestSeq) continue;
                    bestSeq = seq;
                    var p = s.StylusPoints[0];
                    best = new Point(p.X, p.Y);
                }
                return best;
            }
            catch { return null; }
        }

        /// <summary>
        /// 图形插入的统一入口：打标签 → 设笔模式 → 自动选中 → 显示选区控制条。
        /// 三个场景（函数图像 / 面板图形 / 手写识别图形）都调它，行为保证一致。
        ///
        /// 顺序说明（重要，实测结论）：
        /// 必须先设 EditingMode 再 Select。WPF 的 InkCanvas 在 EditingMode 被赋值时
        /// （哪怕赋的是当前同值）会清空选中集合，顺序反了选中就会被自己抹掉。
        ///
        /// 执行时机：用 Dispatcher 异步派发，确保不在 StrokeCollected 等事件内部
        /// 执行选中（事件内部对象状态未稳定，选中有被内部逻辑覆盖的风险）。
        /// </summary>
        private void InsertGraphStrokes(StrokeCollection graphStrokes)
        {
            if (graphStrokes == null || graphStrokes.Count == 0) return;
            TagAsGraph(graphStrokes);

            Dispatcher.BeginInvoke(new Action(() =>
            {
                try
                {
                    //① 先定型笔输入（若已是 Ink 模式，这行是幂等的安全垫）
                    forceEraser = false;
                    //注意：InkCanvasEditingMode 的命名空间是 System.Windows.Controls（不是 Ink）
                    inkCanvas.EditingMode = InkCanvasEditingMode.Ink;

                    //② 再选中（屏蔽 SelectionChanged 的快照副作用，快照由选中后重新捕获）
                    //注意：Select() 会把 EditingMode 切成 Select（WPF 行为），
                    //靠 _isOneShotGraphSelection 标志在取消选中时把笔模式换回来
                    isProgramChangeStrokeSelection = true;
                    try { inkCanvas.Select(graphStrokes); } catch { }
                    isProgramChangeStrokeSelection = false;
                    _isOneShotGraphSelection = inkCanvas.GetSelectedStrokes().Count > 0;

                    //③ 显示选区遮罩与控制条（拖动/缩放/旋转立即可用）
                    GridInkCanvasSelectionCover.Visibility = Visibility.Visible;
                    inkCanvas_SelectionChanged(inkCanvas, EventArgs.Empty);
                    updateBorderStrokeSelectionControlLocation();
                }
                catch (Exception ex)
                {
                    LogHelper.WriteLogToFile("[图形管理] 自动选中失败 " + ex, LogHelper.LogType.Error);
                }
            }));
        }

        #region 整组擦除（渐近线虚线"一擦即整条消失"）

        /// <summary>
        /// 整组擦除的重入保护。
        /// true = 本模块正在主动移除整组笔迹，此时再收到擦除事件一律取消，
        /// 防止"移除过程中又触发移除"的连锁反应。
        /// </summary>
        private bool _isRemovingGraphGroup = false;

        /// <summary>
        /// 初始化：订阅 InkCanvas 的"笔迹即将被擦除"事件。
        /// 在构造函数里调用一次（见 MW_Init.cs），之后常驻生效。
        /// </summary>
        private void InitGraphStrokeGroupErasing()
        {
            inkCanvas.StrokeErasing += InkCanvas_StrokeErasing;
        }

        /// <summary>
        /// 擦除拦截：橡皮碰到"图形组"里的任意一条笔迹时，
        /// 取消默认的"只擦这一条"，改为把同组笔迹一次性全部移除。
        ///
        /// 解决的原始问题（DEVELOPMENT.md 问题二）：
        /// 双曲线渐近线是拆段式虚线（几十条 5px 小笔迹），
        /// 按笔迹擦除时一次只能擦掉一小段，体验极差。
        ///
        /// 为什么不做成"单条虚线笔迹"：
        /// WPF 墨迹模型中一条 Stroke 是连续笔迹（笔尖扫过必留墨），
        /// 数学上无法渲染出"断开的虚线间隙"，原定 OutlineStrokes+DashArray 方案不可行。
        /// 改用本方案后虚线外观不变，且一碰即整条消失。
        ///
        /// 行为说明：
        /// - 普通手写笔迹（无组标签）：完全不干预，走默认擦除；
        /// - 点擦除模式（EraseByPoint）：不干预，保留"精细雕刻"语义；
        /// - 撤销历史：Strokes.Remove(整组) 会触发一次 StrokesChanged，
        ///   TimeMachine 自动把整组记为一个擦除单元，Ctrl+Z 整组恢复，无需额外代码。
        /// </summary>
        private void InkCanvas_StrokeErasing(object sender, InkCanvasStrokeErasingEventArgs e)
        {
            try
            {
                // 自己正在移除整组期间收到的事件：一律取消，避免重入
                if (_isRemovingGraphGroup)
                {
                    e.Cancel = true;
                    return;
                }

                // 只拦截"按笔迹擦除"模式；点擦除保持精细雕刻语义
                // 注意：InkCanvasEditingMode 的命名空间是 System.Windows.Controls
                if (inkCanvas.ActiveEditingMode != InkCanvasEditingMode.EraseByStroke) return;

                // 读取这条笔迹的组标签；没有标签 = 普通手写笔迹，交给默认擦除
                object tag = null;
                try { tag = e.Stroke.GetPropertyData(GraphGroupGuid); }
                catch { /* 笔迹没有该属性时会抛异常，视为无标签 */ }
                if (tag == null) return;

                // 是图形笔迹：取消"只擦这一条"，改为整组移除
                e.Cancel = true;

                // 橡皮一次滑动可能先后碰到组内多条笔迹：
                // 第一条已把整组移除，后面的笔迹已不在画布上，跳过即可
                if (!inkCanvas.Strokes.Contains(e.Stroke)) return;

                // 收集画布上所有带相同组标签的笔迹
                var group = new StrokeCollection();
                foreach (Stroke s in inkCanvas.Strokes)
                {
                    try
                    {
                        if (object.Equals(s.GetPropertyData(GraphGroupGuid), tag))
                            group.Add(s);
                    }
                    catch { /* 无标签的笔迹，跳过 */ }
                }
                if (group.Count == 0) return;

                // 一次性移除整组（重入保护 + 异常兜底，保证标志位一定被复位）
                _isRemovingGraphGroup = true;
                try { inkCanvas.Strokes.Remove(group); }
                finally { _isRemovingGraphGroup = false; }
            }
            catch (Exception ex)
            {
                // 模块自包含异常处理：拦截失败时退回默认擦除，不影响主流程
                LogHelper.WriteLogToFile("[图形管理] 整组擦除失败 " + ex, LogHelper.LogType.Error);
            }
        }

        #endregion 整组擦除
    }
}
