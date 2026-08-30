using System;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;
using System.Windows.Interop;
using Ink_Canvas.Helpers;

namespace Ink_Canvas
{
    // MainWindow 的分部类：函数识别入口（图形面板 fx 按钮）
    //
    // 独立性设计（和 MathGraph 模块同样的原则）：
    // 1. 本文件只做"胶水"——把主画板和数学输入面板连起来
    // 2. 数学逻辑全部在 MathGraph/ 文件夹，本文件不含任何公式解析代码
    // 3. 不想要这个功能时：删本文件 + XAML 里那一个按钮即可，主程序无感知
    //
    // 置顶方案（SetOwnerWindow）：
    // 把数学面板的"属主窗口"设为主窗口。Windows 天然保证属主永远在子窗口之下，
    // 面板自动浮在主窗口上方——不需要任何 Topmost 让位/API 置顶的补丁代码。
    // （早期版本的 Topmost 双向切换 + SetWindowPos 方案已整体废弃）
    // 注：框选识别（右键"识别为函数"）功能已整体删除，只保留 fx 面板直写入口。
    public partial class MainWindow
    {
        /// <summary>
        /// 数学输入面板的 COM 实例
        /// 存成字段原因有二：防 GC 提前回收（面板会消失）；重复点击按钮只显示同一个面板
        /// </summary>
        private Microsoft.MathInput.MathInputControlClass _mathInputPanel;

        /// <summary>
        /// 面板标题（SetCaptionText 设置后即成为 Win32 窗口标题，
        /// 也是位置记忆功能里 FindWindow 定位面板窗口的"身份证"——改标题必须同步改这里）
        /// </summary>
        private const string MathPanelCaption = "手写函数识别作图";

        /// <summary>
        /// 【fx 识别函数】按钮：弹数学输入面板
        /// 使用方式（老师视角）：
        ///   点按钮 → 面板里写公式 → 插入 → 画板上出现函数图像
        ///   面板开着期间仍然可以在画板上批注（面板浮在画板上方）
        /// </summary>
        private void BtnRecognizeFunction_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //面板创建 + 事件订阅统一走 EnsureMathPanelCreated（框选识别共用）
                EnsureMathPanelCreated();

                //显示面板并应用记忆位置（已打开则不重复 Show）
                ShowMathPanel();
            }
            catch (Exception ex)
            {
                //面板是系统 COM 组件，个别精简版 Windows 可能没装——失败时给明确提示而不是崩溃
                MessageBox.Show("无法打开数学输入面板，此功能需要系统组件\"数学识别器\"（micaut.dll）。\n" + ex.Message,
                    "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                LogHelper.WriteLogToFile("[函数识别] 数学面板异常 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// 收起数学面板（唯一的"关面板"出口）
        /// Insert 完成、面板点 X 两条路都走这里——用 Hide 而不是销毁，
        /// 下次点 fx 直接复用同一个面板实例，秒开无冷启动
        /// </summary>
        private void CloseMathPanel()
        {
            try { _mathInputPanel?.Hide(); } catch { }
            //面板关了，fx 图标熄灭（与 ShowMathPanel 里的点亮配对）
            try { BorderShapeIcon_27.Tag = null; } catch { }
        }

        // ==================== 面板位置/大小（原生 SetPosition + Win32 兜底） ====================
        //
        // 【micaut 接口速查】——来自桌面两份经验文件（方案A实施手册 / COM逆向笔记），后续维护先看这里：
        //   · SetPosition(L, T, R, B)  设置面板屏幕坐标（左/上/右/下，物理像素）。
        //     ★ 必须在 Show() 之前调用才保证生效；这是"首次打开位置正确"的唯一可靠手段
        //     （冷启动期间 FindWindow 找不到窗口，Win32 方案全部失效）。
        //   · GetPosition(out L, out T, out R, out B)  读取当前矩形（位置+大小一起拿）。
        //   · Clear()  清空手写区（本文件在 Insert 成功后调用，下次打开是干净面板）。
        //   · SetCaptionText(标题)  设窗口标题。作用不是给人看——是给 FindWindow 当"身份证"。
        //   · Show() / Hide()  显示/隐藏。用 Hide 不销毁：同一实例反复开关，规避冷启动延迟。
        //   · EnableAutoGrow(bool)  长公式自动扩面板（已在初始化时开启）。
        //   · Win32 窗口类名 "MathInputControl_Window"（逆向笔记）：
        //     GetPosition 拿不到时的兜底查找依据（FindWindow 按类名），比按标题找稳定。
        //   · 最小尺寸 344×261：SetPosition 给再小控件也会自己扩到这个值。
        //   · 冷启动慢：首次 Show 后窗口/标题数秒才就绪，任何依赖"窗口已存在"的逻辑都要重试或规避。
        //
        // 记忆策略（会话级，不写盘）：软件开着期间记住"上次关闭时的位置和大小"，
        // 软件重启后回默认（屏幕居中 + 默认尺寸）——课堂场景每天环境可能变，持久化位置反而容易开在屏幕外。
        //
        // 方案说明：
        // ① 初始定位用控件原生的 SetPosition——micaut 自己的 API，冷启动期间也生效，
        //    完全绕开 FindWindow（旧方案"第一次总在左上角"的根因：窗口/标题未就绪时 FindWindow 必失败）。
        // ② 读取当前实际矩形：优先 GetPosition，拿不到再 FindWindow(类名) + GetWindowRect 兜底。

        //Win32 结构体：窗口矩形（物理像素）
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        //Win32 结构体：显示器信息（含工作区 = 去掉任务栏的可用区域）
        [StructLayout(LayoutKind.Sequential)]
        private struct MONITORINFO
        {
            public int cbSize;
            public RECT rcMonitor;
            public RECT rcWork;
            public uint dwFlags;
        }

        /// <summary>控件注册的 Win32 窗口类名（经验文件逆向得到，比标题更稳定——窗口创建瞬间就有）</summary>
        private const string MathPanelClassName = "MathInputControl_Window";

        //经验文件实测：控件客户区最小 344×261，SetPosition 给太小的值控件会自动扩到这个尺寸
        private const int MathPanelMinW = 344;
        private const int MathPanelMinH = 261;

        /// <summary>默认面板尺寸（600×440：写手写公式舒展、不局促）</summary>
        private const int MathPanelDefaultW = 600;
        private const int MathPanelDefaultH = 440;

        /// <summary>
        /// 会话级记忆：本次软件运行期间，面板上次收起时的矩形（位置+大小一起记）。
        /// 只存内存不写盘——软件重启后自动为 null，面板回默认位置和默认大小。
        /// </summary>
        private RECT? _mathSessionRect = null;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        private static extern IntPtr MonitorFromWindow(IntPtr hwnd, uint dwFlags);

        [DllImport("user32.dll")]
        private static extern bool GetMonitorInfo(IntPtr hMonitor, ref MONITORINFO lpmi);

        private const uint MONITOR_DEFAULTTONEAREST = 2;

        /// <summary>
        /// 显示面板并应用位置（点 fx 按钮入口）
        ///
        /// 【关键修复】初始位置不再依赖 FindWindow+SetWindowPos：
        /// 按经验文件：micaut 有原生 SetPosition(L,T,R,B)（屏幕坐标，Show 前调用即可，不用等窗口就绪）——
        /// 这是根治"第一次打开总是左上角"的方案。
        /// </summary>
        private void ShowMathPanel()
        {
            bool visible = false;
            _mathInputPanel.IsVisible(out visible);

            if (!visible)
            {
                //【Show 之前】先算好位置，直接调控件自己的 SetPosition
                //（经验文件 5.1 节明确写了该方法签名：SetPosition(Left,Top,Right,Bottom)，屏幕坐标）
                CalculateMathPanelInitialPosition(out int left, out int top, out int right, out int bottom);
                try { _mathInputPanel.SetPosition(left, top, right, bottom); } catch { }

                _mathInputPanel.Show();
                //补设标题（Show 后更可靠，不影响位置定位）
                try { _mathInputPanel.SetCaptionText(MathPanelCaption); } catch { }
            }

            //fx 图标（BorderShapeIcon_27）亮蓝：面板开着期间保持，
            //告诉用户"函数识别面板还开着"；关闭时在 CloseMathPanel 里熄灭
            try { BorderShapeIcon_27.Tag = "Active"; } catch { }
        }

        /// <summary>
        /// 计算面板的初始屏幕位置（原生 SetPosition 的 LTRB 参数）。
        /// 优先级：会话记忆（本次运行内拖过/调过大小，且仍可见）→ 屏幕正中央 + 默认大小
        /// </summary>
        private void CalculateMathPanelInitialPosition(
            out int left, out int top, out int right, out int bottom)
        {
            var work = GetMainWindowWorkArea();
            int workW = work.Right - work.Left;
            int workH = work.Bottom - work.Top;

            //① 会话级记忆：本次软件运行期间拖过位置/拉过大小 → 原样恢复（位置和大小一起）
            if (_mathSessionRect.HasValue)
            {
                var r = _mathSessionRect.Value;
                int w = r.Right - r.Left, h = r.Bottom - r.Top;
                if (w >= MathPanelMinW && h >= MathPanelMinH &&
                    IsMathPanelPositionVisible(r.Left, r.Top, w, h))
                {
                    left = r.Left; top = r.Top; right = r.Right; bottom = r.Bottom;
                    return;
                }
            }

            //② 默认：中央偏左 + 默认尺寸。
            //不居中而偏左的原因：教师大屏授课右手持笔/持鼠标，面板放左手边，
            //右手的活动区域（画布中央）不被遮挡——插入的图像在中央偏右仍然可见
            int dw = MathPanelDefaultW, dh = MathPanelDefaultH;
            if (dw > workW) dw = Math.Max(MathPanelMinW, workW - 40);
            if (dh > workH) dh = Math.Max(MathPanelMinH, workH - 40);
            //水平位置：工作区宽度的 1/4 处（面板中心 = 工作区 30% 位置，中央偏左）
            left = work.Left + Math.Max(0, (workW - dw) / 5);
            top = work.Top + (workH - dh) / 2;
            right = left + dw;
            bottom = top + dh;
        }

        /// <summary>
        /// 记忆位置是否仍落在可见屏幕内（防止换了分辨率/拔了显示器后面板"开在屏幕外"）。
        /// 判定标准宽松：面板中心点在工作区内即可见。
        /// </summary>
        private bool IsMathPanelPositionVisible(int x, int y, int w, int h)
        {
            var work = GetMainWindowWorkArea();
            int cx = x + w / 2;
            int cy = y + h / 2;
            return cx >= work.Left && cx <= work.Right
                && cy >= work.Top && cy <= work.Bottom;
        }

        /// <summary>主窗口所在显示器的工作区（物理像素，纯 Win32 取值，避免 DPI 换算误差）</summary>
        private RECT GetMainWindowWorkArea()
        {
            var mi = new MONITORINFO { cbSize = System.Runtime.InteropServices.Marshal.SizeOf(typeof(MONITORINFO)) };
            IntPtr mainHwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            IntPtr monitor = MonitorFromWindow(mainHwnd, MONITOR_DEFAULTTONEAREST);
            if (monitor != IntPtr.Zero && GetMonitorInfo(monitor, ref mi))
                return mi.rcWork;
            //兜底：拿不到显示器信息时给全零矩形（面板就停在系统默认位置，不出错）
            return new RECT();
        }

        /// <summary>
        /// 记下面板当前的矩形（位置+大小）到会话变量（关闭/插入前调用，此时窗口还在、坐标可读）。
        /// ★ 会话级记忆：只存内存 _mathSessionRect，不写 Settings——
        ///   本次软件运行期间拖到哪/拉到多大，下次打开就保持那样；软件重启后自动回默认。
        /// 优先用控件原生 GetPosition，拿不到时退回 FindWindow(类名) + GetWindowRect
        /// （比按标题查找稳定，标题在冷启动早期可能还没更新）。
        /// </summary>
        private void SaveMathPanelPosition()
        {
            try
            {
                int L = 0, T = 0, R = 0, B = 0;
                bool ok = false;

                try
                {
                    //首选：控件原生 GetPosition（完全不碰 Win32，最可靠）
                    //注意：tlbimp 生成的互操作签名是 GetPosition(out Left, out Top, out Right, out Bottom)
                    //4 个参数全是 out（逆向笔记 5.1 节验证过）
                    int oL, oT, oR, oB;
                    _mathInputPanel.GetPosition(out oL, out oT, out oR, out oB);
                    L = oL; T = oT; R = oR; B = oB;
                    ok = R > L && B > T;
                }
                catch { ok = false; }

                if (!ok)
                {
                    //兜底：按【窗口类名】查找（经验文件逆向得到 MathInputControl_Window，
                    //比标题查找稳定——窗口创建瞬间就有类名，不存在"标题还没设置好"的时间窗问题）
                    IntPtr hwnd = FindWindow(MathPanelClassName, null);
                    if (hwnd == IntPtr.Zero) return;
                    if (!GetWindowRect(hwnd, out RECT rc)) return;
                    L = rc.Left; T = rc.Top; R = rc.Right; B = rc.Bottom;
                }

                _mathSessionRect = new RECT { Left = L, Top = T, Right = R, Bottom = B };
            }
            catch { }
        }

        /// <summary>确保数学面板已创建（点 fx 按钮触发，实例全局唯一）</summary>
        private void EnsureMathPanelCreated()
        {
            if (_mathInputPanel != null) return;

            _mathInputPanel = new Microsoft.MathInput.MathInputControlClass();

            //开启自动增长：手写内容接近边界时面板自动向右下方扩展，
            //复杂/较长的公式才写得下（面板大小原来锁死，长公式输不进）。
            //位置记忆（FindWindow + SetWindowPos）只在 Show 后执行一次、只挪位置不动大小，
            //与自动增长互不干扰。
            _mathInputPanel.EnableAutoGrow(true);
            //面板内部的英文提示是自绘的改不了，只有标题栏能中文化
            _mathInputPanel.SetCaptionText(MathPanelCaption);

            //设属主窗口：Windows 自动保证面板在主窗口之上（置顶问题的根治方案）
            var hwndSource = (HwndSource)PresentationSource.FromVisual(this);
            _mathInputPanel.SetOwnerWindow(hwndSource.Handle.ToInt64());

            //Insert 事件：老师在面板点"插入" → 画布中央出函数图像
            _mathInputPanel.Insert += (mathml) =>
            {
                //原始 MathML 落盘：这是后续所有解析的唯一输入，必须能事后回看。
                //（此前整条链路零留痕，识别器实际输出什么无从查证，分式/幂解析失败无法定位）
                LogHelper.WriteLogToFile("[函数识别] 收到插入，原始 MathML：" + mathml,
                    LogHelper.LogType.Event);

                Dispatcher.BeginInvoke((Action)(() =>
                {
                    //插入成功前先记住面板当前位置（面板不关闭，记忆用于下次点 fx 时恢复）
                    SaveMathPanelPosition();

                    //【清屏】出图后自动清手写区，老师直接写下一个函数（连续输入流不中断）。
                    //Clear() 是接口原生方法（逆向笔记 5.1 节：void Clear()，清除面板内所有笔迹）
                    try { _mathInputPanel?.Clear(); } catch { }

                    //【面板保持打开】连续输入多个函数：Insert 不关面板，
                    //点 X（Close 事件）才收起——对齐"多函数分次插入"的使用流
                    OnMathInserted(mathml);
                }));
            };

            //Close 事件：老师点 X 放弃
            _mathInputPanel.Close += () =>
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    SaveMathPanelPosition(); //点 X 前也记住面板位置（拖到哪，下次开在哪）
                    CloseMathPanel();
                }));
            };
        }

        /// <summary>
        /// 老师在面板点"插入"后：MathML → 函数图像 → 放到主画板
        /// </summary>
        private void OnMathInserted(string mathml)
        {
            try
            {
                //正方形取景框：所有函数统一 480×480（约 ±8 单位见方）。
                //不用整块画布当取景框的原因：全屏 1920×1080 时横向 ±32 单位、纵向只有 ±18 单位，
                //反比例函数被拉得左右特长上下特短、二次函数窄得像根竖线——比例失衡不美观。
                //统一正方形后所有函数同一比例尺、同一窗口大小，观感一致。
                double w = 480, h = 480;

                //完整链路：MathML → 解析 → 采样 → 坐标系+曲线笔迹
                bool ok = MathGraph.MathGraphModule.TryBuildGraph(mathml, w, h,
                    out StrokeCollection graphStrokes, out string message,
                    inkCanvas.DefaultDrawingAttributes); //曲线颜色/粗细跟随当前画笔（与图形同一大原则）

                if (!ok)
                {
                    //解析失败（含参、下标等暂不支持的情况）：提示原因，画板不动
                    LogHelper.WriteLogToFile("[函数识别] 出图失败：" + message + " | MathML: " + mathml,
                        LogHelper.LogType.Error);
                    MessageBox.Show(message, "暂不能画图", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                //图像平移（多函数共系的核心逻辑）：
                //GraphBuilder 生成的笔迹在 480×480 取景框坐标系里，原点圆点在取景框中心 (w/2, h/2)。
                //① 画布上已有函数图像 → 找到最近插入那组的原点圆点，把新图的原点对齐过去——
                //   y=x 和 y=x² 自动落在同一坐标系上（数学上正确的"共系"行为），不用再拖
                //② 画布上没有函数图像（第一次插入）→ 原点放画布正中央
                //注意：平移目标必须用 inkCanvas 的真实宽高——用取景框的 w/h 会把图挪到画布左上角
                var graphBounds = graphStrokes.GetBounds();
                Point targetOrigin = FindLastGraphOrigin() ?? new Point(inkCanvas.ActualWidth / 2, inkCanvas.ActualHeight / 2);
                double dx = targetOrigin.X - w / 2;
                double dy = targetOrigin.Y - h / 2;
                var matrix = new System.Windows.Media.Matrix();
                matrix.Translate(dx, dy);
                graphStrokes.Transform(matrix, false);

                inkCanvas.Strokes.Add(graphStrokes);

                SelectInsertedGraph(graphStrokes);
            }
            catch (Exception ex)
            {
                //兜底：任何意外都不允许影响主画板
                MessageBox.Show("生成图像失败：" + ex.Message, "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                LogHelper.WriteLogToFile("[函数识别] 出图异常 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// 出图后自动选中生成的图像。
        /// 函数图像由上百条刻度/曲线笔迹组成，老师事后想挪动或缩放时手工框选很难选干净，
        /// 所以插入即选中，直接可拖动、可缩放、可旋转。
        /// 做法参照 MW_CustomShapes.cs 的自定义图形插入（那里已验证可用）。
        /// </summary>
        private void SelectInsertedGraph(StrokeCollection graphStrokes)
        {
            //统一走图形管理入口（打标签 + 自动选中 + 控制条），与其他图形行为一致
            InsertGraphStrokes(graphStrokes);
        }

        /// <summary>
        /// 程序退出时释放数学面板 COM 资源
        /// （窗口关闭时调用——面板可能整个程序生命周期都要用，不提前释放）
        /// </summary>
        public void ReleaseMathInputPanel()
        {
            if (_mathInputPanel != null)
            {
                try
                {
                    _mathInputPanel.Hide();
                    Marshal.ReleaseComObject(_mathInputPanel);
                }
                catch { }
                _mathInputPanel = null;
            }
        }
    }
}
