using Ink_Canvas.Helpers;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Ink;

namespace Ink_Canvas
{
    /// <summary>
    /// 手写公式识别测试窗口·精简版（验证 MathGraph 模块整条链路）
    ///
    /// 窗口职责：点"数学输入面板"→ 手写公式 → 插入 → 画布上出现函数图像
    /// （早期版本的文本识别诊断代码已删除——那套链路验证完毕，使命结束）
    ///
    /// 版本演进备忘（供后期翻阅）：
    /// 1. WPF InkAnalyzer 识别空白（词节点有但 GetRecognizedString 空）→ 放弃
    /// 2. Microsoft.Ink 直连识别器：中文识别完美，但本机无英文/公式识别器 → 放弃
    /// 3. micaut.MathInputControl（微软数学输入面板）：识别公式完美 → 当前方案 ✅
    /// </summary>
    public partial class HandwritingTestWindow : Window
    {
        /// <summary>
        /// 数学输入面板的 COM 控件实例
        /// 必须存成字段：局部变量方法一结束就被 GC，面板窗口会跟着消失
        /// </summary>
        private Microsoft.MathInput.MathInputControlClass _mathInput;

        public HandwritingTestWindow()
        {
            InitializeComponent();

            //画布默认笔迹样式（此窗口主要展示图像，手写都在面板里进行）
            TestInkCanvas.DefaultDrawingAttributes.Color = System.Windows.Media.Colors.Black;
            TestInkCanvas.DefaultDrawingAttributes.Width = 2.5;
            TestInkCanvas.DefaultDrawingAttributes.Height = 2.5;
        }

        /// <summary>
        /// 【数学输入面板】按钮：调用系统自带的数学输入控件（micaut.dll）
        /// 弹出独立手写窗口 → 写公式 → 点"插入" → Insert 事件带回 MathML
        /// </summary>
        private void BtnMathInput_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //第一次点按钮时创建面板并订阅事件，之后只负责显示（保留实例，防 GC 回收）
                if (_mathInput == null)
                {
                    _mathInput = new Microsoft.MathInput.MathInputControlClass();

                    //Insert 事件：用户点"插入"按钮时触发，参数 = 识别出的公式（MathML XML）
                    _mathInput.Insert += (mathml) =>
                    {
                        //COM 事件回调不保证在 WPF UI 线程，操作界面控件要转回 UI 线程
                        Dispatcher.Invoke(() =>
                        {
                            //完整链路：MathML → 解析 → 采样 → 图像笔迹
                            bool ok = Ink_Canvas.MathGraph.MathGraphModule.TryBuildGraph(
                                mathml,
                                TestInkCanvas.ActualWidth,
                                TestInkCanvas.ActualHeight,
                                out var graphStrokes,
                                out string message);

                            if (ok)
                            {
                                //识别成功：清掉旧内容，画布显示函数图像
                                TestInkCanvas.Strokes.Clear();
                                TestInkCanvas.Strokes.Add(graphStrokes);
                            }
                            TextBlockBestResult.Text = message;
                            AppendDiagnostics(new StringBuilder("收到 MathML：\n" + mathml));
                        });
                        LogHelper.WriteLogToFile("[手写测试] MathML: " + mathml, LogHelper.LogType.Event);
                    };

                    //Close 事件：用户点面板右上角的 X
                    _mathInput.Close += () =>
                    {
                        Dispatcher.Invoke(() => AppendDiagnostics(new StringBuilder("面板已关闭")));
                    };
                }

                //显示面板（如果已打开则不重复显示）
                bool visible = false;
                _mathInput.IsVisible(out visible);
                if (!visible)
                {
                    _mathInput.Show();
                }
            }
            catch (System.Exception ex)
            {
                AppendDiagnostics(new StringBuilder("数学面板异常：" + ex.GetType().Name + " —— " + ex.Message));
                LogHelper.WriteLogToFile("[手写测试] 数学面板异常 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>把新的诊断信息显示到诊断区（MathML 原文 / 异常详情）</summary>
        private void AppendDiagnostics(StringBuilder newInfo)
        {
            TextBlockDiagnostics.Text = newInfo.ToString();
            LogHelper.WriteLogToFile("[手写测试] " + newInfo.ToString().Replace("\n", " | "), LogHelper.LogType.Event);
        }

        /// <summary>【清除笔迹】按钮</summary>
        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            TestInkCanvas.Strokes.Clear();
            TextBlockBestResult.Text = "（尚未识别）";
        }

        /// <summary>【橡皮模式】按钮：在手写/擦除间切换</summary>
        private void BtnToggleEraser_Click(object sender, RoutedEventArgs e)
        {
            if (TestInkCanvas.EditingMode == InkCanvasEditingMode.Ink)
            {
                TestInkCanvas.EditingMode = InkCanvasEditingMode.EraseByPoint;
            }
            else
            {
                TestInkCanvas.EditingMode = InkCanvasEditingMode.Ink;
            }
        }

        /// <summary>
        /// 窗口关闭时释放 COM 控件
        /// 不释放的话，COM 非托管资源会一直挂着，程序退出时可能拖泥带水
        /// </summary>
        protected override void OnClosed(System.EventArgs e)
        {
            if (_mathInput != null)
            {
                try
                {
                    _mathInput.Hide();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_mathInput);
                    _mathInput = null;
                }
                catch { }
            }
            base.OnClosed(e);
        }
    }
}
