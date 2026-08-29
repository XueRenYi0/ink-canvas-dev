using System;
using System.Runtime.InteropServices;
using System.Windows;
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
    public partial class MainWindow
    {
        /// <summary>
        /// 数学输入面板的 COM 实例
        /// 存成字段原因有二：防 GC 提前回收（面板会消失）；重复点击按钮只显示同一个面板
        /// </summary>
        private Microsoft.MathInput.MathInputControlClass _mathInputPanel;

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

                //显示面板（已打开则不重复 Show）
                bool visible = false;
                _mathInputPanel.IsVisible(out visible);
                if (!visible)
                {
                    _mathInputPanel.Show();
                }
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
        }

        // ==================== 框选识别（右键"识别为函数"） ====================

        /// <summary>框选识别时暂存被识别的原始笔迹（Insert 成功后删掉它们，换成函数图像）</summary>
        private StrokeCollection _mathSourceStrokes;

        /// <summary>
        /// 框选笔迹后右键 → 识别为函数
        /// 入口：主画板右键菜单"识别为函数"（有选中笔迹时显示，见 EnsureMathContextMenu）
        /// 流程：选中笔迹 → 转成 COM Ink → LoadInk 喂给数学面板 → 老师在面板核对 →
        ///       点"插入"后原笔迹删除、原位置生成函数图像
        /// </summary>
        private void RecognizeSelectedAsFunction()
        {
            var selected = inkCanvas.GetSelectedStrokes();
            if (selected == null || selected.Count == 0) return;

            try
            {
                //先确保面板已创建并挂好事件（和点 fx 按钮是同一个面板）
                EnsureMathPanelCreated();
                bool visible = false;
                _mathInputPanel.IsVisible(out visible);
                if (!visible) _mathInputPanel.Show();

                //把 WPF 笔迹转成 COM Ink 对象（LoadInk 吃这个类型）
                var comInk = BuildComInk(selected);
                if (comInk == null)
                {
                    MessageBox.Show("笔迹转换失败，无法识别。", "提示",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                //记住这次识别的来源笔迹——Insert 成功后删它们换图像
                _mathSourceStrokes = new StrokeCollection();
                foreach (Stroke s in selected) _mathSourceStrokes.Add(s);

                //核心一步：喂给微软数学识别引擎，面板里直接显示识别结果预览
                _mathInputPanel.LoadInk(comInk);
            }
            catch (Exception ex)
            {
                MessageBox.Show("识别启动失败：" + ex.Message, "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                LogHelper.WriteLogToFile("[函数识别] 框选识别异常 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>
        /// WPF StrokeCollection → COM IInkDisp（数学面板 LoadInk 需要的类型）
        ///
        /// 转换采用"ISF 字节流"当桥，原因（原理说明）：
        /// 1. Microsoft.Ink.Ink 是微软的"托管包装类"（.NET 类库），不是 COM 对象本体——
        ///    直接强转 COM 接口 IInkDisp 会报 InvalidCastException（同名不同门）
        /// 2. 两个世界有共同的交换格式：ISF（Ink Serialized Format，墨迹序列化字节流）
        ///    托管 Ink 能 Save 成 ISF 字节，COM InkDisp 能 Load 这些字节
        /// 3. COM InkDisp 本体通过注册表 CLSID 直接创建（ProgID: InkObjCore.msinkaut.InkObject）
        ///
        /// 链路：WPF笔迹 → 托管Ink（逐点复制）→ Save出ISF字节 → COM InkDisp.Load → 完成
        /// </summary>
        private Microsoft.MathInput.IInkDisp BuildComInk(StrokeCollection strokes)
        {
            try
            {
                //第一步：WPF 笔迹 → 托管 Ink（逐条笔迹逐点复制坐标）
                var managedInk = new Microsoft.Ink.Ink();
                foreach (Stroke s in strokes)
                {
                    var pts = s.StylusPoints;
                    var mpts = new System.Drawing.Point[pts.Count];
                    for (int i = 0; i < pts.Count; i++)
                    {
                        mpts[i] = new System.Drawing.Point(
                            (int)Math.Round(pts[i].X), (int)Math.Round(pts[i].Y));
                    }
                    managedInk.CreateStroke(mpts);
                }
                if (managedInk.Strokes.Count == 0) return null;

                //第二步：托管 Ink → ISF 字节流（两种 Ink 世界的共同语言）
                byte[] isfBytes = managedInk.Save(
                    Microsoft.Ink.PersistenceFormat.InkSerializedFormat);

                //第三步：直接创建 COM InkDisp 本体（CLSID 来自注册表 InkObjCore.msinkaut.InkObject）
                Type inkType = Type.GetTypeFromCLSID(
                    new Guid("082C78E1-CD8F-4982-BEB9-BBFE43A0F09A"));
                if (inkType == null) return null;
                object comObj = Activator.CreateInstance(inkType);

                //第四步：把 ISF 字节灌进 COM 对象，再转成数学面板认识的接口类型
                var inkDisp = (Microsoft.MathInput.IInkDisp)comObj;
                inkDisp.Load(isfBytes);
                return inkDisp;
            }
            catch
            {
                return null; //任何转换问题都当"转换失败"处理，不炸主程序
            }
        }

        /// <summary>
        /// 框选识别的 Insert 后处理：删原笔迹 → 在原位置画函数图像
        /// （复用 OnMathInserted 的画图链路，只是位置改成"原来笔迹的地方"）
        /// </summary>
        private void OnMathInsertedFromSelection(string mathml)
        {
            try
            {
                //记下原笔迹的位置和大小（删之前）
                Rect srcBounds = _mathSourceStrokes != null && _mathSourceStrokes.Count > 0
                    ? _mathSourceStrokes.GetBounds() : Rect.Empty;

                //画图尺寸用原笔迹区域（识别的内容画在它原来的地方）
                double w = srcBounds.Width * 1.6;  //稍放大，坐标系需要留边
                double h = srcBounds.Height * 1.6;
                if (w < 200) w = 200; if (h < 150) h = 150;

                bool ok = MathGraph.MathGraphModule.TryBuildGraph(mathml, w, h,
                    out StrokeCollection graphStrokes, out string message);
                if (!ok)
                {
                    MessageBox.Show(message, "暂不能画图", MessageBoxButton.OK, MessageBoxImage.Information);
                    return; //识别失败：原笔迹保留不动（老师可以擦掉重写）
                }

                //图像平移到原笔迹中心
                var graphBounds = graphStrokes.GetBounds();
                double cx = srcBounds.Left + srcBounds.Width / 2;
                double cy = srcBounds.Top + srcBounds.Height / 2;
                var matrix = new System.Windows.Media.Matrix();
                matrix.Translate(cx - (graphBounds.Left + graphBounds.Width / 2),
                                 cy - (graphBounds.Top + graphBounds.Height / 2));
                graphStrokes.Transform(matrix, false);

                //替换：删原笔迹，放函数图像
                if (_mathSourceStrokes != null && _mathSourceStrokes.Count > 0)
                {
                    inkCanvas.Strokes.Remove(_mathSourceStrokes);
                    _mathSourceStrokes = null;
                }
                inkCanvas.Strokes.Add(graphStrokes);
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成图像失败：" + ex.Message, "提示",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                LogHelper.WriteLogToFile("[函数识别] 框选出图异常 " + ex, LogHelper.LogType.Error);
            }
        }

        /// <summary>确保数学面板已创建（点 fx 和框选识别共用同一个面板实例）</summary>
        private void EnsureMathPanelCreated()
        {
            if (_mathInputPanel != null) return;

            _mathInputPanel = new Microsoft.MathInput.MathInputControlClass();

            //设属主窗口：Windows 自动保证面板在主窗口之上（置顶问题的根治方案）
            var hwndSource = (HwndSource)PresentationSource.FromVisual(this);
            _mathInputPanel.SetOwnerWindow(hwndSource.Handle.ToInt64());

            //Insert 事件：分两种来源——面板直接写（fx）和框选识别，各自替换不同的内容
            _mathInputPanel.Insert += (mathml) =>
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    CloseMathPanel();
                    if (_mathSourceStrokes != null && _mathSourceStrokes.Count > 0)
                        OnMathInsertedFromSelection(mathml);  //框选来源：删原笔迹、原位出图
                    else
                        OnMathInserted(mathml);                //面板来源：画布中央出图
                }));
            };

            //Close 事件：老师点 X 放弃。框选来源时原笔迹保留（不删除）
            _mathInputPanel.Close += () =>
            {
                Dispatcher.BeginInvoke((Action)(() =>
                {
                    CloseMathPanel();
                    _mathSourceStrokes = null; //放弃识别，来源笔迹原地不动
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
                //画布当前可视区域（图像画在中间区域）
                double w = inkCanvas.ActualWidth, h = inkCanvas.ActualHeight;

                //完整链路：MathML → 解析 → 采样 → 坐标系+曲线笔迹
                bool ok = MathGraph.MathGraphModule.TryBuildGraph(mathml, w, h,
                    out StrokeCollection graphStrokes, out string message);

                if (!ok)
                {
                    //解析失败（含参、下标等暂不支持的情况）：提示原因，画板不动
                    MessageBox.Show(message, "暂不能画图", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                //图像平移到画布中央偏上位置（GraphBuilder 生成的坐标原点在图像中心）
                var graphBounds = graphStrokes.GetBounds();
                double dx = w / 2 - (graphBounds.Left + graphBounds.Width / 2);
                double dy = h * 0.4 - (graphBounds.Top + graphBounds.Height / 2);
                var matrix = new System.Windows.Media.Matrix();
                matrix.Translate(dx, dy);
                graphStrokes.Transform(matrix, false);

                //插入图像
                //（注：主画板的"选中"是自绘交互层，不是 Stroke.IsSelected 属性——
                // 插入后自动选中属于体验优化，后续接 SelectionGestures 时再加）
                inkCanvas.Strokes.Add(graphStrokes);
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
