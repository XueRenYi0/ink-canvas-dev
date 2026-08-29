using System;
using System.Windows.Ink;

namespace Ink_Canvas.MathGraph
{
    /// <summary>
    /// 数学函数图像模块的唯一对外入口
    ///
    /// 设计原则（模块独立性）：
    /// 1. 整个 MathGraph 文件夹是自包含的：删除文件夹 = 功能彻底移除
    /// 2. 外部代码（测试窗口、未来的主画板集成）只允许调用这一个类
    /// 3. 不依赖 MainWindow、不依赖 TimeMachine、不依赖任何全局状态
    /// 4. 纯托管对象（表达式树、笔迹），窗口关闭即可被整体回收，
    ///    不持有非托管资源——占用大或出问题随时可以弃用，不留垃圾
    /// </summary>
    public static class MathGraphModule
    {
        /// <summary>
        /// 把 MathML 转成函数图像笔迹
        /// </summary>
        /// <param name="mathml">数学识别器输出的 MathML 字符串</param>
        /// <param name="pixelWidth">目标画布宽度</param>
        /// <param name="pixelHeight">目标画布高度</param>
        /// <param name="strokes">成功时输出笔迹集合（坐标系+曲线）</param>
        /// <param name="message">成功时是公式文本，失败时是中文原因</param>
        /// <returns>成功与否（失败不抛异常，方便调用方简单处理）</returns>
        public static bool TryBuildGraph(string mathml, double pixelWidth, double pixelHeight,
            out StrokeCollection strokes, out string message)
        {
            strokes = null;
            try
            {
                //第一段：MathML → 可求值的函数（同时拿到表达式里有没有绝对值的信息）
                var f = MathMLParser.Compile(mathml, out bool hasAbs);

                //第二段：采样函数 → 坐标系 + 曲线笔迹
                //含绝对值的函数（|x| 等）有关键的尖角，必须用直线连接，平滑会把角拽弯
                strokes = GraphBuilder.BuildGraphStrokes(f, pixelWidth, pixelHeight, smooth: !hasAbs);

                message = "已绘制：" + MathMLParser.ToPlainText(mathml);
                return true;
            }
            catch (FormatException fex)
            {
                //格式错误：公式里有暂不支持的记号（含参、下标等），属于预期内的拒绝
                message = "暂不能画图：" + fex.Message;
            }
            catch (Exception ex)
            {
                //其他意外错误：也要兜住，绝不能让模块问题拖垮调用方
                message = "生成图像时出错：" + ex.Message;
            }
            return false;
        }
    }
}
