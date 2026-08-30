using System;
using System.Windows.Ink;
using System.Windows.Media;

namespace Ink_Canvas.MathGraph
{
    /// <summary>
    /// 函数图像 → 画板笔迹（Stroke）生成器
    ///
    /// 输入：一个"给 x 算 y"的函数 + 画布像素尺寸
    /// 输出：一组笔迹（函数曲线 + 原点圆点标记），可直接加进 InkCanvas
    ///
    /// 设计要点：
    /// - 不自带坐标系（避免画板上出现多套坐标系），只在 (0,0) 处画一个圆点标记，
    ///   老师把圆点拖到自己坐标系的原点即可对齐；比例固定 1 单位 = 30 像素
    /// - 曲线遇到间断点（1/x、tan 的渐近线）自动断笔，不会画出竖线假象
    /// - 只依赖 WPF 基础类型，不依赖主窗口任何代码（保持模块独立）
    /// </summary>
    public static class GraphBuilder
    {
        /// <summary>1 个数学单位占多少像素（数值可按屏幕大小微调）</summary>
        private const double PixelsPerUnit = 30.0;

        /// <summary>曲线采样间隔（像素），越小越平滑、笔迹点越多</summary>
        private const double SampleStepPx = 2.0;

        /// <summary>
        /// 生成函数图像笔迹（带"是否平滑"选项）
        /// </summary>
        /// <param name="f">函数：给 x 返回 y</param>
        /// <param name="pixelWidth">目标画布宽度（像素）</param>
        /// <param name="pixelHeight">目标画布高度（像素）</param>
        /// <param name="smooth">true=曲线平滑（适合 sin 等弯曲函数）；false=直线连接（适合 |x| 等折线函数，平滑会把尖角拽弯）</param>
        /// <param name="penAttrs">画笔属性（颜色/粗细跟随主画板当前画笔）；不传则用模块默认的蓝色曲线</param>
        public static StrokeCollection BuildGraphStrokes(Func<double, double> f,
            double pixelWidth, double pixelHeight, bool smooth = true,
            DrawingAttributes penAttrs = null)
        {
            //画布尺寸异常时用兜底值（比如窗口还没渲染完就调用了）
            if (pixelWidth < 10) pixelWidth = 800;
            if (pixelHeight < 10) pixelHeight = 300;

            var result = new StrokeCollection();

            //曲线样式：优先跟随调用方传入的画笔属性（与图形面板"选中什么笔画出什么样"的大原则一致），
            //没传时用模块默认的主题蓝
            DrawingAttributes curveAttrs;
            if (penAttrs != null)
            {
                curveAttrs = penAttrs.Clone();
            }
            else
            {
                curveAttrs = new DrawingAttributes
                {
                    Color = Color.FromRgb(0x00, 0x66, 0xBF), //主题蓝
                    Width = 2.5,
                    Height = 2.5,
                    IsHighlighter = false
                };
            }
            //平滑开关始终由函数形状决定（|x| 的 V 字尖角不能被平滑拽弯），
            //所以这里统一覆盖，不照抄画笔里的 FitToCurve 设置
            curveAttrs.FitToCurve = smooth;

            //第一步：函数曲线（可能被渐近线分成好几段）
            AddCurve(result, f, pixelWidth, pixelHeight, curveAttrs);

            //第二步：原点圆点标记（不画坐标系——老师黑板上有自己的坐标系，
            //把圆点拖到自己坐标系的 (0,0) 上即可对齐；比例 1 单位 = 30 像素）
            var dotAttrs = curveAttrs.Clone();
            dotAttrs.Width = curveAttrs.Width * 1.8; //比曲线稍粗，在曲线旁清晰可见
            dotAttrs.Height = curveAttrs.Height * 1.8;
            dotAttrs.FitToCurve = false; //单个点，无需平滑
            //注意：Stroke 构造时集合必须已含点（空集合连给 Stroke 会抛异常），
            //所以先把点装进集合、再构造
            var dotPoints = new System.Windows.Input.StylusPointCollection();
            dotPoints.Add(new System.Windows.Input.StylusPoint(pixelWidth / 2, pixelHeight / 2));
            var dot = new Stroke(dotPoints) { DrawingAttributes = dotAttrs };
            result.Add(dot);

            return result;
        }

        /// <summary>世界坐标 → 画布像素坐标（原点在画布中心）</summary>
        private static double WorldToPixelY(double y, double h) { return h / 2 - y * PixelsPerUnit; } //y 轴朝上，像素朝下，所以要翻转

        /// <summary>采样函数生成曲线笔迹（含间断点断段逻辑）</summary>
        private static void AddCurve(StrokeCollection result, Func<double, double> f,
            double w, double h, DrawingAttributes attrs)
        {
            //当前段的采样点集合（间断后要开新段）
            var points = new System.Windows.Input.StylusPointCollection();
            //上一段的 y 值（用来检测跳变）
            double lastY = double.NaN;

            //从左到右按像素步进采样
            for (double px = 0; px <= w; px += SampleStepPx)
            {
                double x = (px - w / 2) / PixelsPerUnit; //像素 → 世界坐标
                double y = double.NaN;
                try { y = f(x); } catch { /*求值抛异常按无效处理*/ }

                bool valid = !double.IsNaN(y) && !double.IsInfinity(y)
                             && Math.Abs(y) < (h / PixelsPerUnit); //超出画布高度的点丢弃

                if (valid)
                {
                    //跳变检测：相邻两点 y 差距过大说明中间有渐近线（tan、1/x），
                    //必须断开，否则会画出一条假的竖线
                    bool jump = !double.IsNaN(lastY) && Math.Abs(y - lastY) > 5;
                    if (jump && points.Count > 0)
                    {
                        FlushSegment(result, points, attrs);
                        points = new System.Windows.Input.StylusPointCollection();
                    }
                    points.Add(new System.Windows.Input.StylusPoint(px, WorldToPixelY(y, h)));
                    lastY = y;
                }
                else
                {
                    //无效点（NaN / 无穷 / 出画布）：结束当前段
                    if (points.Count > 0)
                    {
                        FlushSegment(result, points, attrs);
                        points = new System.Windows.Input.StylusPointCollection();
                    }
                    lastY = double.NaN;
                }
            }
            //收尾：最后一段别漏了
            FlushSegment(result, points, attrs);
        }

        /// <summary>把一段采样点真正变成笔迹加入集合（空段直接忽略）</summary>
        private static void FlushSegment(StrokeCollection result,
            System.Windows.Input.StylusPointCollection points, DrawingAttributes attrs)
        {
            if (points == null || points.Count < 2) return;
            result.Add(MakeStroke(points, attrs));
        }

        /// <summary>用给定点集和样式造一条笔迹</summary>
        private static Stroke MakeStroke(System.Windows.Input.StylusPointCollection points,
            DrawingAttributes attrs)
        {
            var s = new Stroke(points);
            s.DrawingAttributes = attrs.Clone(); //每条笔迹独立样式副本，互不影响
            return s;
        }

        /// <summary>重载：用数组造笔迹</summary>
        private static Stroke MakeStroke(System.Windows.Input.StylusPoint[] points,
            DrawingAttributes attrs)
        {
            return MakeStroke(new System.Windows.Input.StylusPointCollection(points), attrs);
        }
    }
}
