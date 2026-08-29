using System;
using System.Windows.Ink;
using System.Windows.Media;

namespace Ink_Canvas.MathGraph
{
    /// <summary>
    /// 函数图像 → 画板笔迹（Stroke）生成器
    ///
    /// 输入：一个"给 x 算 y"的函数 + 画布像素尺寸
    /// 输出：一组笔迹（灰色坐标系 + 蓝色函数曲线），可直接加进 InkCanvas
    ///
    /// 设计要点：
    /// - 坐标系刻度取整数（1 单位固定像素宽），保证图像上的点能和刻度对上
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
        public static StrokeCollection BuildGraphStrokes(Func<double, double> f,
            double pixelWidth, double pixelHeight, bool smooth = true)
        {
            //画布尺寸异常时用兜底值（比如窗口还没渲染完就调用了）
            if (pixelWidth < 10) pixelWidth = 800;
            if (pixelHeight < 10) pixelHeight = 300;

            var result = new StrokeCollection();

            //坐标系和曲线用不同的笔迹样式
            var axisAttrs = new DrawingAttributes
            {
                Color = Color.FromRgb(0x99, 0x99, 0x99), //灰色
                Width = 1.5,
                Height = 1.5,
                FitToCurve = false,
                IsHighlighter = false
            };
            var curveAttrs = new DrawingAttributes
            {
                Color = Color.FromRgb(0x00, 0x66, 0xBF), //主题蓝
                Width = 2.5,
                Height = 2.5,
                //平滑开关：FitToCurve=true 时 WPF 会在采样点之间做贝塞尔过渡，
                //|x| 的 V 字尖角会被"拽弯"，所以折线类函数要关掉平滑
                FitToCurve = smooth
            };

            //第一步：坐标系（横轴、竖轴、刻度）
            AddAxes(result, pixelWidth, pixelHeight, axisAttrs);

            //第二步：函数曲线（可能被渐近线分成好几段）
            AddCurve(result, f, pixelWidth, pixelHeight, curveAttrs);

            return result;
        }

        /// <summary>世界坐标 → 画布像素坐标（原点在画布中心）</summary>
        private static double WorldToPixelX(double x, double w) { return w / 2 + x * PixelsPerUnit; }
        private static double WorldToPixelY(double y, double h) { return h / 2 - y * PixelsPerUnit; } //y 轴朝上，像素朝下，所以要翻转

        /// <summary>生成坐标轴和整刻度小线段</summary>
        private static void AddAxes(StrokeCollection result, double w, double h, DrawingAttributes attrs)
        {
            double cx = w / 2, cy = h / 2; //原点在画布中心

            //横轴：从左到右一条水平线
            result.Add(MakeStroke(new[]
            {
                new System.Windows.Input.StylusPoint(0, cy),
                new System.Windows.Input.StylusPoint(w, cy)
            }, attrs));

            //竖轴：从上到下一条垂直线
            result.Add(MakeStroke(new[]
            {
                new System.Windows.Input.StylusPoint(cx, 0),
                new System.Windows.Input.StylusPoint(cx, h)
            }, attrs));

            //刻度：每 1 单位一个小线段，轴两侧各出 4 像素
            int maxUnitX = (int)(w / 2 / PixelsPerUnit);
            for (int i = -maxUnitX; i <= maxUnitX; i++)
            {
                if (i == 0) continue; //原点不用画刻度
                double px = WorldToPixelX(i, w);
                result.Add(MakeStroke(new[]
                {
                    new System.Windows.Input.StylusPoint(px, cy - 4),
                    new System.Windows.Input.StylusPoint(px, cy + 4)
                }, attrs));
            }
            int maxUnitY = (int)(h / 2 / PixelsPerUnit);
            for (int i = -maxUnitY; i <= maxUnitY; i++)
            {
                if (i == 0) continue;
                double py = WorldToPixelY(i, h);
                result.Add(MakeStroke(new[]
                {
                    new System.Windows.Input.StylusPoint(cx - 4, py),
                    new System.Windows.Input.StylusPoint(cx + 4, py)
                }, attrs));
            }
        }

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
