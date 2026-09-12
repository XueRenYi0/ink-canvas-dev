using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Ink;
using System.Windows.Input;
using Ink_Canvas.Helpers;

namespace Ink_Canvas
{
    /// <summary>
    /// 保角平滑（2026-09-12）：对手写笔迹做"转角锚点 + 段内平滑"。
    ///
    /// 背景：WPF 的 FitToCurve（贝塞尔曲线拟合）会把识别出的图形和汉字的直角"一刀切"磨圆。
    /// 本模块替代它 —— 转折（方折、竖钩、图框角）保留棱角，直线与曲线照常顺滑。
    /// 相比 quintic 贝塞尔平滑：**不增点**（点数不变）、**零相位**（抬笔后可用居中滤波，无滞后）、**可控**。
    ///
    /// 调用点：inkCanvas_StrokeCollected（抬笔后、笔锋计算之前），仅作用于手写笔迹
    /// （程序化图形由 NormalizeAttributesForShapeMode 另行处理，不受影响）。
    /// 只改 StylusPoints（与笔锋改写同一类事件，不触碰 DrawingAttributes，不影响撤销历史）。
    /// </summary>
    public partial class MainWindow
    {
        /// <summary>多小的角算"折"（宏观转角超过此值即标为锚点）。调小 → 更多笔画被当成折而保留；调大 → 更多被平滑。</summary>
        private const double PreserveCornerAngleDeg = 40.0;
        /// <summary>测转角时"前后各取多远"（px，半径；会按实际点距换算成点数）。
        /// 为什么不直接写"取几个点"：手写板一点约 2~5px，而鼠标/触控板合并后可能一点 90px，
        /// 同一个"5 个点"在两种输入下对应的实际距离差 20 倍以上 —— 点稀的会被过度平滑/漏判转折。</summary>
        private const double PreserveCornerMacroPx = 24.0;
        /// <summary>段内平滑窗口的**总宽度**（px；会按点距换算成点数，即 ±宽度/2）。
        /// 20px 的依据：汉字书写能量主峰约 5Hz、噪声在 10Hz 以上（Teulings &amp; Maarse 1984）,
        /// 按常见书写速度折算，要滤掉 10Hz 以上的抖动，窗口量级就是十几到二十几像素。</summary>
        private const double PreserveCornerSmoothPx = 20.0;

        /// <summary>只打一次"首笔自检"日志：用来区分"没生效"和"生效了但看不出差别"。</summary>
        private bool _preserveCornerLogged = false;

        /// <summary>
        /// 对手写笔迹做保角平滑：锚点（转折/两端）不动，锚点之间居中平滑。直接改写 stroke 的 StylusPoints。
        /// </summary>
        private void ApplyPreserveCornerSmoothing(Stroke stroke)
        {
            StylusPointCollection sp = stroke.StylusPoints;
            int n = sp.Count;
            if (n < 6) return;   // 太短（点一下、小标点）不平滑，保持原样

            var pts = new Point[n];
            for (int i = 0; i < n; i++) pts[i] = new Point(sp[i].X, sp[i].Y);

            int anchors, mw, sw; double spacing;
            Point[] sm = PreserveCornerSmoothCore(pts, out anchors, out mw, out sw, out spacing);

            // 启动日志只能证明"开关是开的"，这行才能证明**真的处理到笔迹**了。
            // 平滑效果偏主观，有这行才能区分"没生效"和"生效了但看不出差别"。只打首笔一次，不刷屏。
            if (!_preserveCornerLogged)
            {
                _preserveCornerLogged = true;
                LogHelper.WriteLogToFile(
                    $"[Canvas] 保角平滑首笔：{n} 点 / 平均点距 {spacing:F1}px → 锚点 {anchors} 个"
                    + $"（宏窗口 {mw} 点 ≈ {mw * spacing:F0}px，平滑窗口 {sw} 点 ≈ {sw * spacing:F0}px，阈值 {PreserveCornerAngleDeg}°）",
                    LogHelper.LogType.Event);
            }

            var outp = new StylusPointCollection();
            for (int i = 0; i < n; i++)
                outp.Add(new StylusPoint(sm[i].X, sm[i].Y, sp[i].PressureFactor));   // 压力原样保留
            stroke.StylusPoints = outp;
        }

        private Point[] PreserveCornerSmoothCore(Point[] pts, out int anchorCount,
            out int macroWin, out int smoothWin, out double meanSpacing)
        {
            int n = pts.Length;
            anchorCount = 0; macroWin = 3; smoothWin = 5; meanSpacing = 1.0;
            if (n < 3) return (Point[])pts.Clone();

            // 0) 平均点距 → 把"像素窗口"换算成"点数窗口"。
            //    这是为了让手写板（一点 2~5px）与鼠标/触控板（合并后可能一点 90px）
            //    落在同一个判定尺度上：窗口写死点数的话，点稀的输入会被过度平滑、
            //    且宏观转角会横跨整段直线而漏判转折。
            double len = 0;
            for (int i = 1; i < n; i++) len += PreserveCornerDist(pts[i - 1], pts[i]);
            meanSpacing = len / (n - 1);
            if (meanSpacing < 1e-6) meanSpacing = 1.0;
            int mw = PreserveCornerClamp((int)Math.Round(PreserveCornerMacroPx / meanSpacing), 2, 12);
            int sw = PreserveCornerClamp((int)Math.Round(PreserveCornerSmoothPx / meanSpacing), 2, 9);
            macroWin = mw; smoothWin = sw;

            bool[] anchor = new bool[n];
            anchor[0] = true; anchor[n - 1] = true;

            // 1) 宏观转角检测（宏窗口抗抖），超过阈值的标为锚点候选；相邻候选聚类成一个（取转角最大的）
            var cand = new List<int>();
            for (int i = mw; i < n - mw; i++)
            {
                if (PreserveCornerTurnAngle(pts[i - mw], pts[i], pts[i + mw]) >= PreserveCornerAngleDeg)
                    cand.Add(i);
            }
            foreach (int a in PreserveCornerClusterAnchors(cand, pts, mw)) anchor[a] = true;

            // 2) 锚点之间分段，段内居中滑动平均，锚点保持不动
            var result = (Point[])pts.Clone();
            int[] anchors = Enumerable.Range(0, n).Where(i => anchor[i]).OrderBy(i => i).ToArray();
            anchorCount = anchors.Length;
            for (int k = 0; k < anchors.Length - 1; k++)
            {
                int a = anchors[k], b = anchors[k + 1];
                if (b - a <= 2) continue;
                var seg = new Point[b - a + 1];
                for (int i = a; i <= b; i++) seg[i - a] = pts[i];
                Point[] segSm = PreserveCornerCenteredAvg(seg, sw);
                for (int i = a + 1; i < b; i++) result[i] = segSm[i - a];
            }
            return result;
        }

        /// <summary>net472 没有 Math.Clamp，这里自带一个最小的整数版。</summary>
        private static int PreserveCornerClamp(int v, int lo, int hi)
        { return v < lo ? lo : (v > hi ? hi : v); }

        private static double PreserveCornerDist(Point a, Point b)
        {
            double dx = a.X - b.X, dy = a.Y - b.Y;
            return Math.Sqrt(dx * dx + dy * dy);
        }

        /// <summary>相邻锚点候选聚类：把间距 ≤ mw 的候选合并成一个，取宏观转角最大的那个。</summary>
        private int[] PreserveCornerClusterAnchors(List<int> cand, Point[] pts, int mw)
        {
            var outp = new List<int>();
            int i = 0, n = pts.Length;
            while (i < cand.Count)
            {
                int j = i, best = cand[i]; double bestAng = 0;
                while (j < cand.Count && cand[j] - cand[i] <= mw)
                {
                    int a = cand[j];
                    double ang = PreserveCornerTurnAngle(
                        pts[Math.Max(0, a - mw)], pts[a], pts[Math.Min(n - 1, a + mw)]);
                    if (ang > bestAng) { bestAng = ang; best = a; }
                    j++;
                }
                outp.Add(best);
                i = j;
            }
            return outp.ToArray();
        }

        /// <summary>三点夹角（度）：b 处从 a→b 到 b→c 的转角。</summary>
        private double PreserveCornerTurnAngle(Point a, Point b, Point c)
        {
            var v1 = new Vector(b.X - a.X, b.Y - a.Y);
            var v2 = new Vector(c.X - b.X, c.Y - b.Y);
            if (v1.Length < 1e-9 || v2.Length < 1e-9) return 0;
            v1.Normalize(); v2.Normalize();
            return Math.Acos(Math.Max(-1.0, Math.Min(1.0, Vector.Multiply(v1, v2)))) * 180.0 / Math.PI;
        }

        /// <summary>居中滑动平均（窗口取奇数保证左右对称，即零相位滤波）。</summary>
        private Point[] PreserveCornerCenteredAvg(Point[] pts, int win)
        {
            int n = pts.Length;
            if (win % 2 == 0) win++;
            int half = win / 2;
            var outp = new Point[n];
            for (int i = 0; i < n; i++)
            {
                double sx = 0, sy = 0; int cnt = 0;
                for (int j = i - half; j <= i + half; j++)
                {
                    if (j < 0 || j >= n) continue;
                    sx += pts[j].X; sy += pts[j].Y; cnt++;
                }
                outp[i] = new Point(sx / cnt, sy / cnt);
            }
            return outp;
        }
    }
}
