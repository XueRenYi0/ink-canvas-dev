using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace Ink_Canvas.Helpers
{
    /// <summary>
    /// 启动耗时诊断器（2026-09-11 加）。
    /// 目的：优化启动速度时**先定位再动手**——此前误判"资源字典重复加载"是启动慢的主因，
    /// 实测被否证（WPF 对同 Source 字典有缓存，重复 Add 不重新解析），教训就是不猜。
    ///
    /// 用法：App 构造函数里 Init()（对齐"进程启动"零点），
    /// 需要计时的代码段用 using (StartupProfiler.Measure("名字")) { ... } 包住，
    /// 关键时刻用 Mark("名字") 记瞬时点，最后 Dump("标题") 一次性汇总进日志。
    ///
    /// 开销极小（围绕已有调用的 Begin/End），可常驻正式版：每次启动只多十几行日志。
    /// 不想要了：把最后一个 Dump() 调用注释掉即可（其余打点会自动早退，零副作用）。
    /// </summary>
    internal static class StartupProfiler
    {
        private static readonly Stopwatch Clock = Stopwatch.StartNew();
        private static DateTime _procStart = DateTime.Now;
        private static readonly List<KeyValuePair<string, double>> _marks =
            new List<KeyValuePair<string, double>>();
        private static readonly List<KeyValuePair<string, double>> _sections =
            new List<KeyValuePair<string, double>>();
        private static string _openName;
        private static double _openAt;
        private static bool _dumped;

        /// <summary>最早调用（App 构造函数）：取进程启动时刻作零点</summary>
        public static void Init()
        {
            try { _procStart = Process.GetCurrentProcess().StartTime; }
            catch { _procStart = DateTime.Now; }
            _marks.Clear();
            _sections.Clear();
            _openName = null;
            _dumped = false;
        }

        /// <summary>距进程启动的毫秒数（系统时间被改动时退回单调秒表读数）</summary>
        private static double ElapsedMs
        {
            get
            {
                double v;
                try { v = (DateTime.Now - _procStart).TotalMilliseconds; }
                catch { v = Clock.Elapsed.TotalMilliseconds; }
                return v >= 0 ? v : Clock.Elapsed.TotalMilliseconds;
            }
        }

        /// <summary>记一个瞬时时刻（如"窗口句柄就绪"）</summary>
        public static void Mark(string name)
        {
            if (_dumped) return;
            try { _marks.Add(new KeyValuePair<string, double>(name, ElapsedMs)); }
            catch { }
        }

        /// <summary>开始一段耗时统计</summary>
        public static void Begin(string name)
        {
            if (_dumped) return;
            _openName = name;
            _openAt = ElapsedMs;
        }

        /// <summary>结束当前耗时段</summary>
        public static void End()
        {
            if (_dumped || _openName == null) return;
            try { _sections.Add(new KeyValuePair<string, double>(_openName, ElapsedMs - _openAt)); }
            catch { }
            _openName = null;
        }

        /// <summary>包住一段代码：using (StartupProfiler.Measure("X")) { ... }</summary>
        public static IDisposable Measure(string name)
        {
            Begin(name);
            return new Closer();
        }

        private sealed class Closer : IDisposable
        {
            public void Dispose() { End(); }
        }

        /// <summary>汇总输出（只输一次）。放在"界面就绪"之后调用。</summary>
        public static void Dump(string title)
        {
            if (_dumped) return;
            _dumped = true;
            try
            {
                var sb = new StringBuilder();
                sb.AppendLine($"[Startup] ===== {title}：距进程启动 {ElapsedMs:F0} ms =====");
                sb.AppendLine("[Startup] --- 关键时刻 ---");
                foreach (var m in _marks)
                    sb.AppendLine($"[Startup]   {m.Value,8:F0} ms   {m.Key}");
                sb.AppendLine("[Startup] --- 各段耗时（按顺序）---");
                double accounted = 0;
                foreach (var s in _sections)
                {
                    accounted += s.Value;
                    sb.AppendLine($"[Startup]   {s.Value,8:F0} ms   {s.Key}");
                }
                sb.AppendLine($"[Startup]   {accounted,8:F0} ms   == 上面各段合计 ==");
                LogHelper.WriteLogToFile(sb.ToString().TrimEnd(), LogHelper.LogType.Event);
            }
            catch { }
        }
    }
}
