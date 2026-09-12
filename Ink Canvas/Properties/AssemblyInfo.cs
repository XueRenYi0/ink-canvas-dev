using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows;

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Inkboard · 板书白板")]
[assembly: AssemblyDescription("适用于课堂 PPT 放映和演示场景的板书白板：支持 PPT 批注、自由书写、选择缩放旋转、自定义图形图库、快捷键、笔记滚动、矩形橡皮擦等。")]
[assembly: AssemblyConfiguration("Release")]
[assembly: AssemblyCompany("Inkboard Team")]
[assembly: AssemblyProduct("Inkboard")]
[assembly: AssemblyCopyright("Copyright © Inkboard (formerly Ink Canvas) 2023–2026")]
[assembly: AssemblyTrademark("Inkboard")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

//In order to begin building localizable applications, set
//<UICulture>CultureYouAreCodingWith</UICulture> in your .csproj file
//inside a <PropertyGroup>.  For example, if you are using US english
//in your source files, set the <UICulture> to en-US.  Then uncomment
//the NeutralResourceLanguage attribute below.  Update the "en-US" in
//the line below to match the UICulture setting in the project file.

//[assembly: NeutralResourcesLanguage("en-US", UltimateResourceFallbackLocation.Satellite)]


[assembly: ThemeInfo(
    ResourceDictionaryLocation.None, //where theme specific resource dictionaries are located
                                     //(used if a resource is not found in the page,
                                     // or application resource dictionaries)
    ResourceDictionaryLocation.SourceAssembly //where the generic resource dictionary is located
                                              //(used if a resource is not found in the page,
                                              // app, or any theme specific resource dictionaries)
)]


// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]
// 版本：主版本.次版本.修订(YYYYMMDD).编译号 — 6.0.0：更名 Inkboard（原 Ink Canvas）、
// MathGraph 函数绘图模块（MathML 解析 + 采样绘制 + sin 快捷按钮）、数学公式识别面板、
// 停顿拉直（两条线修复）、选中框旋转手柄、框外直接书写、图形面板位置/图库高度优化、
// 撤销/快捷键统一走 TimeMachine
// 6.0.1：手写板/触摸输入稳定性（停顿拉直修复、矩形触摸橡皮、安全重启系统）、双指滚动、防误触
// 6.0.2：截图三件套（遮罩框选/掀板去墨/本地图片）、笔图标+色点两排三列+粗细面板、清屏优化、Fluent 图标
// 6.1.0：图片操作与墨迹统一（操作条/缩放手柄/翻转/复制模式/适配宽度/插入即选中）、
//        操作条精简为 10 键、双色色点、ISF 剪贴板、笔/橡皮/选择方式独立面板组件
// v6.1.1：自动更新接入国内镜像回退链（gh-proxy → at9 → 直连）、发版校验脚本修复
// v6.2.2：笔设置统一收进笔面板（设置页隐藏粗细/光标/笔锋/橡皮模式/橡皮大小/墨迹识别六项，
//         笔面板新增「墨迹识别 / 显示光标」开关区）、悬浮栏双指手势锁迁入快捷设置弹层、
//         识别出的图形粗细统一（去掉硬编码假笔锋，与圆/椭圆表现一致）、
//         墨迹识别改为「最少笔优先」，减少多笔误合并导致的误判
// v6.2.3：「更多」面板改版（条目加汉字标签、新增「系统」组、检查更新换 Download 图标、去掉恢复默认快捷键）、
//         主栏精简（设置/退出/白板底纹/多人书写收进更多面板）、快捷键对齐 ClassIn
//         （Ctrl+P 笔 / Ctrl+E 橡皮 / Ctrl+M 选择 / 新增 Ctrl+A 全选）、
//         滚动修复（滚动位置按页记忆、滚动后选区跟随、选区内滚轮转发）、
//         滚动胶囊改为按需浮现（去外框、窄长比例、静止淡出、无内容隐藏）、
//         检查更新全程有反馈（正在检查 / 已是最新 / 失败）
// v6.3.3：快捷键切换逻辑修复 —— Ctrl+E 橡皮改为幂等（原来按一下切橡皮、再按又切回画笔）；
//         Ctrl+M 选择改走与"选择墨迹"图标同一条路径（原来实际是隐藏画布，上游遗留的错误实现）；
//         Ctrl+A 全选可直接触发（原来必须先点过选择图标）；笔/橡皮/选择的提示文本同步更新；
//         检查更新失败时显示具体原因并写入日志，便于定位
// v6.5.5：清理与体验整理版 ——
//         ① 上游痕迹清理：移除 ink.wxriw.cn 后台版本上报、删除「马鞍山二中专版」UI 与 4 张图（340KB）、
//            删除首次运行向导 WelcomeWindow（其唯一调用点就是被删的联网块，删后成孤儿）
//         ② 悬浮栏观感：弹层底色透明度规范到 85%（浅色/深色），工具高亮色统一取 InkSpec 的 HighlightBrush
//         ③ 启动优化：主题资源字典改为「幂等 + 就地替换」（启动字典数 12→8，不再随系统主题无限增长）；
//            新增启动耗时分段剖析日志（便于后续定位）
//         ④ 稳定性：工具栏防误触守卫补全 37 处（修掉"按下与松开不在同一元素也误触发"）；
//            弹层行为统一（新增统一关闭入口 + 点外收起登记表，「更多」面板现在点外即收）
// v6.6.1：颜色与面板定位修正版 ——
//         ① 修「点检查更新没反应」：更新弹窗里选过"延迟 30 分钟"（或"跳过此版本"）之后，
//            再点检查更新完全静默 —— AutoUpdater.NET 把延迟时间点存在注册表 RemindLaterAt，
//            并在内存挂了 _remindLaterTimer，两者都会让 Start() 连请求都不发就返回，
//            且不触发结果回调（UI 反馈全在回调里）。现在手动检查会先清掉这两处拦截
//         ② 快捷换色条改版：4 格 × 2 色 = 固定八色（黑/白 · 红/青 · 蓝/黄 · 绿/品红），
//            点哪个用哪个（原来是"单击在对色间跳转"）；大小色区各自独立选中环
//         ③ 画笔四色不再随白板/黑板切两套色（点同一格拿到的颜色会悄悄变）——
//            统一一套出厂色，仍可用 Colors\Colors.ini 覆盖；进板面也不再强制改成红笔
//         ④ 悬浮栏「更多」面板解挂到根层、改用绝对坐标定位（原来吊在缩放容器内，
//            只能靠负 Margin 表达，面板偏右且底边扎出工具条）；内容密集故单开
//            高不透明度背景键 FloatBarBackgroundOpaque（α≈97%），消除文字与背后窗口重影
//         ⑤ 激光笔不参与停顿拉直：拉直生成的直线不经 StrokeCollected，激光淡出永不启动，
//            会永久留在画布上且进不了撤销栈
[assembly: AssemblyVersion("6.6.1.0")]
[assembly: AssemblyFileVersion("6.6.2026.0912")]
[assembly: AssemblyInformationalVersion("6.6.1")]
