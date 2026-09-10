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
[assembly: AssemblyVersion("6.2.2.0")]
[assembly: AssemblyFileVersion("6.2.2026.0910")]
[assembly: AssemblyInformationalVersion("6.2.2")]
