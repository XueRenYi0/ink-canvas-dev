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
[assembly: AssemblyVersion("6.0.0.0")]
[assembly: AssemblyFileVersion("6.0.2026.0830")]
[assembly: AssemblyInformationalVersion("6.0.0")]
