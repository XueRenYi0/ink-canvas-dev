using Newtonsoft.Json;
using System.Collections.Generic;

namespace Ink_Canvas
{
    public class Settings
    {
        [JsonProperty("advanced")]
        public Advanced Advanced { get; set; } = new Advanced();
        [JsonProperty("appearance")]
        public Appearance Appearance { get; set; } = new Appearance();
        [JsonProperty("automation")]
        public Automation Automation { get; set; } = new Automation();
        [JsonProperty("behavior")]
        public PowerPointSettings PowerPointSettings { get; set; } = new PowerPointSettings();
        [JsonProperty("canvas")]
        public Canvas Canvas { get; set; } = new Canvas();
        [JsonProperty("gesture")]
        public Gesture Gesture { get; set; } = new Gesture();
        [JsonProperty("inkToShape")]
        public InkToShape InkToShape { get; set; } = new InkToShape();
        [JsonProperty("startup")]
        public Startup Startup { get; set; } = new Startup();
        [JsonProperty("shortcuts")]
        public ShortcutSettings Shortcuts { get; set; } = new ShortcutSettings();
        [JsonProperty("mathPanel")]
        public MathPanelSettings MathPanel { get; set; } = new MathPanelSettings();
    }

    /// <summary>
    /// 数学输入面板（手写函数识别）的窗口位置记忆。
    /// 用户拖到哪，下次打开就在哪（关闭/插入时保存坐标）；
    /// 首次启动没有记忆时用默认位置（主窗口所在屏中央偏上）。
    /// </summary>
    public class MathPanelSettings
    {
        [JsonProperty("hasPosition")]
        public bool HasPosition { get; set; } = false;
        [JsonProperty("x")]
        public double X { get; set; } = 0;
        [JsonProperty("y")]
        public double Y { get; set; } = 0;
    }

    /// <summary>
    /// 快捷键设置（窗口焦点时生效，用 WPF KeyBinding 动态注册，不走全局热键）。
    /// Key 是动作标识（如 "Pen"），Value 是按键串（如 "F2"、"Ctrl+Z"、"Alt+S"）。
    /// </summary>
    public class ShortcutSettings
    {
        [JsonProperty("bindings")]
        public Dictionary<string, string> Bindings { get; set; } = new Dictionary<string, string>();

        /// <summary>获取某动作的按键串，无配置时返回默认</summary>
        public string Get(string actionId, string defaultGesture)
        {
            if (Bindings != null && Bindings.TryGetValue(actionId, out string g) && !string.IsNullOrWhiteSpace(g))
                return g;
            return defaultGesture;
        }

        /// <summary>设置某动作的按键串</summary>
        public void Set(string actionId, string gesture)
        {
            if (Bindings == null) Bindings = new Dictionary<string, string>();
            if (string.IsNullOrWhiteSpace(gesture))
                Bindings.Remove(actionId);
            else
                Bindings[actionId] = gesture;
        }
    }

    public class Canvas
    {
        [JsonProperty("inkWidth")]
        public double InkWidth { get; set; } = 2.5;
        [JsonProperty("isShowCursor")]
        public bool IsShowCursor { get; set; } = true; //默认开（笔形光标）：教室大屏距离远，笔形比红点好找落点；偏好红点可在设置里关
        [JsonProperty("inkStyle")]
        public int InkStyle { get; set; } = 0;
        /// <summary>笔种类：0=画笔 1=荧光笔 2=激光笔（笔设置面板选择，重启恢复）</summary>
        [JsonProperty("penType")]
        public int PenType { get; set; } = 0;
        /// <summary>荧光笔宽度（与画笔各自独立记忆，重启恢复）</summary>
        [JsonProperty("highlighterWidth")]
        public double HighlighterWidth { get; set; } = 14;
        /// <summary>笔面板自定义颜色（#AARRGGBB 十六进制串；null=未设置，面板显示"＋"入口）</summary>
        [JsonProperty("customColor")]
        public string CustomColor { get; set; } = null;
        /// <summary>选择墨迹方式：0=矩形框选（默认） 1=自由选择（套索）（选择面板选择，重启恢复）</summary>
        [JsonProperty("selectionMode")]
        public int SelectionMode { get; set; } = 0;
        [JsonProperty("eraserSize")]
        public int EraserSize { get; set; } = 2;
        [JsonProperty("eraserType")]
        public int EraserType { get; set; } = 0; // 0 - 图标切换模式      1 - 面积擦     2 - 线条擦
        [JsonProperty("hideStrokeWhenSelecting")]
        public bool HideStrokeWhenSelecting { get; set; } = true;

        [JsonProperty("usingWhiteboard")]
        // 默认开：黑板模式以白色板面启动（用户需求：初始即白板）。已保存过设置的用户不受影响
        public bool UsingWhiteboard { get; set; } = true;

        [JsonProperty("whiteboardPattern")]
        public int WhiteboardPattern { get; set; } = 0; // 白板底纹：0=无 1=方格 2=横线（即改即存，下次启动继承）
        [JsonProperty("whiteboardGridSize")]
        public double WhiteboardGridSize { get; set; } = 40; // 底纹间距（px）：方格边长/横线行距共用

        [JsonProperty("hyperbolaAsymptoteOption")]
        public OptionalOperation HyperbolaAsymptoteOption { get; set; } = OptionalOperation.Ask;
    }

    public enum OptionalOperation
    {
        Yes,
        No,
        Ask
    }

    public class Gesture
    {
        [JsonIgnore]
        public bool IsEnableTwoFingerGesture => IsEnableTwoFingerZoom || IsEnableTwoFingerTranslate || IsEnableTwoFingerRotation;
        [JsonProperty("isDisableLockSmithByDefault")]
        public bool IsDisableLockSmithByDefault { get; set; } = true;
        [JsonProperty("isEnableTwoFingerZoom")]
        public bool IsEnableTwoFingerZoom { get; set; } = true;
        [JsonProperty("isEnableTwoFingerTranslate")]
        public bool IsEnableTwoFingerTranslate { get; set; } = true;
        [JsonProperty("isEnableTwoFingerRotation")]
        public bool IsEnableTwoFingerRotation { get; set; } = false;
        [JsonProperty("isEnableTwoFingerRotationOnSelection")]
        public bool IsEnableTwoFingerRotationOnSelection { get; set; } = false;

    }

    public class Startup
    {
        [JsonProperty("isAutoHideCanvas")]
        public bool IsAutoHideCanvas { get; set; } = true;
        [JsonProperty("isAutoEnterModeFinger")]
        public bool IsAutoEnterModeFinger { get; set; } = false;
        /// <summary>启动时自动检查更新（默认开；关闭后仍可从悬浮条右键菜单手动检查）</summary>
        [JsonProperty("isAutoCheckUpdate")]
        public bool IsAutoCheckUpdate { get; set; } = true;
    }

    public class Appearance
    {
        [JsonProperty("isTransparentButtonBackground")]
        public bool IsTransparentButtonBackground { get; set; } = true;
        [JsonProperty("isShowExitButton")]
        public bool IsShowExitButton { get; set; } = true;
        [JsonProperty("isShowEraserButton")]
        public bool IsShowEraserButton { get; set; } = true;
        [JsonProperty("isShowHideControlButton")]
        public bool IsShowHideControlButton { get; set; } = false;
        [JsonProperty("isShowLRSwitchButton")]
        public bool IsShowLRSwitchButton { get; set; } = false;
        [JsonProperty("isShowModeFingerToggleSwitch")]
        public bool IsShowModeFingerToggleSwitch { get; set; } = true;
        [JsonProperty("theme")]
        public int Theme { get; set; } = 0;
    }

    public class PowerPointSettings
    {
        [JsonProperty("isShowPPTNavigation")]
        public bool IsShowPPTNavigation { get; set; } = true;
        [JsonProperty("powerPointSupport")]
        public bool PowerPointSupport { get; set; } = true;
        [JsonProperty("isShowCanvasAtNewSlideShow")]
        public bool IsShowCanvasAtNewSlideShow { get; set; } = true;
        [JsonProperty("isNoClearStrokeOnSelectWhenInPowerPoint")]
        public bool IsNoClearStrokeOnSelectWhenInPowerPoint { get; set; } = true;
        [JsonProperty("isShowStrokeOnSelectInPowerPoint")]
        public bool IsShowStrokeOnSelectInPowerPoint { get; set; } = false;
        [JsonProperty("isAutoSaveStrokesInPowerPoint")]
        public bool IsAutoSaveStrokesInPowerPoint { get; set; } = true;
        [JsonProperty("isAutoSaveScreenShotInPowerPoint")]
        public bool IsAutoSaveScreenShotInPowerPoint { get; set; } = false;
        [JsonProperty("isNotifyPreviousPage")]
        public bool IsNotifyPreviousPage { get; set; } = false;
        [JsonProperty("isNotifyHiddenPage")]
        public bool IsNotifyHiddenPage { get; set; } = true;
        [JsonProperty("isEnableTwoFingerGestureInPresentationMode")]
        public bool IsEnableTwoFingerGestureInPresentationMode { get; set; } = false;
        [JsonProperty("isEnableFingerGestureSlideShowControl")]
        public bool IsEnableFingerGestureSlideShowControl { get; set; } = true;
        [JsonProperty("isSupportWPS")]
        public bool IsSupportWPS { get; set; } = true;
    }

    public class Automation
    {
        [JsonProperty("isAutoKillPptService")]
        public bool IsAutoKillPptService { get; set; } = false;

        [JsonProperty("isAutoKillEasiNote")]
        public bool IsAutoKillEasiNote { get; set; } = false;

        [JsonProperty("isSaveScreenshotsInDateFolders")]
        public bool IsSaveScreenshotsInDateFolders { get; set; } = false;

        [JsonProperty("isAutoSaveStrokesAtScreenshot")]
        public bool IsAutoSaveStrokesAtScreenshot { get; set; } = false;

        [JsonProperty("isAutoSaveStrokesAtClear")]
        public bool IsAutoSaveStrokesAtClear { get; set; } = false;

        [JsonProperty("isAutoClearWhenExitingWritingMode")]
        public bool IsAutoClearWhenExitingWritingMode { get; set; } = false;

        [JsonProperty("minimumAutomationStrokeNumber")]
        public int MinimumAutomationStrokeNumber { get; set; } = 0;

    }

    public class Advanced
    {
        [JsonProperty("isSpecialScreen")]
        public bool IsSpecialScreen { get; set; } = false;
        [JsonProperty("isQuadIR")]
        public bool IsQuadIR { get; set; } = false;
        [JsonProperty("touchMultiplier")]
        public double TouchMultiplier { get; set; } = 0.25;
        [JsonProperty("eraserBindTouchMultiplier")]
        public bool EraserBindTouchMultiplier { get; set; } = false;
        [JsonProperty("isLogEnabled")]
        public bool IsLogEnabled { get; set; } = true;
    }

    public class InkToShape
    {
        [JsonProperty("isInkToShapeEnabled")]
        public bool IsInkToShapeEnabled { get; set; } = true;
    }
}
