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
        public bool IsShowCursor { get; set; } = false;
        [JsonProperty("inkStyle")]
        public int InkStyle { get; set; } = 0;
        [JsonProperty("eraserSize")]
        public int EraserSize { get; set; } = 2;
        [JsonProperty("eraserType")]
        public int EraserType { get; set; } = 0; // 0 - 图标切换模式      1 - 面积擦     2 - 线条擦
        [JsonProperty("hideStrokeWhenSelecting")]
        public bool HideStrokeWhenSelecting { get; set; } = true;

        [JsonProperty("usingWhiteboard")]
        public bool UsingWhiteboard { get; set; }

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
