using System;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Interop;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：全局快捷键（统一快捷键设置）。
    /// 基于系统级 RegisterHotKey（Helpers/Hotkey.cs），任何应用下按都触发，不依赖画布焦点
    /// ——这正是透明覆盖窗口（焦点不可靠）下做"重启"等操作所需要的。
    /// 与原版窗口级 KeyBinding（Alt+D/E/C 等，仅画布有焦点时生效）互不冲突，并行存在。
    /// 新增快捷键三步：Settings.Shortcuts 加属性 → 本文件 RegisterAll 里注册 → 设置面板 GroupBox 加一行录制控件。
    /// </summary>
    public partial class MainWindow
    {
        #region 全局快捷键

        //持久的回调引用：Hotkey.UnRegist 按 delegate 匹配注销，必须与注册时同一实例
        private readonly Hotkey.HotKeyCallBackHanlder hotkeyCallbackRestart = HotkeyRestartAction;

        /// <summary>启动时加载（MW_Init 调用）：按设置注册全部全局热键</summary>
        private void InitShortcuts()
        {
            if (!Settings.Shortcuts.IsGlobalShortcutsEnabled) return;
            RegisterAllGlobalShortcuts();
        }

        /// <summary>注册全部全局热键。重复调用安全（先全部注销）。</summary>
        private void RegisterAllGlobalShortcuts()
        {
            UnregisterAllGlobalShortcuts();

            var hwnd = new WindowInteropHelper(this).Handle;
            bool ok = Hotkey.Regist(this, ParseGesture(Settings.Shortcuts.Restart).modifiers,
                ParseGesture(Settings.Shortcuts.Restart).key, hotkeyCallbackRestart);
            if (!ok)
            {
                //注册失败=组合键已被其他程序占用（如录屏/截图软件），提示但不禁用开关，用户可改键后重试
                MessageBox.Show($"全局快捷键「{Settings.Shortcuts.Restart}」（重启画板）注册失败：该组合键可能已被其他程序占用。\n" +
                    "可在 设置 → 快捷键 中更换组合键。", "快捷键冲突", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        /// <summary>注销全部全局热键（设置变更/关闭开关/退出前调用）</summary>
        private void UnregisterAllGlobalShortcuts()
        {
            try
            {
                var hwnd = new WindowInteropHelper(this).Handle;
                Hotkey.UnRegist(hwnd, hotkeyCallbackRestart);
            }
            catch { }
        }

        /// <summary>
        /// 快捷键动作：重启画板（复用 BtnRestart_Click 的重启逻辑：拉新实例绕过单实例锁再关自己）。
        /// 与设置面板按钮的差异：热键可能误触，有未保存墨迹时先确认，防一键丢板书。
        /// </summary>
        private static void HotkeyRestartAction()
        {
            Application.Current.Dispatcher.Invoke(() =>
            {
                var mw = Application.Current.MainWindow as MainWindow;
                if (mw == null) return;
                if (mw.inkCanvas.Strokes.Count > 0)
                {
                    var result = MessageBox.Show(
                        $"快捷键重启画板将清除当前 {mw.inkCanvas.Strokes.Count} 笔墨迹（如开启自动保存则会先存档）。\n确定重启吗？",
                        "重启画板", MessageBoxButton.OKCancel, MessageBoxImage.Warning);
                    if (result != MessageBoxResult.OK) return;
                }
                mw.BtnRestart_Click(null, null);
            });
        }

        #region 手势字符串解析与生成

        /// <summary>解析 "Ctrl+Alt+R" → (修饰键, 主键)。解析不了时回退 Ctrl+Alt+R，保证热键永不因脏配置失效。</summary>
        private (HotkeyModifiers modifiers, Key key) ParseGesture(string gesture)
        {
            var fallback = (HotkeyModifiers.MOD_CONTROL | HotkeyModifiers.MOD_ALT, Key.R);
            if (string.IsNullOrWhiteSpace(gesture)) return fallback;

            HotkeyModifiers mods = 0;
            Key mainKey = Key.None;
            foreach (var part in gesture.Split('+'))
            {
                var p = part.Trim();
                if (p.Equals("Ctrl", StringComparison.OrdinalIgnoreCase) || p.Equals("Control", StringComparison.OrdinalIgnoreCase)) mods |= HotkeyModifiers.MOD_CONTROL;
                else if (p.Equals("Alt", StringComparison.OrdinalIgnoreCase)) mods |= HotkeyModifiers.MOD_ALT;
                else if (p.Equals("Shift", StringComparison.OrdinalIgnoreCase)) mods |= HotkeyModifiers.MOD_SHIFT;
                else if (p.Equals("Win", StringComparison.OrdinalIgnoreCase)) mods |= HotkeyModifiers.MOD_WIN;
                else if (Enum.TryParse(p, true, out Key k)) mainKey = k;
            }

            if (mainKey == Key.None) return fallback;
            //防裸键/纯Shift冲突（如单按 R 会吞掉正常打字）：必须含 Ctrl/Alt/Win 之一
            if ((mods & (HotkeyModifiers.MOD_CONTROL | HotkeyModifiers.MOD_ALT | HotkeyModifiers.MOD_WIN)) == 0)
                return fallback;
            return (mods, mainKey);
        }

        /// <summary>录制到的按键组合 → 手势字符串（与解析互逆，用于保存/回显）</summary>
        private static string BuildGesture(ModifierKeys mods, Key key)
        {
            var sb = new StringBuilder();
            if ((mods & ModifierKeys.Control) != 0) sb.Append("Ctrl+");
            if ((mods & ModifierKeys.Alt) != 0) sb.Append("Alt+");
            if ((mods & ModifierKeys.Shift) != 0) sb.Append("Shift+");
            if ((mods & ModifierKeys.Windows) != 0) sb.Append("Win+");
            sb.Append(key.ToString());
            return sb.ToString();
        }

        /// <summary>手势是否有效：主键非空且含 Ctrl/Alt/Win 修饰键（录制时的校验口径，与 ParseGesture 回退口径一致）</summary>
        private static bool IsValidGesture(ModifierKeys mods, Key key)
        {
            if (key == Key.None || key == Key.System || key == Key.LeftCtrl || key == Key.RightCtrl ||
                key == Key.LeftAlt || key == Key.RightAlt || key == Key.LeftShift || key == Key.RightShift ||
                key == Key.LWin || key == Key.RWin) return false;
            var needed = ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Windows;
            return (mods & needed) != 0;
        }

        #endregion

        #region 设置面板交互

        /// <summary>总开关：开=注册全部热键；关=全部注销</summary>
        private void ToggleSwitchGlobalShortcuts_Toggled(object sender, RoutedEventArgs e)
        {
            if (!isLoaded) return;
            Settings.Shortcuts.IsGlobalShortcutsEnabled = ToggleSwitchGlobalShortcuts.IsOn;
            SaveSettingsToFile();

            if (Settings.Shortcuts.IsGlobalShortcutsEnabled) RegisterAllGlobalShortcuts();
            else UnregisterAllGlobalShortcuts();
        }

        private bool isShortcutRecording = false; //录制态标记：窗口级 KeyBinding（Ctrl+Z 等）在此期间不可抢先

        /// <summary>
        /// 录制框获得焦点即进入录制态：清空显示"请按下组合键…"。
        /// TextBox 天然聚焦，PreviewKeyDown 先于窗口 InputBinding，标记 Handled 拦截原窗口级快捷键。
        /// </summary>
        private void TextBoxShortcutRecord_GotFocus(object sender, RoutedEventArgs e)
        {
            isShortcutRecording = true;
            ((TextBox)sender).Text = "请按下组合键（Esc 取消）";
        }

        /// <summary>录制框按键：捕获修饰键+主键组合（Esc 取消并退出录制态）</summary>
        private void TextBoxShortcutRecord_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (!isShortcutRecording) return;
            var tb = (TextBox)sender;
            e.Handled = true; //吞掉一切按键：不进 TextBox、不触发窗口级 KeyBinding

            if (e.Key == Key.Escape)
            {
                RestoreShortcutTextFromSettings(tb);
                return;
            }

            //修饰键单独按下时不响应，等主键
            var key = e.Key == Key.System ? e.SystemKey : e.Key; //Alt 组合会以 System 键上报
            var mods = Keyboard.Modifiers;
            if (!IsValidGesture(mods, key)) return;

            var gesture = BuildGesture(mods, key);
            ApplyShortcutGesture(tb, gesture);
        }

        /// <summary>失去焦点时若仍在录制态（未完成录制），恢复显示当前生效值</summary>
        private void TextBoxShortcutRecord_LostFocus(object sender, RoutedEventArgs e)
        {
            isShortcutRecording = false;
            var tb = (TextBox)sender;
            RestoreShortcutTextFromSettings(tb);
        }

        /// <summary>应用新手势：保存→立即重注册（生效即刻可验证）。冲突时 RegisterAllGlobalShortcuts 已弹框警告。</summary>
        private void ApplyShortcutGesture(TextBox tb, string gesture)
        {
            isShortcutRecording = false;
            Settings.Shortcuts.Restart = gesture;
            SaveSettingsToFile();
            RegisterAllGlobalShortcuts();
            tb.Text = gesture; //MessageBox 若抢走焦点触发 LostFocus，恢复显示的同样是新值（单一来源 Settings）
        }

        /// <summary>从设置恢复录制框显示（非录制态的显示值来源，单一事实源）</summary>
        private void RestoreShortcutTextFromSettings(TextBox tb)
        {
            tb.Text = Settings.Shortcuts.Restart;
        }

        #endregion

        #endregion 全局快捷键
    }
}
