using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace Ink_Canvas
{
    /// <summary>
    /// MainWindow 分部类：动态快捷键系统（窗口焦点时生效，WPF KeyBinding）。
    /// 从 Settings.Shortcuts 读取按键串动态注册，支持自定义；改动后即时生效无需重启。
    /// 不走全局热键（RegisterHotKey），避免与系统/其他软件冲突。
    /// </summary>
    public partial class MainWindow
    {
        #region Dynamic Shortcuts

        /// <summary>
        /// 可自定义的动作目录：ID → (显示名, 默认按键, 执行动作)
        /// 全部复用现有 handler/命令，不重复造轮子。
        /// </summary>
        private List<ShortcutAction> _shortcutActions;

        private class ShortcutAction
        {
            public string Id;          // 配置键，如 "Pen"
            public string Name;        // 界面显示名
            public string DefaultGesture; // 默认按键串，如 "F2"
            public Action Execute;     // 触发时执行（委托给现有 handler）
        }

        /// <summary>初始化动态快捷键（窗口 Loaded 后调用，此时句柄已就绪）</summary>
        private void InitDynamicShortcuts()
        {
            if (_shortcutActions != null) return; // 防重复初始化

            _shortcutActions = new List<ShortcutAction>
            {
                new ShortcutAction { Id = "Pen",       Name = "画笔",       DefaultGesture = "F2",
                    Execute = () => BtnPen_Click(BtnPen, null) },
                new ShortcutAction { Id = "Eraser",    Name = "橡皮",       DefaultGesture = "F3",
                    Execute = () => KeyChangeToEraser(null, null) },
                new ShortcutAction { Id = "Select",    Name = "选择",       DefaultGesture = "F4",
                    Execute = () => KeyChangeToSelect(null, null) },
                new ShortcutAction { Id = "Undo",      Name = "撤销",       DefaultGesture = "Ctrl+Z",
                    Execute = () => back_HotKey(null, null) },
                new ShortcutAction { Id = "Redo",      Name = "重做",       DefaultGesture = "Ctrl+Y",
                    Execute = () => BtnRedo_Click(BtnRedo, null) },
                new ShortcutAction { Id = "Clear",     Name = "清空墨迹",   DefaultGesture = "Ctrl+Del",
                    Execute = () => BtnClear_Click(BtnClear, null) },
                new ShortcutAction { Id = "DrawShape", Name = "图形面板",   DefaultGesture = "F6",
                    Execute = () => ImageDrawShape_MouseUp(null, null) },
                new ShortcutAction { Id = "ToggleBar", Name = "展开/收起悬浮条", DefaultGesture = "F8",
                    Execute = () => SetBorderFloatingBarMainControlsVisibility(!borderFloatingBarMainControlsVisibility) },
                new ShortcutAction { Id = "Settings",  Name = "设置",       DefaultGesture = "F10",
                    Execute = () => SymbolIconSettings_Click(null, null) },
                new ShortcutAction { Id = "Restart",   Name = "重启画板",   DefaultGesture = "Ctrl+Shift+R",
                    Execute = () =>
                    {
                        // 有墨迹时先确认，防误触丢板书（复用现有确认逻辑风格）
                        if (inkCanvas.Strokes.Count > 0)
                        {
                            var r = MessageBox.Show($"画布上有 {inkCanvas.Strokes.Count} 笔墨迹，重启将丢失。\n确定重启画板？",
                                "重启画板", MessageBoxButton.YesNo, MessageBoxImage.Warning);
                            if (r != MessageBoxResult.Yes) return;
                        }
                        BtnRestart_Click(null, null);
                    } },
                new ShortcutAction { Id = "Exit",      Name = "结束放映/退出", DefaultGesture = "Shift+Esc",
                    Execute = () => KeyExit(null, null) },
                new ShortcutAction { Id = "ScaleUp",   Name = "放大批注",   DefaultGesture = "Ctrl+=",
                    Execute = () => ScaleAllOrSelection(1.1) },
                new ShortcutAction { Id = "ScaleDown", Name = "缩小批注",   DefaultGesture = "Ctrl+-",
                    Execute = () => ScaleAllOrSelection(0.9) },
            };

            RebindAllShortcuts();
        }

        /// <summary>清除并重新注册所有动态快捷键（配置改动后调用，即时生效）</summary>
        private void RebindAllShortcuts()
        {
            if (_shortcutActions == null) return;

            // 移除所有动态注册的 KeyBinding（静态 XAML 里的保留）
            var toRemove = InputBindings.OfType<KeyBinding>()
                .Where(kb => kb.Command is RoutedUICommand cmd && cmd.Name.StartsWith("dyn:"))
                .ToList();
            foreach (var kb in toRemove) InputBindings.Remove(kb);

            // 逐个注册：解析按键串 → 构造 KeyBinding → 绑到动作
            foreach (var action in _shortcutActions)
            {
                string gesture = Settings.Shortcuts.Get(action.Id, action.DefaultGesture);
                if (!TryParseGesture(gesture, out ModifierKeys mods, out Key key)) continue;

                // 冲突检测：同键不同动作，后注册的覆盖（以最后保存的为准）
                var existing = InputBindings.OfType<KeyBinding>()
                    .FirstOrDefault(kb => kb.Command is RoutedUICommand cmd2 && cmd2.Name.StartsWith("dyn:")
                        && kb.Modifiers == mods && kb.Key == key);
                if (existing != null) InputBindings.Remove(existing);

                var command = new RoutedUICommand(action.Name, "dyn:" + action.Id, typeof(MainWindow));
                var binding = new KeyBinding(command, key, mods);
                InputBindings.Add(binding);

                //数字小键盘 +/- 也映射到放大/缩小批注（主键盘 OemPlus/OemMinus 之外的第二位置）
                if (action.Id == "ScaleUp" || action.Id == "ScaleDown")
                {
                    InputBindings.Add(new KeyBinding(command,
                        action.Id == "ScaleUp" ? Key.Add : Key.Subtract, mods));
                }

                // 命令 → 动作：用 CommandBinding 挂到窗口，Executed 时调用对应 handler
                CommandBindings.Add(new CommandBinding(command,
                    (s, e) => action.Execute(),
                    (s, e) => e.CanExecute = true));
            }
        }

        /// <summary>
        /// 解析按键串（如 "F2"、"Ctrl+Shift+S"、"Alt+D1"）为 ModifierKeys + Key。
        /// 支持 Ctrl/Alt/Shift/Win 任意组合 + 主键（字母/数字/功能键/方向键/Delete等）。
        /// 解析失败返回 false（不注册，不崩）。
        /// </summary>
        private bool TryParseGesture(string gesture, out ModifierKeys modifiers, out Key key)
        {
            modifiers = ModifierKeys.None;
            key = Key.None;
            if (string.IsNullOrWhiteSpace(gesture)) return false;

            var parts = gesture.Split('+').Select(p => p.Trim()).ToArray();
            if (parts.Length == 0) return false;

            // 最后一段是主键，前面是修饰键
            string mainKeyStr = parts[parts.Length - 1];
            for (int i = 0; i < parts.Length - 1; i++)
            {
                switch (parts[i].ToLowerInvariant())
                {
                    case "ctrl": case "control": modifiers |= ModifierKeys.Control; break;
                    case "alt": modifiers |= ModifierKeys.Alt; break;
                    case "shift": modifiers |= ModifierKeys.Shift; break;
                    case "win": case "windows": modifiers |= ModifierKeys.Windows; break;
                    default: return false;
                }
            }

            // 主键解析：支持单字符（A-Z、0-9）、F1-F24、方向键、Delete、Space、Tab 等
            if (!TryParseKey(mainKeyStr, out key)) return false;
            return true;
        }

        private bool TryParseKey(string s, out Key key)
        {
            key = Key.None;
            if (string.IsNullOrEmpty(s)) return false;

            // 单字符：字母/数字（D0-D9 / A-Z）、加/减号（OemPlus/OemMinus）
            if (s.Length == 1)
            {
                char c = s[0];
                if (char.IsLetter(c)) { key = (Key)Enum.Parse(typeof(Key), c.ToString().ToUpperInvariant()); return true; }
                if (char.IsDigit(c)) { key = (Key)Enum.Parse(typeof(Key), "D" + c); return true; }
                if (c == '=') { key = Key.OemPlus; return true; }
                if (c == '-') { key = Key.OemMinus; return true; }
                return false;
            }

            // 功能键 F1-F24
            if (s.StartsWith("F", StringComparison.OrdinalIgnoreCase) && s.Length <= 3)
            {
                if (int.TryParse(s.Substring(1), out int n) && n >= 1 && n <= 24)
                {
                    key = (Key)Enum.Parse(typeof(Key), "F" + n);
                    return true;
                }
                return false;
            }

            // 命名键
            switch (s.ToLowerInvariant())
            {
                case "space": key = Key.Space; return true;
                case "tab": key = Key.Tab; return true;
                case "enter": case "return": key = Key.Enter; return true;
                case "esc": case "escape": key = Key.Escape; return true;
                case "del": case "delete": key = Key.Delete; return true;
                case "back": case "backspace": key = Key.Back; return true;
                case "up": key = Key.Up; return true;
                case "down": key = Key.Down; return true;
                case "left": key = Key.Left; return true;
                case "right": key = Key.Right; return true;
                case "home": key = Key.Home; return true;
                case "end": key = Key.End; return true;
                case "pageup": key = Key.PageUp; return true;
                case "pagedown": key = Key.PageDown; return true;
                case "insert": key = Key.Insert; return true;
                default: return false;
            }
        }

        /// <summary>获取某动作当前生效的按键串（用于界面显示）</summary>
        public string GetShortcutDisplay(string actionId)
        {
            var action = _shortcutActions?.FirstOrDefault(a => a.Id == actionId);
            if (action == null) return "";
            return Settings.Shortcuts.Get(action.Id, action.DefaultGesture);
        }

        /// <summary>修改某动作的快捷键并立即生效（供设置界面调用）</summary>
        public void SetShortcut(string actionId, string gesture)
        {
            Settings.Shortcuts.Set(actionId, gesture);
            SaveSettingsToFile();
            RebindAllShortcuts();
        }

        /// <summary>恢复某动作为默认快捷键</summary>
        public void ResetShortcut(string actionId)
        {
            Settings.Shortcuts.Bindings?.Remove(actionId);
            SaveSettingsToFile();
            RebindAllShortcuts();
        }

        /// <summary>获取全部可配置动作（供设置界面/右键菜单遍历）</summary>
        public IReadOnlyList<(string id, string name, string gesture)> GetAllShortcuts()
        {
            if (_shortcutActions == null) InitDynamicShortcuts();
            return _shortcutActions.Select(a => (a.Id, a.Name,
                Settings.Shortcuts.Get(a.Id, a.DefaultGesture))).ToList();
        }

        /// <summary>右键菜单：查看当前全部快捷键（只读弹窗）</summary>
        private void MenuItemShowShortcuts_Click(object sender, RoutedEventArgs e)
        {
            var all = GetAllShortcuts();
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("当前快捷键（窗口有焦点时生效）：");
            sb.AppendLine();
            foreach (var (id, name, gesture) in all)
                sb.AppendLine($"  {name,-8} {gesture}");
            sb.AppendLine();
            sb.AppendLine("修改方式：设置 → 快捷键（后续开放自定义界面）");
            sb.AppendLine("           或直接编辑 Settings.json 中 shortcuts.bindings");
            MessageBox.Show(sb.ToString(), "快捷键一览", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        /// <summary>右键菜单：恢复全部快捷键为默认</summary>
        private void MenuItemResetShortcuts_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show("确定恢复所有快捷键为默认值？\n当前自定义配置将被清除。",
                "恢复默认快捷键", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (result != MessageBoxResult.Yes) return;

            Settings.Shortcuts.Bindings?.Clear();
            SaveSettingsToFile();
            RebindAllShortcuts();
            ShowToastNotification("已恢复默认快捷键");
        }

        #endregion
    }
}
