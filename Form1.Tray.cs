using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
namespace ClipInjector
{
    public partial class Form1 : Form
    {
        private const int HOTKEY_ID = 1;
        private const int WM_HOTKEY = 0x0312;
        private const uint MOD_ALT = 0x0001;
        private const uint MOD_CONTROL = 0x0002;
        private const uint MOD_SHIFT = 0x0004;
        private const uint MOD_WIN = 0x0008;

        // SendInput flags
        private const uint KEYEVENTF_KEYUP = 0x0002;
        private const uint KEYEVENTF_UNICODE = 0x0004;
        private const uint INPUT_KEYBOARD = 1;

        // VKs for releasing modifiers
        private const ushort VK_CONTROL = 0x11;
        private const ushort VK_MENU = 0x12;  // Alt
        private const ushort VK_SHIFT = 0x10;
        private const ushort VK_LWIN = 0x5B;

        private NotifyIcon? _trayIcon; // only tray icon
        private string _activeHotkeyDescription = "";

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Opacity = 0;
            InitializeTrayIcon();
            TryRegisterHotkeys();
        }

        private void TryRegisterHotkeys()
        {
            // Preferred order: Win+Shift+V, then Ctrl+Alt+Shift+V, Ctrl+Shift+V, Ctrl+Alt+V
            var attempts = new (uint mods, string desc)[]
            {
                (MOD_WIN | MOD_SHIFT, "Win+Shift+V"),
                (MOD_CONTROL | MOD_ALT | MOD_SHIFT, "Ctrl+Alt+Shift+V"),
                (MOD_CONTROL | MOD_SHIFT, "Ctrl+Shift+V"),
                (MOD_CONTROL | MOD_ALT, "Ctrl+Alt+V")
            };
            foreach (var a in attempts)
            {
                if (RegisterHotKey(Handle, HOTKEY_ID, a.mods, (uint)Keys.V))
                {
                    _activeHotkeyDescription = a.desc;
                    ShowBalloon("ClipInjector", $"Ready ({_activeHotkeyDescription})");
                    return;
                }
            }
            _activeHotkeyDescription = "<none>";
            ShowBalloon("ClipInjector", "Failed to register any hotkey");
        }

        protected override void OnHandleDestroyed(EventArgs e)
        {
            try { UnregisterHotKey(Handle, HOTKEY_ID); } catch { }
            base.OnHandleDestroyed(e);
        }

        private void InitializeTrayIcon()
        {
            _trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Text = "ClipInjector",
                Visible = true
            };
            _trayIcon.DoubleClick += (_, __) => InjectClipboard();
        }

        private void ShowBalloon(string title, string text)
        {
            if (_trayIcon == null) return;
            try
            {
                _trayIcon.BalloonTipTitle = title;
                _trayIcon.BalloonTipText = text;
                _trayIcon.ShowBalloonTip(3000);
            }
            catch { }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
            {
                InjectClipboard();
            }
            base.WndProc(ref m);
        }

        private void InjectClipboard()
        {
            try
            {
                if (!Clipboard.ContainsText()) return;
                var raw = Clipboard.GetText();
                if (string.IsNullOrEmpty(raw)) return;

                var normalized = raw.Replace("\r\n", "\n").Replace("\r", "\n");

                // 1) Make sure hotkey modifiers don't interfere
                ReleaseModifiers();
                Thread.Sleep(25);

                // 2) Try Unicode injection first; if any part fails, fall back
                if (!SendUnicodeString(normalized))
                {
                    SendAsciiFallback(normalized);
                }
            }
            catch { }
        }

        // --- NEW: release modifiers so Ctrl/Alt/Shift/Win don't corrupt typing
        private void ReleaseModifiers()
        {
            var ups = new List<INPUT>
            {
                MakeVk(VK_CONTROL, true),
                MakeVk(VK_MENU,    true),
                MakeVk(VK_SHIFT,   true),
                MakeVk(VK_LWIN,    true),
            };
            if (ups.Count > 0)
                SendInput((uint)ups.Count, ups.ToArray(), Marshal.SizeOf<INPUT>());
        }

        // --- CHANGED: check SendInput return and report; return true when fully sent
        private bool SendUnicodeString(string text)
        {
            var inputs = new List<INPUT>(text.Length * 2);
            foreach (char c in text)
            {
                inputs.Add(MakeInput(c, false));
                inputs.Add(MakeInput(c, true));
            }
            if (inputs.Count == 0) return true;

            uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
            if (sent != inputs.Count)
            {
                int err = Marshal.GetLastWin32Error();
                ShowBalloon("ClipInjector", $"Unicode send partial ({sent}/{inputs.Count}), err {err}.");
                string errorMessage = new Win32Exception(err).Message;
                Debug.WriteLine(errorMessage);
                return false;
            }
            return true;
        }

        // --- NEW: fallback path for apps that ignore Unicode VK_PACKET
        private bool SendAsciiFallback(string text)
        {
            var hkl = GetKeyboardLayout(0);
            var inputs = new List<INPUT>();

            foreach (char ch in text)
            {
                // handle newline explicitly
                if (ch == '\n')
                {
                    inputs.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)Keys.Return, dwFlags = 0 } } });
                    inputs.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = (ushort)Keys.Return, dwFlags = KEYEVENTF_KEYUP } } });
                    continue;
                }

                short vkPacked = VkKeyScanEx(ch, hkl);
                if (vkPacked == -1)
                {
                    // can't map -> fall back to unicode for this char only
                    inputs.Add(MakeInput(ch, false));
                    inputs.Add(MakeInput(ch, true));
                    continue;
                }

                byte vk = (byte)(vkPacked & 0xFF);
                byte shiftState = (byte)((vkPacked >> 8) & 0xFF);

                // press required modifiers
                if ((shiftState & 1) != 0) inputs.Add(MakeVk(VK_SHIFT, false));
                if ((shiftState & 2) != 0) inputs.Add(MakeVk(VK_CONTROL, false));
                if ((shiftState & 4) != 0) inputs.Add(MakeVk(VK_MENU, false));

                // key down/up
                inputs.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = 0 } } });
                inputs.Add(new INPUT { type = INPUT_KEYBOARD, U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = KEYEVENTF_KEYUP } } });

                // release modifiers
                if ((shiftState & 4) != 0) inputs.Add(MakeVk(VK_MENU, true));
                if ((shiftState & 2) != 0) inputs.Add(MakeVk(VK_CONTROL, true));
                if ((shiftState & 1) != 0) inputs.Add(MakeVk(VK_SHIFT, true));
            }

            if (inputs.Count == 0) return true;
            uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
            if (sent != inputs.Count)
            {
                int err = Marshal.GetLastWin32Error();
                ShowBalloon("ClipInjector", $"ASCII fallback partial ({sent}/{inputs.Count}), err {err}.");
                string errorMessage = new Win32Exception(err).Message;
                Debug.WriteLine(errorMessage);
                return false;
            }
            return true;
        }

        private INPUT MakeInput(char c, bool keyUp)
        {
            return new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = 0,
                        wScan = c,
                        dwFlags = KEYEVENTF_UNICODE | (keyUp ? KEYEVENTF_KEYUP : 0),
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
        }

        // helper for modifier VKs
        private INPUT MakeVk(ushort vk, bool keyUp)
        {
            return new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion
                {
                    ki = new KEYBDINPUT
                    {
                        wVk = vk,
                        wScan = 0,
                        dwFlags = keyUp ? KEYEVENTF_KEYUP : 0,
                        time = 0,
                        dwExtraInfo = IntPtr.Zero
                    }
                }
            };
        }

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        // Optional (if you use ASCII fallback elsewhere)
        [DllImport("user32.dll")]
        private static extern IntPtr GetKeyboardLayout(uint idThread);
        [DllImport("user32.dll")]
        private static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

        [StructLayout(LayoutKind.Sequential)]
        private struct INPUT
        {
            public uint type;       // 1 = keyboard
            public InputUnion U;    // largest of mouse/keyboard/hardware
        }

        [StructLayout(LayoutKind.Explicit)]
        private struct InputUnion
        {
            [FieldOffset(0)] public MOUSEINPUT mi;   // ensure union size matches native
            [FieldOffset(0)] public KEYBDINPUT ki;
            [FieldOffset(0)] public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct KEYBDINPUT
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct HARDWAREINPUT
        {
            public uint uMsg;
            public ushort wParamL;
            public ushort wParamH;
        }
    }
}
