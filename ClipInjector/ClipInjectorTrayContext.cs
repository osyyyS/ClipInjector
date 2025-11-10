using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ClipInjector;

internal sealed class ClipInjectorTrayContext : ApplicationContext
{
    private const int HOTKEY_ID = 1;
    private const int WM_HOTKEY = 0x0312;
    private const uint MOD_ALT = 0x0001;
    private const uint MOD_CONTROL = 0x0002;

    private const uint INPUT_KEYBOARD = 1;
    private const uint KEYEVENTF_KEYUP = 0x0002;
    private const uint KEYEVENTF_UNICODE = 0x0004;

    private const ushort VK_CONTROL = 0x11;
    private const ushort VK_MENU = 0x12; // Alt
    private const ushort VK_SHIFT = 0x10;
    private const ushort VK_LWIN = 0x5B;

    private readonly TrayWindow _window;
    private NotifyIcon? _trayIcon;
    private ContextMenuStrip? _trayMenu;
    private string _activeHotkeyDescription = string.Empty;

    public ClipInjectorTrayContext()
    {
        _window = new TrayWindow(this);

        InitializeTrayIcon();
        RegisterHotkey();
    }

    protected override void ExitThreadCore()
    {
        try
        {
            UnregisterHotKey(_window.Handle, HOTKEY_ID);
        }
        catch
        {
            // ignored
        }

        if (_trayIcon != null)
        {
            _trayIcon.DoubleClick -= TrayIconOnDoubleClick;
            _trayIcon.Visible = false;
            _trayIcon.Dispose();
            _trayIcon = null;
        }

        if (_trayMenu != null)
        {
            foreach (ToolStripItem item in _trayMenu.Items)
            {
                if (item is ToolStripMenuItem menuItem)
                {
                    menuItem.Click -= HandleTrayMenuClick;
                }
            }

            _trayMenu.Dispose();
            _trayMenu = null;
        }

        _window.DestroyHandle();

        base.ExitThreadCore();
    }

    private void RegisterHotkey()
    {
        const uint modifiers = MOD_CONTROL | MOD_ALT;
        if (RegisterHotKey(_window.Handle, HOTKEY_ID, modifiers, (uint)Keys.V))
        {
            _activeHotkeyDescription = "Ctrl+Alt+V";
            ShowBalloon("ClipInjector", $"Ready ({_activeHotkeyDescription})");
            UpdateTrayText();
        }
        else
        {
            _activeHotkeyDescription = "<none>";
            ShowBalloon("ClipInjector", "Failed to register Ctrl+Alt+V hotkey");
            UpdateTrayText();
        }
    }

    private void InitializeTrayIcon()
    {
        _trayMenu = new ContextMenuStrip();

        var exitItem = new ToolStripMenuItem("Exit")
        {
            Tag = "exit"
        };
        exitItem.Click += HandleTrayMenuClick;

        _trayMenu.Items.Add(new ToolStripSeparator());
        _trayMenu.Items.Add(exitItem);

        _trayIcon = new NotifyIcon
        {
            Icon = SystemIcons.Application,
            Visible = true,
            ContextMenuStrip = _trayMenu
        };

        var assembly = Assembly.GetExecutingAssembly();
        using var stream = assembly.GetManifestResourceStream("ClipInjector.AppIcon.ico");
        var icon = new Icon(stream);
        _trayIcon.Icon = icon;

        _trayIcon.DoubleClick += TrayIconOnDoubleClick;
        UpdateTrayText();
    }

    private void TrayIconOnDoubleClick(object? sender, EventArgs e) => InjectClipboard();

    private void HandleTrayMenuClick(object? sender, EventArgs e)
    {
        if (sender is not ToolStripItem item)
        {
            return;
        }

        switch (item.Tag as string)
        {
            case "inject":
                InjectClipboard();
                break;
            case "exit":
                ExitThread();
                break;
        }
    }

    private void UpdateTrayText()
    {
        if (_trayIcon == null)
        {
            return;
        }

        string suffix = string.IsNullOrWhiteSpace(_activeHotkeyDescription)
            ? "<no hotkey>"
            : _activeHotkeyDescription;

        _trayIcon.Text = $"ClipInjector ({suffix})";
    }

    private void ShowBalloon(string title, string text)
    {
        if (_trayIcon == null)
        {
            return;
        }

        try
        {
            _trayIcon.BalloonTipTitle = title;
            _trayIcon.BalloonTipText = text;
            _trayIcon.ShowBalloonTip(3000);
        }
        catch
        {
            // ignored
        }
    }

    private void HandleHotkey(int hotkeyId)
    {
        if (hotkeyId == HOTKEY_ID)
        {
            InjectClipboard();
        }
    }

    private void InjectClipboard()
    {
        try
        {
            if (!Clipboard.ContainsText())
            {
                return;
            }

            string raw = Clipboard.GetText();
            if (string.IsNullOrEmpty(raw))
            {
                return;
            }

            string normalized = raw.Replace("\r\n", "\n").Replace("\r", "\n");

            ReleaseModifiers();
            Thread.Sleep(25);

            if (!SendUnicodeString(normalized))
            {
                SendAsciiFallback(normalized);
            }
        }
        catch
        {
            // ignored
        }
    }

    private void ReleaseModifiers()
    {
        var ups = new List<INPUT>
        {
            MakeVk(VK_CONTROL, true),
            MakeVk(VK_MENU, true),
            MakeVk(VK_SHIFT, true),
            MakeVk(VK_LWIN, true)
        };

        if (ups.Count > 0)
        {
            SendInput((uint)ups.Count, ups.ToArray(), Marshal.SizeOf<INPUT>());
        }
    }

    private bool SendUnicodeString(string text)
    {
        var inputs = new List<INPUT>(text.Length * 2);
        foreach (char c in text)
        {
            inputs.Add(MakeInput(c, false));
            inputs.Add(MakeInput(c, true));
        }

        if (inputs.Count == 0)
        {
            return true;
        }

        uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
        if (sent != inputs.Count)
        {
            int err = Marshal.GetLastWin32Error();
            string errorMessage = new Win32Exception(err).Message;
            ShowBalloon("ClipInjector", $"Unicode send partial ({sent}/{inputs.Count}): {errorMessage} (0x{err:X}).");
            return false;
        }

        return true;
    }

    private bool SendAsciiFallback(string text)
    {
        IntPtr layout = GetKeyboardLayout(0);
        var inputs = new List<INPUT>();

        foreach (char ch in text)
        {
            if (ch == '\n')
            {
                inputs.Add(new INPUT
                {
                    type = INPUT_KEYBOARD,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT { wVk = (ushort)Keys.Return, dwFlags = 0 }
                    }
                });
                inputs.Add(new INPUT
                {
                    type = INPUT_KEYBOARD,
                    U = new InputUnion
                    {
                        ki = new KEYBDINPUT { wVk = (ushort)Keys.Return, dwFlags = KEYEVENTF_KEYUP }
                    }
                });
                continue;
            }

            short vkPacked = VkKeyScanEx(ch, layout);
            if (vkPacked == -1)
            {
                inputs.Add(MakeInput(ch, false));
                inputs.Add(MakeInput(ch, true));
                continue;
            }

            byte vk = (byte)(vkPacked & 0xFF);
            byte shiftState = (byte)((vkPacked >> 8) & 0xFF);

            if ((shiftState & 1) != 0)
            {
                inputs.Add(MakeVk(VK_SHIFT, false));
            }
            if ((shiftState & 2) != 0)
            {
                inputs.Add(MakeVk(VK_CONTROL, false));
            }
            if ((shiftState & 4) != 0)
            {
                inputs.Add(MakeVk(VK_MENU, false));
            }

            inputs.Add(new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = 0 } }
            });
            inputs.Add(new INPUT
            {
                type = INPUT_KEYBOARD,
                U = new InputUnion { ki = new KEYBDINPUT { wVk = vk, dwFlags = KEYEVENTF_KEYUP } }
            });

            if ((shiftState & 4) != 0)
            {
                inputs.Add(MakeVk(VK_MENU, true));
            }
            if ((shiftState & 2) != 0)
            {
                inputs.Add(MakeVk(VK_CONTROL, true));
            }
            if ((shiftState & 1) != 0)
            {
                inputs.Add(MakeVk(VK_SHIFT, true));
            }
        }

        if (inputs.Count == 0)
        {
            return true;
        }

        uint sent = SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf<INPUT>());
        if (sent != inputs.Count)
        {
            int err = Marshal.GetLastWin32Error();
            string errorMessage = new Win32Exception(err).Message;
            ShowBalloon("ClipInjector", $"ASCII fallback partial ({sent}/{inputs.Count}): {errorMessage} (0x{err:X}).");
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

    private sealed class TrayWindow : NativeWindow
    {
        private readonly ClipInjectorTrayContext _owner;

        public TrayWindow(ClipInjectorTrayContext owner)
        {
            _owner = owner;
            CreateHandle(new CreateParams());
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_HOTKEY)
            {
                _owner.HandleHotkey(m.WParam.ToInt32());
            }

            base.WndProc(ref m);
        }
    }

    [DllImport("user32.dll")]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll")]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    [DllImport("user32.dll")]
    private static extern IntPtr GetKeyboardLayout(uint idThread);

    [DllImport("user32.dll")]
    private static extern short VkKeyScanEx(char ch, IntPtr dwhkl);

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint type;
        public InputUnion U;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)] public MOUSEINPUT mi;
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
