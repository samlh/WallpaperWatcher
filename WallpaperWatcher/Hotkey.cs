using System;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace WallpaperWatcher
{
    // Based on public domain code from http://bloggablea.wordpress.com/2007/05/01/global-hotkeys-with-net/
    internal class Hotkey : IMessageFilter
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, Keys vk);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int UnregisterHotKey(IntPtr hWnd, int id);

        private const uint WM_HOTKEY = 0x312;

        private const uint MOD_ALT = 0x1;
        private const uint MOD_CONTROL = 0x2;
        private const uint MOD_SHIFT = 0x4;
        private const uint MOD_WIN = 0x8;

        private const uint ERROR_HOTKEY_ALREADY_REGISTERED = 1409;
        private const uint ERROR_HOTKEY_NOT_REGISTERED = 1419;

        private readonly Keys keyCode;
        private readonly uint modifiers;
        private readonly int id;
        private bool registered;

        public event HandledEventHandler Pressed;

        public Hotkey(int id, Keys keyCode, bool windowsKeyModifier = false)
        {
            this.keyCode = keyCode & Keys.KeyCode;

            this.id = id;
            this.modifiers =
                (keyCode.HasFlag(Keys.Alt) ? Hotkey.MOD_ALT : 0) |
                (keyCode.HasFlag(Keys.Control) ? Hotkey.MOD_CONTROL : 0) |
                (keyCode.HasFlag(Keys.Shift) ? Hotkey.MOD_SHIFT : 0) |
                (windowsKeyModifier ? Hotkey.MOD_WIN : 0);
            this.registered = false;

            // Register us as a message filter
            this.Register();
            Application.AddMessageFilter(this);
        }

        ~Hotkey()
        {
            // Unregister the hotkey if necessary
            if (this.Registered)
            {
                this.Unregister();
            }
        }

        public bool Register()
        {
            // Register the hotkey
            if (Hotkey.RegisterHotKey(IntPtr.Zero, this.id, this.modifiers, this.keyCode) == 0)
            {
                // Is the error that the hotkey is registered?
                if (Marshal.GetLastWin32Error() == ERROR_HOTKEY_ALREADY_REGISTERED)
                {
                    return false;
                }
                else
                {
                    throw new Win32Exception();
                }
            }

            // We successfully registered
            this.registered = true;
            return true;
        }

        public bool Unregister()
        {
            // Clean up after ourselves
            if (Hotkey.UnregisterHotKey(IntPtr.Zero, this.id) == 0)
            {
                if (Marshal.GetLastWin32Error() == ERROR_HOTKEY_NOT_REGISTERED)
                {
                    this.registered = false;
                    return false;
                }
                else
                {
                    throw new Win32Exception();
                }
            }
            this.registered = false;
            return true;
        }

        public bool PreFilterMessage(ref Message message)
        {
            // Only process WM_HOTKEY messages
            if (message.Msg != Hotkey.WM_HOTKEY)
            {
                return false;
            }

            // Check that the ID is our key and we are registerd
            if (this.registered && (message.WParam.ToInt32() == this.id))
            {
                // Fire the event and pass on the event if our handlers didn't handle it
                return this.OnPressed();
            }
            else
            {
                return false;
            }
        }

        private bool OnPressed()
        {
            // Fire the event if we can
            var handledEventArgs = new HandledEventArgs(false);
            this.Pressed?.Invoke(this, handledEventArgs);

            // Return whether we handled the event or not
            return handledEventArgs.Handled;
        }

        public bool Registered
        {
            get { return this.registered; }
        }
    }

}

