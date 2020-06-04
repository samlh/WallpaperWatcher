using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Configuration;
using Microsoft.Win32;

namespace WallpaperWatcher
{
    public class Program
    {
        public static void Main()
        {
            var imageChangeDelay = decimal.Parse(ConfigurationManager.AppSettings["ImageChangeDelay"]);

            MiscWindowsAPIs.SetPriorityClass(Process.GetCurrentProcess().Handle,
                                             MiscWindowsAPIs.PriorityClass.PROCESS_MODE_BACKGROUND_BEGIN);

            var wallpaperChanger = new WallpaperChanger();

            var timer = new System.Timers.Timer((int)(1000 * 60 * imageChangeDelay));
            timer.Elapsed += (s, e) =>
            {
                var active = 0;
                MiscWindowsAPIs.SystemParametersInfo(
                    (int)MiscWindowsAPIs.SPI.GETSCREENSAVERRUNNING, 0, ref active, 0);
                if (active != 0)
                {
                    return;
                }

                wallpaperChanger.UpdateWallpaper();
            };

            wallpaperChanger.WallpaperChanged += (sender, e) =>
            {
                timer.Stop();
                timer.Start();
            };

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Next image", null, (s, e) => wallpaperChanger.UpdateWallpaper());
            trayMenu.Items.Add("Debug info", null, (s, e) =>
            {
                var form = new Form()
                {
                    Width = 800,
                    Height = 600,
                };
                var text = new TextBox()
                {
                    Text = wallpaperChanger.GetDebugData(),
                    Multiline = true,
                    ScrollBars = ScrollBars.Vertical,
                    ReadOnly = true,
                    Dock = DockStyle.Fill,
                };
                form.Controls.Add(text);

                void UpdateDebugText(object sender, EventArgs evd)
                {
                    text.Text = wallpaperChanger.GetDebugData();
                }
                wallpaperChanger.WallpaperChanged += UpdateDebugText;
                form.ShowDialog();
                wallpaperChanger.WallpaperChanged -= UpdateDebugText;
            });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Delete current wallpaper", null, (s, e) => wallpaperChanger.DeleteCurrentWallpaper());
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, (s, e) => Application.Exit());

            var trayIcon = new NotifyIcon()
            {
                Text = "Wallpaper Watcher",
                Icon = new Icon(typeof(Program), "Icon.ico"),
                ContextMenuStrip = trayMenu,
                Visible = true
            };
            trayIcon.DoubleClick += (s, e) => wallpaperChanger.UpdateWallpaper();

            var nextWallpaperHotkey = new Hotkey(1, Keys.N, true);
            nextWallpaperHotkey.Pressed += (sender, e) => wallpaperChanger.UpdateWallpaper();

            var deleteWallpaperHotkey = new Hotkey(2, Keys.D | Keys.Shift, true);
            deleteWallpaperHotkey.Pressed += (sender, e) => wallpaperChanger.DeleteCurrentWallpaper();

            Application.ApplicationExit += (sender, e) =>
            {
                timer.Stop();
                trayIcon.Dispose();
                nextWallpaperHotkey.Unregister();
                deleteWallpaperHotkey.Unregister();
            };
            SystemEvents.SessionSwitch += (sender, e) =>
            {
                switch (e.Reason)
                {
                    case SessionSwitchReason.ConsoleDisconnect:
                    case SessionSwitchReason.RemoteDisconnect:
                    case SessionSwitchReason.SessionLock:
                    case SessionSwitchReason.SessionLogoff:
                        timer.Stop();
                        break;

                    case SessionSwitchReason.ConsoleConnect:
                    case SessionSwitchReason.RemoteConnect:
                    case SessionSwitchReason.SessionUnlock:
                    case SessionSwitchReason.SessionLogon:
                        timer.Start();
                        break;
                }
            };

            wallpaperChanger.UpdateWallpaper();

            Application.Run();
        }
    }
}