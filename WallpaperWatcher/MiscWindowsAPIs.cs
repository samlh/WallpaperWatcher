using System;
using System.Runtime.InteropServices;
using System.Drawing;
using Microsoft.Win32;

namespace WallpaperWatcher
{
    internal static class MiscWindowsAPIs
    {
        [DllImport("user32.dll")]
        private static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SystemParametersInfo(int action, int param, ref int retval, int updini);

        [DllImport("user32.dll")]
        private static extern int GetSystemMetrics(int nIndex);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool SetPriorityClass(IntPtr handle, PriorityClass priorityClass);

        public enum SPI : int
        {
            GETSCREENSAVERRUNNING = 0x0072,
        }

        public enum PriorityClass : uint
        {
            ABOVE_NORMAL_PRIORITY_CLASS = 0x8000,
            BELOW_NORMAL_PRIORITY_CLASS = 0x4000,
            HIGH_PRIORITY_CLASS = 0x80,
            IDLE_PRIORITY_CLASS = 0x40,
            NORMAL_PRIORITY_CLASS = 0x20,
            PROCESS_MODE_BACKGROUND_BEGIN = 0x100000,
            PROCESS_MODE_BACKGROUND_END = 0x200000,
            REALTIME_PRIORITY_CLASS = 0x100
        }

        public static void SetDesktopBackgroundColor(Color color)
        {
            int[] elements = { 1 };
            int[] colors = { ColorTranslator.ToWin32(color) };
            SetSysColors(1, elements, colors);
        }

        public static Size GetScreenSize()
        {
            return new Size(GetSystemMetrics(0), GetSystemMetrics(1));
        }

        public static void SetWallpaperStyleRegKeys(ActiveDesktop.WallpaperStyle style)
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\Desktop", true);

            switch (style)
            {
                case ActiveDesktop.WallpaperStyle.Center:
                    key.SetValue("WallpaperStyle", "0");
                    key.SetValue("TileWallpaper", "0");
                    break;
                case ActiveDesktop.WallpaperStyle.Tile:
                    key.SetValue("WallpaperStyle", "0");
                    key.SetValue("TileWallpaper", "1");
                    break;
                case ActiveDesktop.WallpaperStyle.Stretch:
                    key.SetValue("WallpaperStyle", "2");
                    key.SetValue("TileWallpaper", "0");
                    break;
                case ActiveDesktop.WallpaperStyle.Fit: // (Windows 7 and later)
                    key.SetValue("WallpaperStyle", "6");
                    key.SetValue("TileWallpaper", "0");
                    break;
                case ActiveDesktop.WallpaperStyle.Fill: // (Windows 7 and later)
                    key.SetValue("WallpaperStyle", "10");
                    key.SetValue("TileWallpaper", "0");
                    break;

                default:
                    throw new NotImplementedException();
            }

            key.Close();
        }

    }
}

