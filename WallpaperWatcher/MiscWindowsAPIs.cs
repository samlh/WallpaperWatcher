using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace WallpaperWatcher
{
    internal static class MiscWindowsAPIs
    {
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

        public static Size GetScreenSize()
        {
            return new Size(GetSystemMetrics(0), GetSystemMetrics(1));
        }
    }
}

