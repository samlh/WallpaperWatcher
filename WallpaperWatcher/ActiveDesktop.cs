using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace WallpaperWatcher
{
    internal static class ActiveDesktop
    {
        private static readonly Guid CLSID_ActiveDesktop = new Guid("{75048700-EF1F-11D0-9888-006097DEACF9}");

        public static void SetDesktopWallpaper(string wallpaperLocation)
        {
            var thread = new Thread(() =>
            {
                ActiveDesktop.IActiveDesktop activeDesktop = ActiveDesktop.GetActiveDesktop();
                activeDesktop.SetWallpaper(wallpaperLocation, 0);
                activeDesktop.ApplyChanges(ActiveDesktop.ApplyFlags.Save | ActiveDesktop.ApplyFlags.Force);
            });
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            thread.Join();
        }

        public static IActiveDesktop GetActiveDesktop()
        {
            Type typeActiveDesktop = Type.GetTypeFromCLSID(CLSID_ActiveDesktop);
            return Activator.CreateInstance(typeActiveDesktop) as IActiveDesktop;
        }

        [ComImport]
        [Guid("F490EB00-1240-11D1-9888-006097DEACF9")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IActiveDesktop
        {
            [PreserveSig]
            int ApplyChanges(ApplyFlags dwFlags);

            // [PreserveSig]
            // int GetWallpaper(
            //     [MarshalAs(UnmanagedType.LPWStr)]
            //     System.Text.StringBuilder pwszWallpaper,
            //     int cchWallpaper,
            //     int dwReserved);

            [PreserveSig]
            int SetWallpaper(
                [MarshalAs(UnmanagedType.LPWStr)]
                string pwszWallpaper,
                int dwReserved);

            // [PreserveSig]
            // int GetWallpaperOptions(ref WallpaperOpt pwpo, int dwReserved);

            // [PreserveSig]
            // int SetWallpaperOptions(ref WallpaperOpt pwpo, int dwReserved);

            // [PreserveSig]
            // int GetPattern(
            //     [MarshalAs(UnmanagedType.LPWStr)]
            //     System.Text.StringBuilder pwszPattern,
            //     int cchPattern,
            //     int dwReserved);

            // [PreserveSig]
            // int SetPattern(
            //     [MarshalAs(UnmanagedType.LPWStr)]
            //     string pwszPattern,
            //     int dwReserved);
        }

        // [StructLayout(LayoutKind.Sequential)]
        // public struct WallpaperOpt
        // {
        //     public static readonly int SizeOf = Marshal.SizeOf(typeof(WallpaperOpt));
        //     public WallpaperStyle dwStyle;
        // }

        public enum WallpaperStyle : int
        {
            Center = 0,
            Tile = 1,
            Stretch = 2,
            // Windows 7
            Fit = 3,
            Fill = 4,
            // Windows 8
            Span = 5
        }

        [Flags]
        public enum ApplyFlags : int
        {
            Save = 0x00000001,
            HtmlGen = 0x00000002,
            Refresh = 0x00000004,
            All = Save | HtmlGen | Refresh,
            Force = 0x00000008,
            BufferedRefresh = 0x00000010,
            DynamicRefresh = 0x00000020
        }
    }
}

