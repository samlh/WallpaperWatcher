using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Configuration;
using Microsoft.Win32;

namespace WallpaperWatcher
{

    public class Program
    {
        const int thumbSize = 120;
        const decimal imageSideFractionToAnalyze = 0.4m;
        const decimal tiedColorDifferenceMargin = 0.4m;
        const int colorBits = 3;

        static decimal imageChangeDelay;
        static decimal maxScaleFactor;
        static decimal maxFractionOffscreen;

        static System.Timers.Timer timer;
        static StringWriter debugData;
        static event EventHandler WallpaperChanged;

        public static void Main()
        {
            var filePathConfig = ConfigurationManager.AppSettings["ImagePath"];
            imageChangeDelay = decimal.Parse(ConfigurationManager.AppSettings["ImageChangeDelay"]);
            maxScaleFactor = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxScaleFactor"]);
            maxFractionOffscreen = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxFractionOffscreen"]);

            MiscWindowsAPIs.SetPriorityClass(Process.GetCurrentProcess().Handle,
                                             MiscWindowsAPIs.PriorityClass.PROCESS_MODE_BACKGROUND_BEGIN);

            var fileList = GetFileList(filePathConfig.Split(Path.PathSeparator));

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Next image", null, (s, e) => UpdateWallpaper(fileList));
            trayMenu.Items.Add("Debug info", null, (s, e) =>
                               {
                                   var form = new Form();
                                   var text = new TextBox()
                                   {
                                       Text = debugData.ToString(),
                                       Multiline = true,
                                       ScrollBars = ScrollBars.Vertical,
                                       ReadOnly = true,
                                       Dock = DockStyle.Fill,
                                   };
                                   form.Controls.Add(text);
                                   void WallpaperChangedEventHandler(object sender, EventArgs evd) => text.Text = debugData.ToString();
                                   WallpaperChanged += WallpaperChangedEventHandler;
                                   form.ShowDialog();
                                   WallpaperChanged -= WallpaperChangedEventHandler;
                               });
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, (s, e) => Application.Exit());

            var trayIcon = new NotifyIcon()
            {
                Text = "Wallpaper Watcher",
                Icon = new Icon(SystemIcons.Application, 40, 40),
                ContextMenuStrip = trayMenu,
                Visible = true
            };
            trayIcon.DoubleClick += (s, e) => UpdateWallpaper(fileList);

            timer = new System.Timers.Timer((int)(1000 * 60 * imageChangeDelay));
            timer.Elapsed += (s, e) =>
            {
                int active = 0;
                MiscWindowsAPIs.SystemParametersInfo(
                    (int)MiscWindowsAPIs.SPI.GETSCREENSAVERRUNNING, 0, ref active, 0);
                if (active != 0)
                    return;

                UpdateWallpaper(fileList);

                timer.Start();
            };

            var nextWallpaperHotkey = new Hotkey(1, Keys.N, true);
            nextWallpaperHotkey.Pressed += (sender, e) => UpdateWallpaper(fileList);
            Application.ApplicationExit += (sender, e) => nextWallpaperHotkey.Unregister();
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

            UpdateWallpaper(fileList);

            timer.Start();
            Application.Run();
        }

        private static void UpdateWallpaper(List<string> fileList)
        {
            Bitmap img;
            string wallpaperLocation;

            if (!fileList.Any())
                return;

            wallpaperLocation = fileList[new Random().Next(fileList.Count)];
            debugData = new StringWriter();
            debugData.WriteLine("{0}", wallpaperLocation);

            try
            {
                img = new Bitmap(wallpaperLocation);
            }
            catch (Exception)
            {
                Console.WriteLine("File not read, skipping: {0}", wallpaperLocation);
                fileList.Remove(wallpaperLocation);

                UpdateWallpaper(fileList);
                return;
            }

            var style = GetWallpaperStyle(img, out var matchHorz, out var matchVert);

            Color? bgColor = null;
            if (style != DesktopWallpaperPosition.Fill)
                bgColor = GetBGColor(img, matchHorz, matchVert);

            img.Dispose();

            var wallpaper = (IDesktopWallpaper)(new DesktopWallpaperClass());

            //MiscWindowsAPIs.SetWallpaperStyleRegKeys(style);
            wallpaper.SetPosition(style);

            if (bgColor.HasValue)
            {
                //MiscWindowsAPIs.SetDesktopBackgroundColor(bgColor.Value);
                wallpaper.SetBackgroundColor((uint)ColorTranslator.ToWin32(bgColor.Value));
            }

            //ActiveDesktop.SetDesktopWallpaper(wallpaperLocation);
            for (uint i = 0; i < wallpaper.GetMonitorDevicePathCount(); i++)
            {
                wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(i), wallpaperLocation);
            }

            WallpaperChanged?.Invoke(null, null);
        }

        private static List<string> GetFileList(params string[] searchDirs)
        {
            return searchDirs
                .SelectMany(d =>
                            new DirectoryInfo(d)
                            .GetFiles("*", SearchOption.AllDirectories)
                            .Select(f => f.FullName))
                .AsParallel()
                .ToList();
        }

        private static DesktopWallpaperPosition GetWallpaperStyle(Image img, out bool matchHorz, out bool matchVert)
        {
            Size screenSize = MiscWindowsAPIs.GetScreenSize();

            // The percent of the image height that would be chopped off if fit to width
            decimal scaledW = (decimal)img.Width * screenSize.Height / img.Height / screenSize.Width;

            // The scale factor for fit
            decimal fitScaleFactor = Math.Min((decimal)screenSize.Width / img.Width, (decimal)screenSize.Height / img.Height);

            // The scale factor for fill
            decimal fillScaleFactor = Math.Max((decimal)screenSize.Width / img.Width, (decimal)screenSize.Height / img.Height);

            debugData.WriteLine("{0}x{1} / {4:P}x{5:P} -> {2:P},{3:P}",
                                img.Width, img.Height,
                                scaledW, 1m / scaledW,
                                (decimal)img.Width / screenSize.Width, (decimal)img.Height / screenSize.Height);

            DesktopWallpaperPosition style;

            if (fitScaleFactor > maxScaleFactor)
            {
                debugData.WriteLine("Center");
                style = DesktopWallpaperPosition.Center;
            }
            else if (fillScaleFactor > maxScaleFactor || Math.Max(scaledW, 1m / scaledW) - 1m > maxFractionOffscreen)
            {
                debugData.WriteLine("Fit");
                style = DesktopWallpaperPosition.Fit;
            }
            else
            {
                debugData.WriteLine("Fill");
                style = DesktopWallpaperPosition.Fill;
            }

            matchHorz = style == DesktopWallpaperPosition.Center || (style == DesktopWallpaperPosition.Fit && scaledW < 1);
            matchVert = style == DesktopWallpaperPosition.Center || (style == DesktopWallpaperPosition.Fit && scaledW > 1);

            return style;
        }

        private static Color? GetBGColor(Image img, bool matchHorz, bool matchVert)
        {
            Color chosen;

            using (var thumb = img.GetThumbnailImage(thumbSize, thumbSize, () => false, IntPtr.Zero) as Bitmap)
            {
                if (thumb == null || thumb.Width == 0 || thumb.Height == 0)
                {
                    return null;
                }

                int allPixelCount = thumb.Width * thumb.Height;
                decimal imageAreaFractionToAnalyze = (1m -
                                                 (matchHorz ? (1m - imageSideFractionToAnalyze) : 1m) *
                                                 (matchVert ? (1m - imageSideFractionToAnalyze) : 1m));
                int pixelCount = (int)(allPixelCount * imageAreaFractionToAnalyze);

                //var hist = new Dictionary<Color, uint>();
                var histArray = new ushort[1 << colorBits << colorBits << colorBits];

                var orderedHist =
                    Enumerable.Range(0, thumb.Height)
                        .SelectMany(y =>
                                    Enumerable.Range(0, thumb.Width)
                                        .Where(x =>
                                            matchHorz && OutsideCenterFraction(x, thumb.Width, imageSideFractionToAnalyze) ||
                                            matchVert && OutsideCenterFraction(y, thumb.Height, imageSideFractionToAnalyze))
                                        .Select(x => thumb.GetPixel(x, y))
                                        .Select(c => Tuple.Create(ColorToBucket(c), c)))
                        .GroupBy(tuple => tuple.Item1,
                                (bucketKey, bucketedTuples) =>
                                {
                                    var commonColor = bucketedTuples
                                         .GroupBy(tuple => tuple.Item2, (c, tuples) => new Tuple<Color, int>(c, tuples.Count()))
                                         .OrderByDescending(tuple => tuple.Item2)
                                         .First()
                                         .Item1;
                                    return Tuple.Create(bucketKey, commonColor, bucketedTuples.Count());
                                })
                        .OrderByDescending(p => p.Item3)
                        .AsParallel();

                if (!orderedHist.Any())
                {
                    return null;
                }

                var minFrequency = orderedHist.First().Item3 * tiedColorDifferenceMargin;
                var tiedForBest = orderedHist.TakeWhile(p => p.Item3 >= minFrequency).ToList();

                tiedForBest = tiedForBest.OrderBy(p => p.Item2.GetBrightness()).ToList();

                foreach (var v in tiedForBest)
                {
                    debugData.WriteLine("{0:X} {1} {2:P1}", v.Item2.ToArgb(), v.Item3, (decimal)v.Item3 / pixelCount);
                }

                chosen = tiedForBest.First().Item2;

                debugData.WriteLine("chosen: {0:X}", chosen.ToArgb());

            }

            return chosen;
        }

        private static bool OutsideCenterFraction(int v, int max, decimal fraction)
        {
            return Math.Abs(v - max / 2) * 2 >= max - max * fraction;
        }

        private static int ColorToBucket(Color c)
        {
            return c.R >> 8 - colorBits << colorBits << colorBits | c.G >> 8 - colorBits << colorBits | c.B >> 8 - colorBits;
        }
    }
}


