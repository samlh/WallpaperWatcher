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
        const int ThumbSize = 120;
        const decimal ImageSideFractionToAnalyze = 0.4m;
        const decimal TiedColorDifferenceMargin = 0.08m;
        const int ColorBits = 4;

        static decimal imageChangeDelay;
        static decimal maxScaleFactor;
        static decimal skipScaleFactor;
        static decimal maxFractionOffscreen;

        static StringWriter debugData;
        static event EventHandler WallpaperChanged;

        public static void Main()
        {
            var filePathConfig = ConfigurationManager.AppSettings["ImagePath"];
            imageChangeDelay = decimal.Parse(ConfigurationManager.AppSettings["ImageChangeDelay"]);
            maxScaleFactor = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxScaleFactor"]);
            skipScaleFactor = decimal.Parse(ConfigurationManager.AppSettings["ImageSkipScaleFactor"]);
            maxFractionOffscreen = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxFractionOffscreen"]);

            MiscWindowsAPIs.SetPriorityClass(Process.GetCurrentProcess().Handle,
                                             MiscWindowsAPIs.PriorityClass.PROCESS_MODE_BACKGROUND_BEGIN);

            var fileList = GetFileList(filePathConfig.Split(Path.PathSeparator));

            var trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Next image", null, (s, e) => UpdateWallpaper(fileList));
            trayMenu.Items.Add("Debug info", null, (s, e) =>
            {
                var form = new Form()
                {
                    Width = 800,
                    Height = 600,
                };
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
                Icon = new Icon(typeof(Program), "Icon.ico"),
                ContextMenuStrip = trayMenu,
                Visible = true
            };
            trayIcon.DoubleClick += (s, e) => UpdateWallpaper(fileList);

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

                UpdateWallpaper(fileList);

                timer.Start();
            };

            var nextWallpaperHotkey = new Hotkey(1, Keys.N, true);
            nextWallpaperHotkey.Pressed += (sender, e) => UpdateWallpaper(fileList);
            Application.ApplicationExit += (sender, e) =>
            {
                trayIcon.Dispose();
                nextWallpaperHotkey.Unregister();
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

            UpdateWallpaper(fileList);

            timer.Start();
            Application.Run();
        }

        private static void UpdateWallpaper(List<string> fileList)
        {
            debugData = new StringWriter();
            while (fileList.Any())
            {
                var wallpaperLocation = fileList[new Random().Next(fileList.Count)];
                debugData.WriteLine("{0}", wallpaperLocation);

                Bitmap img;
                try
                {
                    img = new Bitmap(wallpaperLocation);
                }
                catch (Exception)
                {
                    debugData.WriteLine("File not read, skipping: {0}", wallpaperLocation);
                    fileList.Remove(wallpaperLocation);

                    continue;
                }

                DesktopWallpaperPosition? style;
                Color? bgColor = null;

                using (img)
                {
                    style = GetWallpaperStyle(img, out var matchHorz, out var matchVert);
                    if (style == null)
                    {
                        fileList.Remove(wallpaperLocation);
                        continue;
                    }

                    if (style != DesktopWallpaperPosition.Fill)
                    {
                        bgColor = GetBGColor(img, matchHorz, matchVert);
                    }
                }

                var wallpaper = (IDesktopWallpaper)new DesktopWallpaperClass();

                wallpaper.SetPosition(style.Value);

                if (bgColor.HasValue)
                {
                    wallpaper.SetBackgroundColor((uint)ColorTranslator.ToWin32(bgColor.Value));
                }

                for (uint i = 0; i < wallpaper.GetMonitorDevicePathCount(); i++)
                {
                    wallpaper.SetWallpaper(wallpaper.GetMonitorDevicePathAt(i), wallpaperLocation);
                }

                WallpaperChanged?.Invoke(null, null);

                return;
            }
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

        private static DesktopWallpaperPosition? GetWallpaperStyle(Image img, out bool matchLeftAndRight, out bool matchTopAndBottom)
        {
            var screenSize = MiscWindowsAPIs.GetScreenSize();

            var widthScaleFactor = (decimal)screenSize.Width / img.Width;
            var heightScaleFactor = (decimal)screenSize.Height / img.Height;

            var fitScaleFactor = Math.Min(widthScaleFactor, heightScaleFactor);
            var fillScaleFactor = Math.Max(widthScaleFactor, heightScaleFactor);

            var fillFractionOffscreen = 1m - Math.Min(widthScaleFactor / heightScaleFactor, heightScaleFactor / widthScaleFactor);

            debugData.WriteLine($"w={img.Width} h={img.Height} sw={screenSize.Width} sh={screenSize.Height}");
            debugData.WriteLine($"fitScaleFactor={fitScaleFactor:F2} fillScaleFactor={fillScaleFactor:F2}");
            debugData.WriteLine($"fillFractionOffscreen={fillFractionOffscreen:F2}");

            DesktopWallpaperPosition? style;

            if (fitScaleFactor > skipScaleFactor)
            {
                style = null;
            }
            else if (fitScaleFactor > maxScaleFactor)
            {
                style = DesktopWallpaperPosition.Center;
            }
            else if (fillScaleFactor > maxScaleFactor || fillFractionOffscreen > maxFractionOffscreen)
            {
                style = DesktopWallpaperPosition.Fit;
            }
            else
            {
                style = DesktopWallpaperPosition.Fill;
            }

            matchLeftAndRight = style == DesktopWallpaperPosition.Center || (style == DesktopWallpaperPosition.Fit && widthScaleFactor > heightScaleFactor);
            matchTopAndBottom = style == DesktopWallpaperPosition.Center || (style == DesktopWallpaperPosition.Fit && heightScaleFactor > widthScaleFactor);

            debugData.WriteLine($"style={style}");
            debugData.WriteLine($"matchLeftAndRight={matchLeftAndRight} matchTopAndBottom={matchTopAndBottom}");

            return style;
        }

        private static Color? GetBGColor(Image img, bool matchLeftAndRight, bool matchTopAndBottom)
        {
            Color chosen;

            using (var thumb = img.GetThumbnailImage(ThumbSize, ThumbSize, () => false, IntPtr.Zero) as Bitmap)
            {
                if (thumb == null || thumb.Width == 0 || thumb.Height == 0)
                {
                    return null;
                }

                var histogramValues = new List<Color>[1 << ColorBits << ColorBits << ColorBits];

                for (var x = 0; x < thumb.Width; x++)
                {
                    for (var y = 0; y < thumb.Width; y++)
                    {
                        if (matchLeftAndRight && OutsideCenterFraction(x, thumb.Width, ImageSideFractionToAnalyze)
                         || matchTopAndBottom && OutsideCenterFraction(y, thumb.Height, ImageSideFractionToAnalyze))
                        {
                            var c = thumb.GetPixel(x, y);
                            var b = ColorToBucket(c, ColorBits);
                            if (histogramValues[b] == null)
                            {
                                histogramValues[b] = new List<Color>();
                            }
                            histogramValues[b].Add(c);
                        }
                    }
                }

                var histogram = new int[1 << ColorBits << ColorBits << ColorBits];
                var histogramValue = new Color[1 << ColorBits << ColorBits << ColorBits];

                for (var i = 0; i < histogramValues.Length; i++)
                {
                    var bucket = histogramValues[i];
                    if (bucket != null)
                    {
                        histogram[i] = bucket.Count();
                        histogramValue[i] = bucket.GroupBy(c => ColorToBucket(c, ColorBits + 2)).OrderByDescending(g => g.Count()).First().First();
                    }
                }

                var orderedHist = histogram
                    .Select((n, i) => (n, i))
                    .Where(b => b.n != 0)
                    .OrderByDescending(b => b.n)
                    .Select(b => (b.n, c: histogramValue[b.i]))
                    .ToArray();

                if (!orderedHist.Any())
                {
                    return null;
                }

                var maxFrequency = orderedHist.First().n;
                var minFrequencyOk = (int)Math.Floor(maxFrequency * TiedColorDifferenceMargin);
                debugData.WriteLine($"maxFrequency={maxFrequency} minFrequencyOk={minFrequencyOk}");

                var tiedForBest = orderedHist
                    .TakeWhile(bucket => bucket.n >= minFrequencyOk)
                    .Select(bucket => (bucket.n, bucket.c, l: GetL(bucket.c)))
                    .OrderBy(bucket => Math.Round(bucket.l * 8))
                    .ThenBy(bucket => Math.Round(bucket.c.GetSaturation() * 8))
                    .ThenBy(bucket => bucket.l)
                    .ToArray();

                foreach (var bucket in tiedForBest)
                {
                    debugData.WriteLine($"{bucket.c.ToArgb():X} s={bucket.c.GetSaturation():F2} v={bucket.c.GetBrightness():F2} l={bucket.l:F2} {bucket.n}");
                }

                chosen = tiedForBest[0].c;

                debugData.WriteLine($"chosen: {chosen.ToArgb():X}");

            }

            return chosen;
        }

        private static bool OutsideCenterFraction(int v, int max, decimal fraction)
        {
            return Math.Abs(v - max / 2) * 2 >= max - max * fraction;
        }

        private static float GetL(Color color)
        {
            return (2 - color.GetSaturation()) * color.GetBrightness() / 2;
        }

        private static int ColorToBucket(Color c, int bits)
        {
            return (c.R >> (8 - bits) << bits << bits) | (c.G >> (8 - bits) << bits) | (c.B >> (8 - bits));
        }
    }
}


