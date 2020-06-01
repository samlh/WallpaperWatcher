using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using WallpaperWatcher;

class WallpaperChanger
{
    const int ThumbSize = 120;
    const decimal ImageSideFractionToAnalyze = 0.4m;
    const decimal TiedColorDifferenceMargin = 0.08m;
    const int ColorBits = 4;

    readonly decimal maxScaleFactor;
    readonly decimal skipScaleFactor;
    readonly decimal maxFractionOffscreen;
    readonly List<string> fileList;

    string currentFile;
    StringWriter debugData;

    public event EventHandler WallpaperChanged;

    public WallpaperChanger()
    {
        this.maxScaleFactor = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxScaleFactor"]);
        this.skipScaleFactor = decimal.Parse(ConfigurationManager.AppSettings["ImageSkipScaleFactor"]);
        this.maxFractionOffscreen = decimal.Parse(ConfigurationManager.AppSettings["ImageMaxFractionOffscreen"]);

        var filePathConfig = ConfigurationManager.AppSettings["ImagePath"];
        this.fileList = filePathConfig
            .Split(Path.PathSeparator)
            .SelectMany(d =>
                        new DirectoryInfo(d)
                        .GetFiles("*", SearchOption.AllDirectories)
                        .Select(f => f.FullName))
            .AsParallel()
            .ToList();
    }

    public void UpdateWallpaper()
    {
        this.debugData = new StringWriter();
        while (fileList.Any())
        {
            var wallpaperLocation = fileList[new Random().Next(fileList.Count)];
            this.debugData.WriteLine("{0}", wallpaperLocation);

            Bitmap img;
            try
            {
                img = new Bitmap(wallpaperLocation);
            }
            catch (Exception)
            {
                this.debugData.WriteLine("File not read, skipping: {0}", wallpaperLocation);
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

            this.currentFile = wallpaperLocation;

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

    public void DeleteCurrentWallpaper()
    {
        fileList.Remove(this.currentFile);
        File.Delete(this.currentFile);
        this.UpdateWallpaper();
    }

    public string GetDebugData()
    {
        return this.debugData.ToString();
    }

    private DesktopWallpaperPosition? GetWallpaperStyle(Image img, out bool matchLeftAndRight, out bool matchTopAndBottom)
    {
        var screenSize = MiscWindowsAPIs.GetScreenSize();

        var widthScaleFactor = (decimal)screenSize.Width / img.Width;
        var heightScaleFactor = (decimal)screenSize.Height / img.Height;

        var fitScaleFactor = Math.Min(widthScaleFactor, heightScaleFactor);
        var fillScaleFactor = Math.Max(widthScaleFactor, heightScaleFactor);

        var fillFractionOffscreen = 1m - Math.Min(widthScaleFactor / heightScaleFactor, heightScaleFactor / widthScaleFactor);

        this.debugData.WriteLine($"w={img.Width} h={img.Height} sw={screenSize.Width} sh={screenSize.Height}");
        this.debugData.WriteLine($"fitScaleFactor={fitScaleFactor:F2} fillScaleFactor={fillScaleFactor:F2}");
        this.debugData.WriteLine($"fillFractionOffscreen={fillFractionOffscreen:F2}");

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

        this.debugData.WriteLine($"style={style}");
        this.debugData.WriteLine($"matchLeftAndRight={matchLeftAndRight} matchTopAndBottom={matchTopAndBottom}");

        return style;
    }

    private Color? GetBGColor(Image img, bool matchLeftAndRight, bool matchTopAndBottom)
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
            this.debugData.WriteLine($"maxFrequency={maxFrequency} minFrequencyOk={minFrequencyOk}");

            var tiedForBest = orderedHist
                .TakeWhile(bucket => bucket.n >= minFrequencyOk)
                .Select(bucket => (bucket.n, bucket.c, l: GetL(bucket.c)))
                .OrderBy(bucket => Math.Round(bucket.l * 8))
                .ThenBy(bucket => Math.Round(bucket.c.GetSaturation() * 8))
                .ThenBy(bucket => bucket.l)
                .ToArray();

            foreach (var bucket in tiedForBest)
            {
                this.debugData.WriteLine($"{bucket.c.ToArgb():X} s={bucket.c.GetSaturation():F2} v={bucket.c.GetBrightness():F2} l={bucket.l:F2} {bucket.n}");
            }

            chosen = tiedForBest[0].c;

            this.debugData.WriteLine($"chosen: {chosen.ToArgb():X}");
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