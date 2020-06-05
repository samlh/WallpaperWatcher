using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using WallpaperWatcher;

public class WallpaperChanger
{
    const int ThumbSize = 120;
    const decimal ImageSideFractionToAnalyze = 0.4m;
    const decimal TiedColorDifferenceMargin = 0.08m;

    const int Bits = 4;
    const int BitsMask = (1 << Bits) - 1;
    const int BitsRemoved = 8 - Bits;
    const int BitsRemovedMask = (1 << BitsRemoved) - 1;

    readonly decimal maxScaleFactor;
    readonly decimal skipScaleFactor;
    readonly decimal maxFractionOffscreen;
    readonly List<string> fileList;

    string currentFile;
    StringWriter debugData;
    Stopwatch stopWatch;

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
        this.stopWatch = new Stopwatch();
    }

    public void UpdateWallpaper()
    {
        this.debugData = new StringWriter();
        this.stopWatch.Restart();
        while (fileList.Any())
        {
            var wallpaperLocation = fileList[new Random().Next(fileList.Count)];
            WriteLine($"Loading file {wallpaperLocation}");

            var (style, bgColor) = GetStyleAndBgColor(wallpaperLocation);
            if (style == null)
            {
                fileList.Remove(wallpaperLocation);
                continue;
            }

            this.currentFile = wallpaperLocation;

            var desktopWallpaper = (IDesktopWallpaper)new DesktopWallpaperClass();

            desktopWallpaper.SetPosition(style.Value);
            WriteLine("Set position");

            if (bgColor.HasValue)
            {
                desktopWallpaper.SetBackgroundColor((uint)ColorTranslator.ToWin32(bgColor.Value));
                WriteLine("Set bg color");
            }

            for (uint i = 0; i < desktopWallpaper.GetMonitorDevicePathCount(); i++)
            {
                var monitorID = desktopWallpaper.GetMonitorDevicePathAt(i);
                desktopWallpaper.SetWallpaper(monitorID, wallpaperLocation);
                WriteLine($"Set wallpaper on monitor {i}: {monitorID}");
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

    private void WriteLine(string message)
    {
        this.debugData.WriteLine($"{this.stopWatch.ElapsedMilliseconds}: {message}");
    }

    private (DesktopWallpaperPosition? style, Color? bgColor) GetStyleAndBgColor(string wallpaperLocation)
    {
        Bitmap img;
        try
        {
            // Note: this is the slowest line of the program
            img = new Bitmap(wallpaperLocation);
            WriteLine("Opened file");
        }
        catch (Exception)
        {
            WriteLine("File not read, skipping");
            return (null, null);
        }

        var style = GetWallpaperStyle(img.Size, out var matchLeftAndRight, out var matchTopAndBottom);

        PixelBuffer pixelBuffer = null;
        using (img)
        {
            if (style == null || style == DesktopWallpaperPosition.Fill)
            {
                return (style, null);
            }

            using (var thumb = img.GetThumbnailImage(ThumbSize, ThumbSize, () => false, IntPtr.Zero) as Bitmap)
            {
                WriteLine("Got thumbnail");
                pixelBuffer = new PixelBuffer(thumb);
                WriteLine("Got pixel buffer");
            }
        }

        WriteLine("Closed file");
        var bgColor = GetBGColor(pixelBuffer, matchLeftAndRight, matchTopAndBottom);
        return (style, bgColor);
    }

    private DesktopWallpaperPosition? GetWallpaperStyle(Size size, out bool matchLeftAndRight, out bool matchTopAndBottom)
    {
        var screenSize = MiscWindowsAPIs.GetScreenSize();

        var widthScaleFactor = (decimal)screenSize.Width / size.Width;
        var heightScaleFactor = (decimal)screenSize.Height / size.Height;

        var fitScaleFactor = Math.Min(widthScaleFactor, heightScaleFactor);
        var fillScaleFactor = Math.Max(widthScaleFactor, heightScaleFactor);

        var fillFractionOffscreen = 1m - Math.Min(widthScaleFactor / heightScaleFactor, heightScaleFactor / widthScaleFactor);

        this.WriteLine($"w={size.Width} h={size.Height} sw={screenSize.Width} sh={screenSize.Height}");
        this.WriteLine($"fitScaleFactor={fitScaleFactor:F2} fillScaleFactor={fillScaleFactor:F2}");
        this.WriteLine($"fillFractionOffscreen={fillFractionOffscreen:F2}");

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

        this.WriteLine($"style={style}");
        this.WriteLine($"matchLeftAndRight={matchLeftAndRight} matchTopAndBottom={matchTopAndBottom}");

        return style;
    }

    private Color? GetBGColor(PixelBuffer pixelBuffer, bool matchLeftAndRight, bool matchTopAndBottom)
    {
        var (rect1, rect2) = GetRectsToAnalyze(pixelBuffer.Size, matchLeftAndRight, matchTopAndBottom);

        var histogram = new int[1 << Bits << Bits << Bits];
        AddToHistogram(pixelBuffer, rect1, histogram);
        AddToHistogram(pixelBuffer, rect2, histogram);
        WriteLine("Computed histogram");

        var histogramBuckets = Enumerable.Range(0, histogram.Length).ToArray();
        Array.Sort(histogram, histogramBuckets, Comparer<int>.Create((x, y) => y.CompareTo(x)));

        var maxFrequency = histogram[0];
        var minFrequencyOk = (int)Math.Floor(maxFrequency * TiedColorDifferenceMargin);
        this.WriteLine($"maxFrequency={maxFrequency} minFrequencyOk={minFrequencyOk}");

        var bucketCount = Array.FindIndex(histogram, n => n < minFrequencyOk);
        if (bucketCount == -1)
        {
            bucketCount = histogram.Length;
        }

        // TODO: slow on first execution
        var tiedForBest = histogram
            .Take(bucketCount)
            .Zip(histogramBuckets, (n, bucket) =>
            {
                var c = BucketToColor(bucket);
                var s = c.GetSaturation();
                var b = c.GetBrightness();
                var l = (2 - s) * b / 2;
                return (n, c, s, b, l);
            })
            .OrderBy(bucket => Math.Round(bucket.l * 8))
            .ThenBy(bucket => Math.Round(bucket.s * 8))
            .ThenBy(bucket => bucket.l)
            .ToArray();

        foreach (var bucket in tiedForBest)
        {
            this.WriteLine($"{bucket.c.ToArgb():X} s={bucket.s:F2} v={bucket.b:F2} l={bucket.l:F2} {bucket.n}");
        }

        var chosenCoarse = tiedForBest[0].c;

        var finalColorHist = new int[1 << BitsRemoved << BitsRemoved << BitsRemoved];
        AddToSubbucketHistogram(pixelBuffer, rect1, chosenCoarse, finalColorHist);
        AddToSubbucketHistogram(pixelBuffer, rect2, chosenCoarse, finalColorHist);

        WriteLine("Computed final histogram");

        var finalColorHistBuckets = Enumerable.Range(0, finalColorHist.Length).ToArray();
        Array.Sort(finalColorHist, finalColorHistBuckets, Comparer<int>.Create((x, y) => y.CompareTo(x)));

        var adjustment = finalColorHistBuckets[0];
        var chosen = ApplySubbucketAdjustment(chosenCoarse, adjustment);

        this.WriteLine($"chosen: {chosen.ToArgb():X} (coarse {chosenCoarse.ToArgb():X})");
        return chosen;
    }

    private static (Rectangle, Rectangle) GetRectsToAnalyze(Size size, bool matchLeftAndRight, bool matchTopAndBottom)
    {
        if (matchLeftAndRight)
        {
            var w = (int)Math.Floor(size.Width * ImageSideFractionToAnalyze / 2);
            return (new Rectangle(0, 0, w, size.Height), new Rectangle(size.Width - w, 0, w, size.Height));
        }
        else if (matchTopAndBottom)
        {
            var h = (int)Math.Floor(size.Height * ImageSideFractionToAnalyze / 2);
            return (new Rectangle(0, 0, size.Width, h), new Rectangle(0, size.Height - h, size.Width, h));
        }

        return (new Rectangle(0, 0, 0, 0), new Rectangle(0, 0, 0, 0));
    }

    private static void AddToHistogram(PixelBuffer pixelBuffer, Rectangle rect, int[] histogram)
    {
        for (var y = rect.Top; y < rect.Bottom; y++)
        {
            for (var x = rect.Left; x < rect.Right; x++)
            {
                var bucket = ColorToBucket(pixelBuffer.GetPixel(x, y));
                histogram[bucket]++;
            }
        }
    }

    private static void AddToSubbucketHistogram(PixelBuffer pixelBuffer, Rectangle rect, Color coarseColor, int[] histogram)
    {
        var bucket = ColorToBucket(coarseColor);
        for (var y = rect.Top; y < rect.Bottom; y++)
        {
            for (var x = rect.Left; x < rect.Right; x++)
            {
                var color = pixelBuffer.GetPixel(x, y);
                if (bucket == ColorToBucket(color))
                {
                    var subbucket = ColorToSubbucket(color);
                    histogram[subbucket]++;
                }
            }
        }
    }

    private static int ColorToBucket(Color c)
    {
        return c.R >> BitsRemoved << Bits << Bits
             | c.G >> BitsRemoved << Bits
             | c.B >> BitsRemoved;
    }

    private static int ColorToSubbucket(Color c)
    {
        return (c.R & BitsRemovedMask) << BitsRemoved << BitsRemoved
             | (c.G & BitsRemovedMask) << BitsRemoved
             | (c.B & BitsRemovedMask);
    }

    private static Color BucketToColor(int bucket)
    {
        return Color.FromArgb(
            (bucket >> Bits >> Bits) << BitsRemoved,
            (bucket >> Bits & BitsMask) << BitsRemoved,
            (bucket & BitsMask) << BitsRemoved);
    }

    private static Color ApplySubbucketAdjustment(Color color, int adjustment)
    {
        var r = adjustment >> BitsRemoved >> BitsRemoved;
        var g = adjustment >> BitsRemoved & BitsRemovedMask;
        var b = adjustment & BitsRemovedMask;

        return Color.FromArgb(color.R + r, color.G + g, color.B + b);
    }
}