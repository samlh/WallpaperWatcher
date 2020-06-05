using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

public class PixelBuffer
{
    public PixelBuffer(Bitmap bitmap)
    {
        var depth = Bitmap.GetPixelFormatSize(bitmap.PixelFormat);
        if (depth != 24 && depth != 32)
        {
            // TODO
            throw new NotImplementedException();
        }

        var bitmapData = bitmap.LockBits(new Rectangle(Point.Empty, bitmap.Size), ImageLockMode.ReadOnly, bitmap.PixelFormat);

        var stride = Math.Abs(bitmapData.Stride);
        var pixels = new byte[ystep * bitmap.Height];

        Marshal.Copy(bitmapData.Scan0, pixels, 0, pixels.Length);
        bitmap.UnlockBits(bitmapData);

        this.Size = bitmap.Size;
        this.xstep = depth / 8;
        this.ystep = stride;
        this.pixels = pixels;
    }

    byte[] pixels;
    int xstep, ystep;

    public Size Size { get; }

    public Color GetPixel(int x, int y)
    {
        var i = x * xstep + y * ystep;
        return Color.FromArgb(pixels[i + 2], pixels[i + 1], pixels[i]);
    }
}