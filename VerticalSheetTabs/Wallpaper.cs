using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;

namespace VerticalSheetTabs
{
    // How the picture is fitted into the background area.
    public enum WallpaperFit
    {
        Original,  // the pixels of the picture as they are
        FitWidth,  // as wide as the area (the ratio is kept)
        FitHeight, // as high as the area (the ratio is kept)
        FitBoth    // as wide and as high as the area: stretched, or (KeepRatio) enlarged until it covers the area
    }

    public enum WallpaperHAlign { Left, Center, Right }

    public enum WallpaperVAlign { Top, Center, Bottom }

    // The wallpaper settings of the vertical tabs (all kept for the whole PC).
    public class WallpaperOptions
    {
        public string Path = "";
        public WallpaperFit Fit = WallpaperFit.FitBoth;
        public bool KeepRatio = true;   // FitBoth only: cover the area without changing the ratio
        public bool Tile;               // repeat the picture (not with FitBoth)
        public WallpaperHAlign HAlign = WallpaperHAlign.Center;
        public WallpaperVAlign VAlign = WallpaperVAlign.Center;
        public WallpaperCrop Crop = WallpaperCrop.Full; // the part of the picture that is used (the rest is left out)
        public WallpaperRotation Rotation;              // how the picture is turned before the part is cut out (default: not turned)

        public bool SameAs(WallpaperOptions o)
        {
            return o != null && Path == o.Path && Fit == o.Fit && KeepRatio == o.KeepRatio && Tile == o.Tile && HAlign == o.HAlign && VAlign == o.VAlign && Crop.SameAs(o.Crop) && Rotation.SameAs(o.Rotation);
        }

        public WallpaperOptions Clone()
        {
            return (WallpaperOptions)MemberwiseClone();
        }

        // Which of the settings the chosen way of fitting leaves to the user (the dialog disables the others).
        public static bool TileApplies(WallpaperFit fit) { return fit != WallpaperFit.FitBoth; }
        public static bool KeepRatioApplies(WallpaperFit fit) { return fit == WallpaperFit.FitBoth; }
        public static bool HAlignApplies(WallpaperFit fit, bool keepRatio) { return fit == WallpaperFit.Original || fit == WallpaperFit.FitHeight || (fit == WallpaperFit.FitBoth && keepRatio); }
        public static bool VAlignApplies(WallpaperFit fit, bool keepRatio) { return fit == WallpaperFit.Original || fit == WallpaperFit.FitWidth || (fit == WallpaperFit.FitBoth && keepRatio); }
    }

    // The size and place of the picture (or of one tile of it) in an area of areaW x areaH pixels.
    public struct WallpaperLayout
    {
        public int Width, Height; // the size of the picture (a tile)
        public int X, Y;          // where its top left corner is (it can be negative: the picture is cut)
        public bool Tile;
    }

    // The transparency of the rows: 0 to 90 percent, in steps of 5.
    public static class RowTransparencyScale
    {
        public const int Max = 90, Step = 5;

        public static int Clamp(int v)
        {
            v = Math.Max(0, Math.Min(Max, v));
            return v / Step * Step;
        }
    }

    public static class Wallpaper
    {
        // Pictures larger than this (the longer side, pixels) are made smaller when they are read, to keep the memory in check.
        public const int MaxSide = 2560;
        private const long MaxFileBytes = 100L * 1024 * 1024;

        public static bool IsSupportedFile(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".jpg" || ext == ".jpeg" || ext == ".png";
        }

        // Reads the picture (jpg or png). Returns null for anything that cannot be read (no file, another kind of file, damaged
        // data, too large): the vertical tabs then have no wallpaper. The file is not kept open.
        public static Bitmap Load(string path)
        {
            try
            {
                if (!IsSupportedFile(path)) return null;
                var info = new FileInfo(path);
                if (!info.Exists || info.Length <= 0 || info.Length > MaxFileBytes) return null;
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (Image image = Image.FromStream(stream, false, true))
                {
                    int w = image.Width, h = image.Height;
                    if (w < 1 || h < 1) return null;
                    double k = Math.Min(1.0, MaxSide / (double)Math.Max(w, h));
                    int nw = Math.Max(1, (int)Math.Round(w * k)), nh = Math.Max(1, (int)Math.Round(h * k));
                    var bmp = new Bitmap(nw, nh, PixelFormat.Format32bppPArgb);
                    try
                    {
                        using (Graphics g = Graphics.FromImage(bmp))
                        {
                            g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                            g.CompositingMode = CompositingMode.SourceOver;
                            g.DrawImage(image, new Rectangle(0, 0, nw, nh), 0, 0, w, h, GraphicsUnit.Pixel);
                        }
                        return bmp;
                    }
                    catch (Exception)
                    {
                        bmp.Dispose();
                        return null;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static int Place(int area, int size, int align)
        {
            if (align == 0) return 0;
            int diff = area - size;
            return align == 1 ? (int)Math.Floor(diff / 2.0) : diff;
        }

        // Works out the size and the place of the picture for the options (pure arithmetic: no drawing).
        public static WallpaperLayout Layout(int srcW, int srcH, int areaW, int areaH, WallpaperOptions o)
        {
            var r = new WallpaperLayout();
            int hAlign = (int)o.HAlign, vAlign = (int)o.VAlign;
            srcW = Math.Max(1, srcW);
            srcH = Math.Max(1, srcH);
            areaW = Math.Max(1, areaW);
            areaH = Math.Max(1, areaH);
            bool placeX = true, placeY = true;
            switch (o.Fit)
            {
                case WallpaperFit.Original:
                    r.Width = srcW;
                    r.Height = srcH;
                    break;
                case WallpaperFit.FitWidth:
                    r.Width = areaW;
                    r.Height = Math.Max(1, (int)Math.Round(srcH * (double)areaW / srcW));
                    placeX = false;
                    break;
                case WallpaperFit.FitHeight:
                    r.Height = areaH;
                    r.Width = Math.Max(1, (int)Math.Round(srcW * (double)areaH / srcH));
                    placeY = false;
                    break;
                default:
                    if (o.KeepRatio)
                    {
                        double k = Math.Max(areaW / (double)srcW, areaH / (double)srcH);
                        r.Width = Math.Max(areaW, (int)Math.Round(srcW * k));
                        r.Height = Math.Max(areaH, (int)Math.Round(srcH * k));
                    }
                    else
                    {
                        r.Width = areaW;
                        r.Height = areaH;
                        placeX = false;
                        placeY = false;
                    }
                    break;
            }
            r.X = placeX ? Place(areaW, r.Width, hAlign) : 0;
            r.Y = placeY ? Place(areaH, r.Height, vAlign) : 0;
            r.Tile = o.Tile && WallpaperOptions.TileApplies(o.Fit);
            return r;
        }

        // A part cut out of a picture is drawn with its edge pixels mirrored (see Render). Null when the whole picture is drawn.
        private static ImageAttributes EdgeAttributes(bool cropped)
        {
            if (!cropped) return null;
            var attrs = new ImageAttributes();
            attrs.SetWrapMode(WrapMode.TileFlipXY);
            return attrs;
        }

        // The part of the picture that is used when it is turned: the frame (src, in the pixels of the picture, as the preview shows it)
        // lies on the turned picture that was made smaller by the zoom, so it is a part (1 / zoom as large, round the middle of the
        // picture) of the picture turned as it is. Returns that part's place; the picture is not made smaller for the wallpaper.
        private static RectangleF TurnedPartPlace(Rectangle src, Size pic, WallpaperRotation n)
        {
            double k = n.Zoom <= 0 ? 1.0 : n.Zoom;
            double cx = pic.Width / 2.0, cy = pic.Height / 2.0;
            return new RectangleF((float)(cx + (src.X - cx) / k), (float)(cy + (src.Y - cy) / k), (float)(src.Width / k), (float)(src.Height / k));
        }

        // The part place of the picture after it is turned round its middle (see WallpaperRotation), cut out of the result into a
        // bitmap of the size size. Null when it cannot be drawn.
        private static Bitmap RotatedPart(Bitmap source, RectangleF place, Size size, WallpaperRotation rotation)
        {
            WallpaperRotation n = rotation.Normalize();
            if (size.Width < 1 || size.Height < 1 || (long)size.Width * size.Height > 64L * 1024 * 1024) return null;
            var part = new Bitmap(size.Width, size.Height, PixelFormat.Format32bppPArgb);
            try
            {
                using (Graphics g = Graphics.FromImage(part))
                using (var m = new Matrix())
                using (var attrs = new ImageAttributes())
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    attrs.SetWrapMode(WrapMode.TileFlipXY);
                    m.RotateAt((float)n.Degrees, new PointF(source.Width / 2f, source.Height / 2f));
                    m.Translate(-place.X, -place.Y, MatrixOrder.Append);
                    g.Transform = m;
                    g.DrawImage(source, new Rectangle(0, 0, source.Width, source.Height), 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, attrs);
                }
                return part;
            }
            catch (Exception)
            {
                part.Dispose();
                return null;
            }
        }

        // Draws the wallpaper into a transparent bitmap of the size of the area. The caller draws it where the area is.
        public static Bitmap Render(Bitmap source, int areaW, int areaH, WallpaperOptions o)
        {
            if (source == null || areaW < 1 || areaH < 1) return null;
            // The part of the picture that is used: the layout works with its size, the drawing takes its pixels.
            Rectangle src = o.Crop.ToPixels(source.Width, source.Height);
            bool rotated = !o.Rotation.IsIdentity;
            bool cropped = rotated || src.X != 0 || src.Y != 0 || src.Width != source.Width || src.Height != source.Height;
            // (turned: the part is cut from the turned picture, and its size can differ from the frame's by the zoom)
            WallpaperRotation turn = o.Rotation.Normalize();
            RectangleF place = rotated ? TurnedPartPlace(src, source.Size, turn) : RectangleF.Empty;
            Size used = rotated ? new Size((int)Math.Ceiling(place.Width - 1e-3), (int)Math.Ceiling(place.Height - 1e-3)) : src.Size;
            if (used.Width < 1 || used.Height < 1) return null;
            WallpaperLayout lay = Layout(used.Width, used.Height, areaW, areaH, o);
            // The smoothing looks at the neighbours of a pixel, even outside the part that is drawn: the part is cut out into a bitmap
            // of its own (drawn whole, with its edge pixels mirrored), so that nothing from outside it bleeds into its edges.
            Bitmap picture = source;
            bool ownPicture = false;
            if (cropped)
            {
                try
                {
                    picture = rotated ? RotatedPart(source, place, used, turn) : source.Clone(src, PixelFormat.Format32bppPArgb);
                    if (picture == null) return null;
                    ownPicture = true;
                }
                catch (Exception)
                {
                    return null;
                }
            }
            var bmp = new Bitmap(areaW, areaH, PixelFormat.Format32bppPArgb);
            try
            {
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.Clear(Color.Transparent);
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    if (lay.Tile && (long)lay.Width * lay.Height <= 8L * 1024 * 1024)
                    {
                        Bitmap tile = picture;
                        bool own = false;
                        if (lay.Width != picture.Width || lay.Height != picture.Height)
                        {
                            tile = new Bitmap(lay.Width, lay.Height, PixelFormat.Format32bppPArgb);
                            own = true;
                            using (Graphics tg = Graphics.FromImage(tile))
                            using (ImageAttributes attrs = EdgeAttributes(cropped))
                            {
                                tg.InterpolationMode = InterpolationMode.HighQualityBicubic;
                                tg.PixelOffsetMode = PixelOffsetMode.HighQuality;
                                tg.DrawImage(picture, new Rectangle(0, 0, lay.Width, lay.Height), 0, 0, picture.Width, picture.Height, GraphicsUnit.Pixel, attrs);
                            }
                        }
                        try
                        {
                            using (var brush = new TextureBrush(tile, WrapMode.Tile))
                            {
                                brush.TranslateTransform(lay.X, lay.Y);
                                g.FillRectangle(brush, 0, 0, areaW, areaH);
                            }
                        }
                        finally
                        {
                            if (own) tile.Dispose();
                        }
                    }
                    else
                    {
                        using (ImageAttributes attrs = EdgeAttributes(cropped))
                        {
                            g.DrawImage(picture, new Rectangle(lay.X, lay.Y, lay.Width, lay.Height), 0, 0, picture.Width, picture.Height, GraphicsUnit.Pixel, attrs);
                        }
                    }
                }
                return bmp;
            }
            catch (Exception)
            {
                bmp.Dispose();
                return null;
            }
            finally
            {
                if (ownPicture) picture.Dispose();
            }
        }
    }
}
