using System;
using System.Drawing;
using System.Globalization;

namespace VerticalSheetTabs
{
    // The part of the wallpaper picture that is used, as fractions (0 to 1) of its width and height: left, top, right, bottom.
    // The whole picture (0,0,1,1) is "no cropping". A value that cannot be a part of the picture (the default value of the struct,
    // a right edge that is not right of the left edge, and so on) counts as "no cropping" too, so that a wallpaper never vanishes
    // because a value was left out.
    public struct WallpaperCrop
    {
        // The smallest width and height of the part (a fraction of the picture).
        public const double MinSide = 0.02;

        public readonly double L, T, R, B;

        public WallpaperCrop(double l, double t, double r, double b)
        {
            L = l;
            T = t;
            R = r;
            B = b;
        }

        public static WallpaperCrop Full { get { return new WallpaperCrop(0, 0, 1, 1); } }

        // Fits the values into the picture (0 to 1), rounds them to four decimals (what is saved and what is shown are the same
        // values) and returns "no cropping" for a part that is too small or not a part at all.
        public WallpaperCrop Normalize()
        {
            if (IsBad(L) || IsBad(T) || IsBad(R) || IsBad(B)) return Full;
            double l = Round(L), t = Round(T), r = Round(R), b = Round(B);
            if (r - l < MinSide - 1e-9 || b - t < MinSide - 1e-9) return Full;
            return new WallpaperCrop(l, t, r, b);
        }

        public bool IsFull
        {
            get
            {
                WallpaperCrop n = Normalize();
                return n.L == 0 && n.T == 0 && n.R == 1 && n.B == 1;
            }
        }

        public bool SameAs(WallpaperCrop o)
        {
            WallpaperCrop a = Normalize(), b = o.Normalize();
            return a.L == b.L && a.T == b.T && a.R == b.R && a.B == b.B;
        }

        // The pixels of a picture of w x h that the part covers. Rounded outwards; at least one pixel; never outside the picture.
        public Rectangle ToPixels(int w, int h)
        {
            w = Math.Max(1, w);
            h = Math.Max(1, h);
            WallpaperCrop n = Normalize();
            int x0 = Math.Max(0, Math.Min(w - 1, (int)Math.Floor(n.L * w + 1e-9)));
            int y0 = Math.Max(0, Math.Min(h - 1, (int)Math.Floor(n.T * h + 1e-9)));
            int x1 = Math.Max(x0 + 1, Math.Min(w, (int)Math.Ceiling(n.R * w - 1e-9)));
            int y1 = Math.Max(y0 + 1, Math.Min(h, (int)Math.Ceiling(n.B * h - 1e-9)));
            return new Rectangle(x0, y0, x1 - x0, y1 - y0);
        }

        // "L,T,R,B": the point is a period, at most four decimals (the same in every culture).
        public string Format()
        {
            WallpaperCrop n = Normalize();
            return n.L.ToString("0.####", CultureInfo.InvariantCulture) + "," + n.T.ToString("0.####", CultureInfo.InvariantCulture) + ","
                + n.R.ToString("0.####", CultureInfo.InvariantCulture) + "," + n.B.ToString("0.####", CultureInfo.InvariantCulture);
        }

        // Reads "L,T,R,B". Anything that cannot be read (not four numbers) is "no cropping".
        public static WallpaperCrop Parse(string text)
        {
            WallpaperCrop c;
            return TryParse(text, out c) ? c.Normalize() : Full;
        }

        public static bool TryParse(string text, out WallpaperCrop crop)
        {
            crop = Full;
            if (string.IsNullOrWhiteSpace(text)) return false;
            string[] parts = text.Split(',');
            if (parts.Length != 4) return false;
            var v = new double[4];
            for (int i = 0; i < 4; i++)
            {
                if (!double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v[i])) return false;
            }
            crop = new WallpaperCrop(v[0], v[1], v[2], v[3]);
            return true;
        }

        private static bool IsBad(double v) { return double.IsNaN(v) || double.IsInfinity(v); }

        private static double Round(double v)
        {
            return Math.Round(Math.Max(0, Math.Min(1, v)), 4, MidpointRounding.AwayFromZero);
        }
    }

    // What the pointer holds of the cropping frame.
    public enum CropHandle { None, Move, N, S, W, E, NW, NE, SW, SE, Rotate }

    // What the pointer does to the cropping frame (pure arithmetic in the pixels of the preview: no drawing).
    public static class CropDrag
    {
        // What is under the point p: a corner, an edge, the dotted lines through the middle of the frame (the picture is turned there),
        // the inside, or nothing. grab = how far from the line of the frame the frame
        // can be held; it is made narrower for a small frame, so that the inside can still be held. canMove = false when the frame
        // fills the whole picture (the inside cannot be moved).
        public static CropHandle HitTest(RectangleF frame, PointF p, float grab, bool canMove)
        {
            float g = Math.Max(1f, Math.Min(grab, Math.Min(frame.Width, frame.Height) / 3f));
            bool inX = p.X >= frame.Left - g && p.X <= frame.Right + g;
            bool inY = p.Y >= frame.Top - g && p.Y <= frame.Bottom + g;
            if (!inX || !inY) return CropHandle.None;
            bool nearL = Math.Abs(p.X - frame.Left) <= g, nearR = Math.Abs(p.X - frame.Right) <= g;
            bool nearT = Math.Abs(p.Y - frame.Top) <= g, nearB = Math.Abs(p.Y - frame.Bottom) <= g;
            if (nearT && nearL) return CropHandle.NW;
            if (nearT && nearR) return CropHandle.NE;
            if (nearB && nearL) return CropHandle.SW;
            if (nearB && nearR) return CropHandle.SE;
            if (nearT) return CropHandle.N;
            if (nearB) return CropHandle.S;
            if (nearL) return CropHandle.W;
            if (nearR) return CropHandle.E;
            if (frame.Contains(p))
            {
                float mx = frame.Left + frame.Width / 2f, my = frame.Top + frame.Height / 2f;
                if (Math.Abs(p.X - mx) <= g || Math.Abs(p.Y - my) <= g) return CropHandle.Rotate;
            }
            if (canMove && frame.Contains(p)) return CropHandle.Move;
            return CropHandle.None;
        }

        // The frame after the pointer went from startPt to pt while it held the handle h of the frame start. The frame stays in
        // bounds (the picture), is at least minW x minH (but never more than half of the bounds) and does not turn inside out.
        public static RectangleF Drag(CropHandle h, RectangleF start, PointF startPt, PointF pt, RectangleF bounds, float minW, float minH)
        {
            float dx = pt.X - startPt.X, dy = pt.Y - startPt.Y;
            minW = Math.Min(minW, bounds.Width / 2f);
            minH = Math.Min(minH, bounds.Height / 2f);
            float l = start.Left, t = start.Top, r = start.Right, b = start.Bottom;
            if (h == CropHandle.Move)
            {
                float w = start.Width, hh = start.Height;
                l = Math.Max(bounds.Left, Math.Min(bounds.Right - w, start.Left + dx));
                t = Math.Max(bounds.Top, Math.Min(bounds.Bottom - hh, start.Top + dy));
                return new RectangleF(l, t, w, hh);
            }
            if (h == CropHandle.W || h == CropHandle.NW || h == CropHandle.SW) l = Math.Max(bounds.Left, Math.Min(r - minW, start.Left + dx));
            if (h == CropHandle.E || h == CropHandle.NE || h == CropHandle.SE) r = Math.Min(bounds.Right, Math.Max(l + minW, start.Right + dx));
            if (h == CropHandle.N || h == CropHandle.NW || h == CropHandle.NE) t = Math.Max(bounds.Top, Math.Min(b - minH, start.Top + dy));
            if (h == CropHandle.S || h == CropHandle.SW || h == CropHandle.SE) b = Math.Min(bounds.Bottom, Math.Max(t + minH, start.Bottom + dy));
            return new RectangleF(l, t, r - l, b - t);
        }
    }

    // Where the picture lies in the preview, and the frame in pixels of the preview (pure arithmetic).
    public static class CropPreviewMath
    {
        // The largest rectangle of the ratio of the picture that fits into a client area of clientW x clientH minus the margin on every
        // side, in the middle. A small picture is enlarged. Empty when there is nothing to show.
        public static RectangleF FitRect(int clientW, int clientH, int imgW, int imgH, int margin)
        {
            if (imgW < 1 || imgH < 1) return RectangleF.Empty;
            float availW = clientW - 2f * margin, availH = clientH - 2f * margin;
            if (availW < 1f || availH < 1f) return RectangleF.Empty;
            float s = Math.Min(availW / imgW, availH / imgH);
            float w = imgW * s, h = imgH * s;
            return new RectangleF((clientW - w) / 2f, (clientH - h) / 2f, w, h);
        }

        public static RectangleF ToFrame(WallpaperCrop crop, RectangleF bounds)
        {
            WallpaperCrop n = crop.Normalize();
            return new RectangleF(bounds.Left + (float)n.L * bounds.Width, bounds.Top + (float)n.T * bounds.Height,
                (float)(n.R - n.L) * bounds.Width, (float)(n.B - n.T) * bounds.Height);
        }

        public static WallpaperCrop FromFrame(RectangleF frame, RectangleF bounds)
        {
            if (bounds.Width <= 0 || bounds.Height <= 0) return WallpaperCrop.Full;
            return new WallpaperCrop((frame.Left - bounds.Left) / bounds.Width, (frame.Top - bounds.Top) / bounds.Height,
                (frame.Right - bounds.Left) / bounds.Width, (frame.Bottom - bounds.Top) / bounds.Height).Normalize();
        }
    }
}
