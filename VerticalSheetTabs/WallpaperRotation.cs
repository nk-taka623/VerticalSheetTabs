using System;
using System.Drawing;
using System.Globalization;

namespace VerticalSheetTabs
{
    // How the picture of the wallpaper is turned in the preview of the settings dialog: the angle (clockwise is positive, in degrees),
    // and how much the turned picture is made smaller to stay inside the preview (Zoom, 1 = not smaller). The picture is always turned
    // round its own middle, which is the middle of the preview. The zoom follows from the angle and the size of the preview; it is kept
    // as well so that the wallpaper is cut from the picture as it was seen. Not turned (angle 0, zoom 1) is the default of the struct.
    public struct WallpaperRotation
    {
        public readonly double Degrees, Zoom;

        public WallpaperRotation(double degrees, double zoom)
        {
            Degrees = degrees;
            Zoom = zoom;
        }

        public static WallpaperRotation None { get { return new WallpaperRotation(0, 1); } }

        // The smallest zoom that is kept (a picture much longer than wide would otherwise be made too small to see).
        public const double MinZoom = 0.1;

        // The angle in the range -180 (excluded) to 180 (included).
        public static double WrapDegrees(double d)
        {
            d = d % 360.0;
            if (d > 180.0) d -= 360.0;
            if (d <= -180.0) d += 360.0;
            return d;
        }

        // What is saved: the angle in two decimals, the zoom (MinZoom to 1) in four. Anything that is not a number is "not turned"; a
        // zoom that is not given (0, the default of the struct) is 1.
        public WallpaperRotation Normalize()
        {
            if (IsBad(Degrees) || IsBad(Zoom)) return None;
            double a = Math.Round(WrapDegrees(Degrees), 2, MidpointRounding.AwayFromZero);
            if (a <= -180.0) a = 180.0;
            double z = Zoom <= 0 ? 1.0 : Math.Round(Math.Max(MinZoom, Math.Min(1.0, Zoom)), 4, MidpointRounding.AwayFromZero);
            return new WallpaperRotation(a + 0.0, z);
        }

        // The same turn on the picture turned over left to right (the angle goes the other way).
        public WallpaperRotation Mirror()
        {
            WallpaperRotation n = Normalize();
            return new WallpaperRotation(-n.Degrees, n.Zoom).Normalize();
        }

        public bool IsIdentity
        {
            get
            {
                WallpaperRotation n = Normalize();
                return n.Degrees == 0 && n.Zoom == 1.0;
            }
        }

        public bool SameAs(WallpaperRotation o)
        {
            WallpaperRotation a = Normalize(), b = o.Normalize();
            return a.Degrees == b.Degrees && a.Zoom == b.Zoom;
        }

        // "Degrees,Zoom": the point is a period (the same in every culture).
        public string Format()
        {
            WallpaperRotation n = Normalize();
            return n.Degrees.ToString("0.##", CultureInfo.InvariantCulture) + "," + n.Zoom.ToString("0.####", CultureInfo.InvariantCulture);
        }

        // Reads "Degrees,Zoom" (or "Degrees" alone: zoom 1). Anything that cannot be read is "not turned".
        public static WallpaperRotation Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return None;
            string[] parts = text.Split(',');
            if (parts.Length < 1 || parts.Length > 2) return None;
            var v = new double[] { 0, 1 };
            for (int i = 0; i < parts.Length; i++)
            {
                if (!double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v[i])) return None;
            }
            return new WallpaperRotation(v[0], v[1]).Normalize();
        }

        private static bool IsBad(double v) { return double.IsNaN(v) || double.IsInfinity(v); }
    }

    // The arithmetic of turning the picture in the preview (pure: no drawing). Coordinates are the pixels of the preview; "bounds" is
    // where the not turned picture lies (CropPreviewMath.FitRect), which has the ratio of the picture.
    public static class RotationMath
    {
        private static PointF Turn(double x, double y, double degrees)
        {
            double a = degrees * Math.PI / 180.0, c = Math.Cos(a), s = Math.Sin(a);
            return new PointF((float)(x * c - y * s), (float)(x * s + y * c));
        }

        // The middle of the preview (the picture is turned round it).
        public static PointF Center(RectangleF bounds)
        {
            return new PointF(bounds.Left + bounds.Width / 2f, bounds.Top + bounds.Height / 2f);
        }

        // How much the picture (w x h, turned by degrees round its middle) is made smaller so that its four corners stay inside an area
        // of the same middle (areaW x areaH): 1 when it fits as it is. It follows from the angle alone (it is the same again after 180
        // degrees), so a picture that was made smaller grows again as the angle goes on, and is made smaller again when it would stick out.
        public static double FitZoom(double degrees, double w, double h, double areaW, double areaH)
        {
            if (w < 1e-6 || h < 1e-6 || areaW < 1e-6 || areaH < 1e-6) return 1.0;
            double a = degrees * Math.PI / 180.0, c = Math.Abs(Math.Cos(a)), s = Math.Abs(Math.Sin(a));
            double z = Math.Min(1.0, Math.Min(areaW / (w * c + h * s), areaH / (w * s + h * c)));
            if (z > 1.0 - 1e-9) return 1.0;
            return Math.Max(WallpaperRotation.MinZoom, z);
        }

        // The four corners of the turned picture (made smaller by the zoom), clockwise from the top left one.
        public static PointF[] Corners(WallpaperRotation r, RectangleF bounds)
        {
            WallpaperRotation n = r.Normalize();
            PointF q = Center(bounds);
            float hw = (float)(bounds.Width / 2.0 * n.Zoom), hh = (float)(bounds.Height / 2.0 * n.Zoom);
            double[][] c = { new double[] { -hw, -hh }, new double[] { hw, -hh }, new double[] { hw, hh }, new double[] { -hw, hh } };
            var res = new PointF[4];
            for (int i = 0; i < 4; i++)
            {
                PointF t = Turn(c[i][0], c[i][1], n.Degrees);
                res[i] = new PointF(q.X + t.X, q.Y + t.Y);
            }
            return res;
        }
        // Is the point inside the (convex, clockwise) corners, at least inset pixels away from every side?
        public static bool Inside(PointF[] corners, PointF p, double inset)
        {
            for (int i = 0; i < 4; i++)
            {
                PointF a = corners[i], b = corners[(i + 1) % 4];
                double ex = b.X - a.X, ey = b.Y - a.Y;
                double len = Math.Sqrt(ex * ex + ey * ey);
                if (len < 1e-9) return false;
                double dist = (ex * (p.Y - a.Y) - ey * (p.X - a.X)) / len;
                if (dist < inset - 1e-6) return false;
            }
            return true;
        }

        public static bool FrameInside(PointF[] corners, RectangleF f, double inset)
        {
            return Inside(corners, new PointF(f.Left, f.Top), inset) && Inside(corners, new PointF(f.Right, f.Top), inset)
                && Inside(corners, new PointF(f.Right, f.Bottom), inset) && Inside(corners, new PointF(f.Left, f.Bottom), inset);
        }

        // The angle (degrees, clockwise is positive: right = 0, down = 90) of the point p seen from pivot.
        public static double AngleOf(PointF pivot, PointF p)
        {
            return Math.Atan2(p.Y - pivot.Y, p.X - pivot.X) * 180.0 / Math.PI;
        }

        // The difference from angle a to angle b, in the range -180 (excluded) to 180 (included).
        public static double Delta(double a, double b)
        {
            return WallpaperRotation.WrapDegrees(b - a);
        }

        // Makes the frame smaller (round its middle, with its ratio) until its four corners lie inside the picture. A frame that
        // is inside is not changed, and a frame is never made larger. False when that cannot be done (the middle is outside).
        public static bool FitFrame(RectangleF frame, PointF[] corners, double inset, out RectangleF result)
        {
            result = frame;
            if (FrameInside(corners, frame, inset)) return true;
            float cx = frame.Left + frame.Width / 2f, cy = frame.Top + frame.Height / 2f;
            if (!Inside(corners, new PointF(cx, cy), inset)) return false;
            double lo = 0, hi = 1;
            for (int i = 0; i < 40; i++)
            {
                double m = (lo + hi) / 2;
                if (FrameInside(corners, Scaled(cx, cy, frame, m), inset)) lo = m; else hi = m;
            }
            result = Scaled(cx, cy, frame, lo);
            return true;
        }

        private static RectangleF Scaled(float cx, float cy, RectangleF f, double m)
        {
            float w = (float)(f.Width * m), h = (float)(f.Height * m);
            return new RectangleF(cx - w / 2f, cy - h / 2f, w, h);
        }

        // The frame the pointer asks for (candidate), kept inside the turned picture. It moves from the frame that was valid until now
        // towards the candidate, one part at a time (a move: across, then down; a change of size: left, top, right, bottom), as far as
        // it stays inside: so a frame slides along a side of the picture. A valid frame that is not inside to begin with is not limited.
        public static RectangleF LimitFrame(RectangleF valid, RectangleF candidate, PointF[] corners, double inset, bool move)
        {
            // (a frame that was kept inside and then rounded to four decimals can be a hair outside: a little room is allowed)
            // (the frame the search finds lies on the limit; the frame is then rounded to four decimals, which can put it a little
            // outside the limit: so a frame is taken as valid with a larger room than the limit it is kept to)
            if (!FrameInside(corners, valid, inset - 1.0)) return candidate;
            // (a frame that is a hair outside the limit is kept at least as far in as it is now: so it still slides along the side)
            inset = Math.Min(inset - 0.25, MinDistance(corners, valid));
            var cur = new double[] { valid.Left, valid.Top, valid.Right, valid.Bottom };
            var cand = new double[] { candidate.Left, candidate.Top, candidate.Right, candidate.Bottom };
            if (move)
            {
                double dx = cand[0] - cur[0], dy = cand[1] - cur[1];
                cur = Approach(cur, new[] { cur[0] + dx, cur[1], cur[2] + dx, cur[3] }, corners, inset);
                cur = Approach(cur, new[] { cur[0], cur[1] + dy, cur[2], cur[3] + dy }, corners, inset);
            }
            else
            {
                for (int k = 0; k < 4; k++)
                {
                    double[] target = (double[])cur.Clone();
                    target[k] = cand[k];
                    cur = Approach(cur, target, corners, inset);
                }
            }
            return new RectangleF((float)cur[0], (float)cur[1], (float)(cur[2] - cur[0]), (float)(cur[3] - cur[1]));
        }

        // How far the nearest corner of the frame is from the sides of the picture (negative: outside).
        private static double MinDistance(PointF[] corners, RectangleF f)
        {
            PointF[] pts = { new PointF(f.Left, f.Top), new PointF(f.Right, f.Top), new PointF(f.Right, f.Bottom), new PointF(f.Left, f.Bottom) };
            double min = double.MaxValue;
            for (int i = 0; i < 4; i++)
            {
                PointF a = corners[i], b = corners[(i + 1) % 4];
                double ex = b.X - a.X, ey = b.Y - a.Y;
                double len = Math.Sqrt(ex * ex + ey * ey);
                if (len < 1e-9) return double.MinValue;
                foreach (PointF p in pts) min = Math.Min(min, (ex * (p.Y - a.Y) - ey * (p.X - a.X)) / len);
            }
            return min;
        }

        private static bool Ok(double[] f, PointF[] corners, double inset)
        {
            return FrameInside(corners, new RectangleF((float)f[0], (float)f[1], (float)(f[2] - f[0]), (float)(f[3] - f[1])), inset);
        }

        // From a frame that is inside towards a target, as far as it stays inside (the inside is convex, so one search finds it).
        private static double[] Approach(double[] from, double[] to, PointF[] corners, double inset)
        {
            if (Ok(to, corners, inset)) return to;
            double lo = 0, hi = 1;
            for (int i = 0; i < 40; i++)
            {
                double t = (lo + hi) / 2;
                var f = new double[4];
                for (int k = 0; k < 4; k++) f[k] = from[k] + (to[k] - from[k]) * t;
                if (Ok(f, corners, inset)) lo = t; else hi = t;
            }
            var res = new double[4];
            for (int k = 0; k < 4; k++) res[k] = from[k] + (to[k] - from[k]) * lo;
            return res;
        }
    }
}
