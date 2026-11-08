using System;
using System.Drawing;
using System.Globalization;

namespace VerticalSheetTabs
{
    // How the picture of the wallpaper is turned in the preview of the settings dialog: the angle (clockwise is positive, in degrees)
    // and where the middle of the picture lies (DX, DY: the shift from the middle of the preview, as fractions of its width and
    // height). The picture is turned round its own middle by Degrees and then moved by (DX, DY). All zero (the default value of the
    // struct) means "not turned". The shift is kept because the picture is turned round the middle of the frame, which is not
    // always the middle of the picture.
    public struct WallpaperRotation
    {
        public readonly double Degrees, DX, DY;

        public WallpaperRotation(double degrees, double dx, double dy)
        {
            Degrees = degrees;
            DX = dx;
            DY = dy;
        }

        public static WallpaperRotation None { get { return new WallpaperRotation(0, 0, 0); } }

        // The angle in the range -180 (excluded) to 180 (included).
        public static double WrapDegrees(double d)
        {
            d = d % 360.0;
            if (d > 180.0) d -= 360.0;
            if (d <= -180.0) d += 360.0;
            return d;
        }

        // What is saved: the angle in two decimals, the shift (-1 to 1) in four. Anything that is not a number is "not turned".
        public WallpaperRotation Normalize()
        {
            if (IsBad(Degrees) || IsBad(DX) || IsBad(DY)) return None;
            double a = Math.Round(WrapDegrees(Degrees), 2, MidpointRounding.AwayFromZero);
            if (a <= -180.0) a = 180.0;
            double x = Math.Round(Math.Max(-1, Math.Min(1, DX)), 4, MidpointRounding.AwayFromZero);
            double y = Math.Round(Math.Max(-1, Math.Min(1, DY)), 4, MidpointRounding.AwayFromZero);
            return new WallpaperRotation(a + 0.0, x + 0.0, y + 0.0);
        }

        public bool IsIdentity
        {
            get
            {
                WallpaperRotation n = Normalize();
                return n.Degrees == 0 && n.DX == 0 && n.DY == 0;
            }
        }

        public bool SameAs(WallpaperRotation o)
        {
            WallpaperRotation a = Normalize(), b = o.Normalize();
            return a.Degrees == b.Degrees && a.DX == b.DX && a.DY == b.DY;
        }

        // "Degrees,DX,DY": the point is a period (the same in every culture).
        public string Format()
        {
            WallpaperRotation n = Normalize();
            return n.Degrees.ToString("0.##", CultureInfo.InvariantCulture) + "," + n.DX.ToString("0.####", CultureInfo.InvariantCulture) + ","
                + n.DY.ToString("0.####", CultureInfo.InvariantCulture);
        }

        // Reads "Degrees,DX,DY". Anything that cannot be read (not three numbers) is "not turned".
        public static WallpaperRotation Parse(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return None;
            string[] parts = text.Split(',');
            if (parts.Length != 3) return None;
            var v = new double[3];
            for (int i = 0; i < 3; i++)
            {
                if (!double.TryParse(parts[i].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out v[i])) return None;
            }
            return new WallpaperRotation(v[0], v[1], v[2]).Normalize();
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

        // The middle of the turned picture.
        public static PointF Center(WallpaperRotation r, RectangleF bounds)
        {
            WallpaperRotation n = r;
            return new PointF(bounds.Left + bounds.Width / 2f + (float)(n.DX * bounds.Width), bounds.Top + bounds.Height / 2f + (float)(n.DY * bounds.Height));
        }

        // The four corners of the turned picture, clockwise from the top left one.
        public static PointF[] Corners(WallpaperRotation r, RectangleF bounds)
        {
            PointF q = Center(r, bounds);
            float hw = bounds.Width / 2f, hh = bounds.Height / 2f;
            double[][] c = { new double[] { -hw, -hh }, new double[] { hw, -hh }, new double[] { hw, hh }, new double[] { -hw, hh } };
            var res = new PointF[4];
            for (int i = 0; i < 4; i++)
            {
                PointF t = Turn(c[i][0], c[i][1], r.Degrees);
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

        // The picture turned further by delta degrees (clockwise is positive) round the point pivot. The angle is not rounded.
        public static WallpaperRotation RotateAbout(WallpaperRotation r, RectangleF bounds, PointF pivot, double deltaDegrees)
        {
            PointF q = Center(r, bounds);
            PointF t = Turn(q.X - pivot.X, q.Y - pivot.Y, deltaDegrees);
            double nx = pivot.X + t.X, ny = pivot.Y + t.Y;
            double mx = bounds.Left + bounds.Width / 2.0, my = bounds.Top + bounds.Height / 2.0;
            double w = Math.Max(1e-6, bounds.Width), h = Math.Max(1e-6, bounds.Height);
            return new WallpaperRotation(WallpaperRotation.WrapDegrees(r.Degrees + deltaDegrees), (nx - mx) / w, (ny - my) / h);
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
