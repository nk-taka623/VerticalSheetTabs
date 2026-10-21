using System.Drawing;
using System.Drawing.Drawing2D;

namespace VerticalSheetTabs
{
    // The icon of the ribbon button: a sheet with a column of vertical tabs at its left edge (a tab, a
    // folder, two indented tabs in it, two more tabs). It is drawn in a 32 x 32 design grid and scaled to
    // the requested size, so that it stays sharp at any DPI.
    public static class RibbonIcon
    {
        private const float Grid = 32f;

        public static Bitmap Create(int size)
        {
            var bmp = new Bitmap(size, size, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.Clear(Color.Transparent);
                float s = size / Grid;
                g.ScaleTransform(s, s);

                using (var path = RoundedRect(0.5f, 0.5f, 31f, 31f, 3f))
                using (var fill = new SolidBrush(Color.FromArgb(243, 242, 243)))
                using (var pen = new Pen(Color.FromArgb(180, 178, 169), 0.6f))
                {
                    g.FillPath(fill, path);
                    g.DrawPath(pen, path);
                }

                // The sheet and its lines.
                using (var path = RoundedRect(11f, 3f, 18.5f, 26f, 1.2f))
                using (var fill = new SolidBrush(Color.White))
                using (var pen = new Pen(Color.FromArgb(136, 135, 128), 0.8f))
                {
                    g.FillPath(fill, path);
                    g.DrawPath(pen, path);
                }
                using (var pen = new Pen(Color.FromArgb(209, 209, 209), 0.6f))
                {
                    for (int y = 8; y <= 24; y += 4) g.DrawLine(pen, 14f, y, 27f, y);
                }

                // The tab column: tab, folder, two indented tabs, two tabs.
                Tab(g, 3f, 3.6f, 9f, Color.FromArgb(74, 143, 212));
                Folder(g);
                Tab(g, 5.6f, 12.4f, 6.4f, Color.FromArgb(232, 145, 60));
                Tab(g, 5.6f, 16.8f, 6.4f, Color.FromArgb(16, 124, 65));
                Tab(g, 3f, 21.2f, 9f, Color.FromArgb(74, 143, 212));
                Tab(g, 3f, 25.6f, 9f, Color.FromArgb(154, 152, 144));
            }
            return bmp;
        }

        private static void Tab(Graphics g, float x, float y, float w, Color color)
        {
            using (var path = RoundedRect(x, y, w, 3.4f, 0.9f))
            using (var brush = new SolidBrush(color))
            {
                g.FillPath(brush, path);
            }
        }

        private static void Folder(Graphics g)
        {
            var points = new[]
            {
                new PointF(3f, 8f), new PointF(7.6f, 8f), new PointF(8.5f, 8.8f),
                new PointF(12f, 8.8f), new PointF(12f, 11.4f), new PointF(3f, 11.4f)
            };
            using (var brush = new SolidBrush(Color.FromArgb(232, 184, 74)))
            using (var pen = new Pen(Color.FromArgb(184, 138, 30), 0.5f))
            {
                pen.LineJoin = LineJoin.Round;
                g.FillPolygon(brush, points);
                g.DrawPolygon(pen, points);
            }
        }

        private static GraphicsPath RoundedRect(float x, float y, float w, float h, float r)
        {
            float d = r * 2;
            var path = new GraphicsPath();
            path.AddArc(x, y, d, d, 180, 90);
            path.AddArc(x + w - d, y, d, d, 270, 90);
            path.AddArc(x + w - d, y + h - d, d, d, 0, 90);
            path.AddArc(x, y + h - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }
    }
}
