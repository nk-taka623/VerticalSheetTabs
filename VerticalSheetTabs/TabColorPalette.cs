using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    public enum TabColorKind
    {
        None,   // "no color"
        Rgb,    // a fixed color
        Theme   // a theme color with a tint, so it follows the workbook's theme like Excel's own
    }

    public class TabColorChoice
    {
        public TabColorKind Kind;
        public int Bgr;          // Rgb: the color as Excel stores it (0xBBGGRR)
        public int ThemeIndex;   // Theme: XlThemeColor (1 dark1, 2 light1, 3 dark2, 4 light2, 5..10 accent1..6)
        public double Tint;      // Theme: -1..1 (TintAndShade)
    }

    // The color palette of the "tab color" menu item: the same layout as Excel's own (theme colors
    // with five tints each, standard colors, "no color", "more colors").
    public class TabColorPalette : Control
    {
        // Excel's column order of the theme colors and the XlThemeColor index of each.
        private static readonly int[] ThemeIndexes = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };
        private static readonly Color[] OfficeTheme =
        {
            Color.FromArgb(255, 255, 255), Color.FromArgb(0, 0, 0), Color.FromArgb(231, 230, 230), Color.FromArgb(68, 84, 106),
            Color.FromArgb(68, 114, 196), Color.FromArgb(237, 125, 49), Color.FromArgb(165, 165, 165), Color.FromArgb(255, 192, 0),
            Color.FromArgb(91, 155, 213), Color.FromArgb(112, 173, 71)
        };
        private static readonly Color[] StandardColors =
        {
            Color.FromArgb(192, 0, 0), Color.FromArgb(255, 0, 0), Color.FromArgb(255, 192, 0), Color.FromArgb(255, 255, 0),
            Color.FromArgb(146, 208, 80), Color.FromArgb(0, 176, 80), Color.FromArgb(0, 176, 240), Color.FromArgb(0, 112, 192),
            Color.FromArgb(0, 32, 96), Color.FromArgb(112, 48, 160)
        };
        // Tints of the five rows below the base color, as Excel's palette shows them (measured in Excel):
        // the white and black columns have their own; the other columns use one of three sets, chosen
        // by the lightness of the theme color (very light: darker shades, very dark: lighter tints).
        private static readonly double[] NormalTints = { 0, 0.8, 0.6, 0.4, -0.25, -0.5 };
        private static readonly double[] WhiteTints = { 0, -0.05, -0.15, -0.25, -0.35, -0.5 };
        private static readonly double[] BlackTints = { 0, 0.5, 0.35, 0.25, 0.15, 0.05 };
        private static readonly double[] VeryLightTints = { 0, -0.1, -0.25, -0.5, -0.75, -0.9 };
        private static readonly double[] VeryDarkTints = { 0, 0.9, 0.75, 0.5, 0.25, 0.1 };

        private static double[] TintsFor(int col, Color themeColor)
        {
            if (col == 0) return WhiteTints;
            if (col == 1) return BlackTints;
            double h, l, sat;
            RgbToHls(themeColor, out h, out l, out sat);
            if (l > 0.8) return VeryLightTints;
            if (l < 0.2) return VeryDarkTints;
            return NormalTints;
        }

        private class Cell
        {
            public Rectangle Bounds;
            public Color Color;
            public TabColorChoice Choice; // null for the text buttons
            public string Text;
            public bool More;
        }

        private readonly List<Cell> _cells = new List<Cell>();
        private readonly List<KeyValuePair<string, Point>> _titles = new List<KeyValuePair<string, Point>>();
        private int _hot = -1;
        private Font _bold;
        private readonly List<int> _dividers = new List<int>();
        private Color[] _theme = OfficeTheme;

        public event Action<TabColorChoice> ColorChosen;

        public TabColorPalette()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
            Font font = SystemFonts.MenuFont;
            // the pane's WinForms scaling is 96 dpi, Excel's menus use the real one
            Font = new Font(font.FontFamily, font.SizeInPoints * 0.92f * Dpi / 96f, FontStyle.Regular, GraphicsUnit.Point);
            _bold = new Font(Font, FontStyle.Bold);
            Rebuild();
        }

        // The ten theme colors of the workbook in Excel's order (white, black, two backgrounds, six accents).
        public void SetTheme(Color[] theme)
        {
            _theme = theme != null && theme.Length == 10 ? theme : OfficeTheme;
            Rebuild();
            Invalidate();
        }

        private int Dpi
        {
            get { return DpiHelper.GetDpi(IsHandleCreated ? Handle : IntPtr.Zero); }
        }

        // DIP to pixels at the real DPI.
        private int D(double dip)
        {
            return (int)Math.Round(dip * Dpi / 96.0);
        }

        // The layout is Excel's own, measured at 125% and given here in DIP: swatches 24 x 24 with 3.2
        // between columns, the five shades touch each other, a header band of 34 above every group.
        private void Rebuild()
        {
            _cells.Clear();
            _titles.Clear();
            _dividers.Clear();
            int x0 = D(2.4), pitch = D(27.2), size = D(24), headerH = D(34), itemH = D(26);
            int width = 9 * pitch + size;
            int y = 0;

            _titles.Add(new KeyValuePair<string, Point>("テーマの色", new Point(x0 + D(2), y + headerH / 2)));
            y += headerH;
            for (int row = 0; row < 6; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    double tint = TintsFor(col, _theme[col])[row];
                    var choice = new TabColorChoice { Kind = TabColorKind.Theme, ThemeIndex = ThemeIndexes[col], Tint = tint };
                    _cells.Add(new Cell
                    {
                        Bounds = new Rectangle(x0 + col * pitch, y, size, size),
                        Color = ApplyTint(_theme[col], tint),
                        Choice = choice
                    });
                }
                y += size;
                if (row == 0) y += D(12.8);
            }

            y += D(12.8);
            _dividers.Add(y);
            _titles.Add(new KeyValuePair<string, Point>("標準の色", new Point(x0 + D(2), y + headerH / 2)));
            y += headerH;
            for (int col = 0; col < StandardColors.Length; col++)
            {
                Color c = StandardColors[col];
                _cells.Add(new Cell
                {
                    Bounds = new Rectangle(x0 + col * pitch, y, size, size),
                    Color = c,
                    Choice = new TabColorChoice { Kind = TabColorKind.Rgb, Bgr = c.R | (c.G << 8) | (c.B << 16) }
                });
            }
            y += size + D(11.2);

            _dividers.Add(y);
            y += D(3);
            _cells.Add(new Cell
            {
                Bounds = new Rectangle(0, y, width + 2 * x0, itemH),
                Text = "色なし(&N)",
                Choice = new TabColorChoice { Kind = TabColorKind.None }
            });
            y += itemH;
            _cells.Add(new Cell { Bounds = new Rectangle(0, y, width + 2 * x0, itemH), Text = "その他の色(&M)...", More = true });
            y += itemH + D(3);

            Size = new Size(width + 2 * x0, y);
        }

        // Excel's tint: lighten (tint > 0) or darken (tint < 0) the luminance in HLS space.
        public static Color ApplyTint(Color color, double tint)
        {
            if (tint == 0) return color;
            double h, l, s;
            RgbToHls(color, out h, out l, out s);
            l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
            return HlsToRgb(h, Math.Max(0, Math.Min(1, l)), s);
        }

        private static void RgbToHls(Color c, out double h, out double l, out double s)
        {
            double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
            double max = Math.Max(r, Math.Max(g, b)), min = Math.Min(r, Math.Min(g, b));
            l = (max + min) / 2;
            double d = max - min;
            if (d == 0) { h = 0; s = 0; return; }
            s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
            if (max == r) h = ((g - b) / d + (g < b ? 6 : 0)) / 6;
            else if (max == g) h = ((b - r) / d + 2) / 6;
            else h = ((r - g) / d + 4) / 6;
        }

        private static Color HlsToRgb(double h, double l, double s)
        {
            if (s == 0) { int v = (int)Math.Round(l * 255); return Color.FromArgb(v, v, v); }
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            return Color.FromArgb(
                (int)Math.Round(255 * Hue(p, q, h + 1.0 / 3)),
                (int)Math.Round(255 * Hue(p, q, h)),
                (int)Math.Round(255 * Hue(p, q, h - 1.0 / 3)));
        }

        private static double Hue(double p, double q, double t)
        {
            if (t < 0) t += 1;
            if (t > 1) t -= 1;
            if (t < 1.0 / 6) return p + (q - p) * 6 * t;
            if (t < 0.5) return q;
            if (t < 2.0 / 3) return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            foreach (var t in _titles)
            {
                TextRenderer.DrawText(e.Graphics, t.Key, _bold, new Point(t.Value.X, t.Value.Y - _bold.Height / 2), Color.Black, TextFormatFlags.NoPadding);
            }
            using (var line = new Pen(Color.FromArgb(225, 225, 225), 1))
            {
                foreach (int y in _dividers) e.Graphics.DrawLine(line, 0, y, Width, y);
            }
            for (int i = 0; i < _cells.Count; i++)
            {
                Cell c = _cells[i];
                bool hot = i == _hot;
                if (c.Text != null)
                {
                    if (hot) using (var b = new SolidBrush(Color.FromArgb(240, 240, 240))) e.Graphics.FillRectangle(b, c.Bounds);
                    TextRenderer.DrawText(e.Graphics, c.Text.Replace("&", ""), Font,
                        new Rectangle(c.Bounds.X + D(16), c.Bounds.Y, c.Bounds.Width - D(16), c.Bounds.Height), Color.Black,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                    continue;
                }
                using (var b = new SolidBrush(c.Color)) e.Graphics.FillRectangle(b, c.Bounds);
                Rectangle r = c.Bounds;
                if (hot)
                {
                    using (var pen = new Pen(Color.FromArgb(60, 60, 60), 2))
                    {
                        r.Inflate(-1, -1);
                        e.Graphics.DrawRectangle(pen, r);
                    }
                }
                else if (c.Color.GetBrightness() > 0.97f)
                {
                    // a white swatch needs an outline, the others have none (like Excel's)
                    using (var pen = new Pen(Color.FromArgb(190, 190, 190), 1))
                    {
                        r.Width -= 1; r.Height -= 1;
                        e.Graphics.DrawRectangle(pen, r);
                    }
                }
            }
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int hot = _cells.FindIndex(c => c.Bounds.Contains(e.Location));
            if (hot != _hot) { _hot = hot; Invalidate(); }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_hot != -1) { _hot = -1; Invalidate(); }
        }

        // Asks for a color with Excel's dialog: 0xBBGGRR, or null when it was cancelled.
        public Func<int?> MoreColorsProvider { get; set; }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            Cell cell = _cells.Find(c => c.Bounds.Contains(e.Location));
            if (cell == null) return;
            Choose(cell);
        }

        private void Choose(Cell cell)
        {
            TabColorChoice choice = cell.Choice;
            if (cell.More)
            {
                // Excel's own "Colors" dialog (the one of the horizontal tab's "More Colors") when it can be
                // opened; the system color dialog otherwise.
                if (MoreColorsProvider != null)
                {
                    int? picked = MoreColorsProvider();
                    if (picked == null) return;
                    choice = new TabColorChoice { Kind = TabColorKind.Rgb, Bgr = picked.Value };
                    if (ColorChosen != null) ColorChosen(choice);
                    return;
                }
                using (var dialog = new ColorDialog { FullOpen = true, AnyColor = true })
                {
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    Color c = dialog.Color;
                    choice = new TabColorChoice { Kind = TabColorKind.Rgb, Bgr = c.R | (c.G << 8) | (c.B << 16) };
                }
            }
            if (choice != null && ColorChosen != null) ColorChosen(choice);
        }
    }
}
