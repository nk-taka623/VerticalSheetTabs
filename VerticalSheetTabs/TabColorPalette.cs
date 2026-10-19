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
        // Tints of the rows below the base color; white and black columns use their own, as in Excel.
        private static readonly double[] NormalTints = { 0, 0.8, 0.6, 0.4, -0.25, -0.5 };
        private static readonly double[] LightTints = { 0, -0.05, -0.15, -0.25, -0.35, -0.5 };
        private static readonly double[] DarkTints = { 0, 0.5, 0.35, 0.25, 0.15, 0.05 };

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
        private Color[] _theme = OfficeTheme;

        public event Action<TabColorChoice> ColorChosen;

        public TabColorPalette()
        {
            DoubleBuffered = true;
            BackColor = SystemColors.Window;
            Font = SystemFonts.MenuFont;
            Rebuild();
        }

        // The ten theme colors of the workbook in Excel's order (white, black, two backgrounds, six accents).
        public void SetTheme(Color[] theme)
        {
            _theme = theme != null && theme.Length == 10 ? theme : OfficeTheme;
            Rebuild();
            Invalidate();
        }

        private int S(int v)
        {
            return (int)Math.Round(v * DeviceDpi / 96.0);
        }

        private void Rebuild()
        {
            _cells.Clear();
            _titles.Clear();
            int pad = S(6), size = S(18), gap = S(3), x0 = pad, y = pad;
            int titleH = Font.Height + S(4);

            _titles.Add(new KeyValuePair<string, Point>("テーマの色", new Point(x0, y)));
            y += titleH;
            for (int row = 0; row < 6; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    double[] tints = col == 0 ? LightTints : col == 1 ? DarkTints : NormalTints;
                    double tint = tints[row];
                    var choice = new TabColorChoice { Kind = TabColorKind.Theme, ThemeIndex = ThemeIndexes[col], Tint = tint };
                    _cells.Add(new Cell
                    {
                        Bounds = new Rectangle(x0 + col * (size + gap), y, size, size),
                        Color = ApplyTint(_theme[col], tint),
                        Choice = choice
                    });
                }
                y += size + (row == 0 ? gap + S(3) : gap);
            }

            y += S(4);
            _titles.Add(new KeyValuePair<string, Point>("標準の色", new Point(x0, y)));
            y += titleH;
            for (int col = 0; col < StandardColors.Length; col++)
            {
                Color c = StandardColors[col];
                _cells.Add(new Cell
                {
                    Bounds = new Rectangle(x0 + col * (size + gap), y, size, size),
                    Color = c,
                    Choice = new TabColorChoice { Kind = TabColorKind.Rgb, Bgr = c.R | (c.G << 8) | (c.B << 16) }
                });
            }
            y += size + S(8);

            int width = 10 * (size + gap) - gap;
            _cells.Add(new Cell
            {
                Bounds = new Rectangle(x0, y, width, titleH + S(2)),
                Text = "色なし(&N)",
                Choice = new TabColorChoice { Kind = TabColorKind.None }
            });
            y += titleH + S(2);
            _cells.Add(new Cell { Bounds = new Rectangle(x0, y, width, titleH + S(2)), Text = "その他の色(&M)...", More = true });
            y += titleH + S(2);

            Size = new Size(x0 + width + pad, y + pad);
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
                TextRenderer.DrawText(e.Graphics, t.Key, Font, t.Value, SystemColors.ControlText, TextFormatFlags.NoPadding);
            }
            for (int i = 0; i < _cells.Count; i++)
            {
                Cell c = _cells[i];
                bool hot = i == _hot;
                if (c.Text != null)
                {
                    if (hot) using (var b = new SolidBrush(SystemColors.Highlight)) e.Graphics.FillRectangle(b, c.Bounds);
                    TextRenderer.DrawText(e.Graphics, c.Text.Replace("&", ""), Font,
                        Rectangle.Inflate(c.Bounds, -S(4), 0), hot ? SystemColors.HighlightText : SystemColors.ControlText,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix);
                    continue;
                }
                using (var b = new SolidBrush(c.Color)) e.Graphics.FillRectangle(b, c.Bounds);
                using (var pen = new Pen(hot ? Color.FromArgb(255, 170, 0) : Color.FromArgb(170, 170, 170), hot ? 2 : 1))
                {
                    Rectangle r = c.Bounds;
                    r.Width -= 1; r.Height -= 1;
                    e.Graphics.DrawRectangle(pen, r);
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
                // The system color dialog stands in for Excel's "more colors" dialog.
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
