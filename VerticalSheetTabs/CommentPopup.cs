using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // One note or comment of a sheet, as it is shown (read-only) by CommentPopup.
    public class CommentEntry
    {
        public string Sheet;
        public string Address;          // the cell, e.g. "B3" (no sheet name)
        public string Author;
        public string Text;
        public bool Threaded;           // a threaded comment (with replies), not a legacy note
        public bool Done;               // the thread is resolved
        public List<string> Replies = new List<string>();

        public string ListText
        {
            get
            {
                string t = (Text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
                if (t.Length > 28) t = t.Substring(0, 28) + "…";
                return Address + "  " + (string.IsNullOrEmpty(Author) ? "" : Author + ": ") + t;
            }
        }
    }

    // The pop-ups of the note / comment badge: a list of the comments at the click position (when there are several), and the content
    // of one comment, read-only, with a button that goes to its cell. Nothing here touches Excel: the caller gives the comments and
    // the action of the button.
    public static class CommentPopup
    {
        public static void ShowList(Point screen, IWin32Window owner, List<CommentEntry> items, Action<CommentEntry> goTo)
        {
            var menu = new ContextMenuStrip();
            CommentEntry picked = null;
            bool pickedGoTo = false;
            float scale = DpiHelper.GetDpi(IntPtr.Zero) / 96f;
            foreach (CommentEntry c in items)
            {
                CommentEntry item = c;
                var mi = new GoToMenuItem(c.ListText, scale);
                mi.Click += (s, e) => { picked = item; pickedGoTo = mi.IconClicked; };
                menu.Items.Add(mi);
            }
            // the content is shown when the menu has closed (a form opened while the menu is closing would be closed with it)
            menu.Closed += (s, e) =>
            {
                var timer = new Timer { Interval = 60 };
                timer.Tick += (s2, e2) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    CommentEntry c = picked; // (read now: the item's Click can come after Closed)
                    bool go = pickedGoTo;
                    menu.Dispose();
                    if (c == null) return;
                    if (go) goTo(c); // the icon at the right end: straight to the cell
                    else ShowContent(screen, owner, c, goTo);
                };
                timer.Start();
            };
            menu.Show(screen);
        }

        // A list row with the "go to the cell" icon (a table with an arrow, 16 DIP) at its right end. The icon is drawn here (GDI+).
        private class GoToMenuItem : ToolStripMenuItem
        {
            private readonly float _scale;
            private bool _hot;
            public bool IconClicked;

            public GoToMenuItem(string text, float scale) : base(text) { _scale = scale; }

            private int Px(float v) { return (int)Math.Round(v * _scale); }

            private Rectangle IconBox()
            {
                int box = Px(22);
                int right = Width;
                if (Owner != null) right = Math.Min(right, Owner.ClientSize.Width - Bounds.X); // (the menu can be narrower than the item)
                return new Rectangle(right - box - Px(4), (Height - box) / 2, box, box);
            }

            public override Size GetPreferredSize(Size constrainingSize)
            {
                Size s = base.GetPreferredSize(constrainingSize);
                return new Size(s.Width + Px(30), Math.Max(s.Height, Px(26)));
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                bool hot = IconBox().Contains(e.Location);
                if (hot != _hot) { _hot = hot; Invalidate(); }
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                if (_hot) { _hot = false; Invalidate(); }
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                IconClicked = IconBox().Contains(e.Location);
                base.OnMouseDown(e);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                try
                {
                    Rectangle box = IconBox();
                    var g = e.Graphics;
                    g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    Color ink = Color.FromArgb(96, 96, 96);
                    if (_hot)
                    {
                        ink = Color.FromArgb(16, 124, 65);
                        using (var fill = new SolidBrush(Color.White))
                        using (var pen = new Pen(ink, Math.Max(1f, _scale * 0.8f)))
                        {
                            Rectangle r = new Rectangle(box.X, box.Y, box.Width - 1, box.Height - 1);
                            g.FillRectangle(fill, r);
                            g.DrawRectangle(pen, r);
                        }
                    }
                    // The icon is drawn on a 24 x 24 grid (a table with an arrow at the lower left), shrunk to 16 DIP and centered in the box.
                    float k = 16f / 24f * _scale;
                    float ox = box.X + (box.Width - 24 * k) / 2f, oy = box.Y + (box.Height - 24 * k) / 2f;
                    Func<float, float> X = v => ox + v * k;
                    Func<float, float> Y = v => oy + v * k;
                    using (var pen = new Pen(ink, Math.Max(1f, 2f * k)))
                    {
                        pen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                        pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                        pen.LineJoin = System.Drawing.Drawing2D.LineJoin.Round;
                        // the frame: open at the lower left (from the left side up, round the corners, down the right side, back along the bottom)
                        using (var frame = new System.Drawing.Drawing2D.GraphicsPath())
                        {
                            frame.AddLine(X(3), Y(13), X(3), Y(5));
                            frame.AddArc(X(3), Y(3), 4 * k, 4 * k, 180, 90);
                            frame.AddLine(X(5), Y(3), X(19), Y(3));
                            frame.AddArc(X(17), Y(3), 4 * k, 4 * k, 270, 90);
                            frame.AddLine(X(21), Y(5), X(21), Y(19));
                            frame.AddArc(X(17), Y(17), 4 * k, 4 * k, 0, 90);
                            frame.AddLine(X(19), Y(21), X(11), Y(21));
                            g.DrawPath(pen, frame);
                        }
                        g.DrawLine(pen, X(3), Y(10), X(21), Y(10));   // the header line
                        g.DrawLine(pen, X(10), Y(3), X(10), Y(14));   // the vertical line
                        g.DrawLine(pen, X(2), Y(22), X(7), Y(17));    // the arrow: shaft
                        g.DrawLine(pen, X(7), Y(21.5f), X(7), Y(17)); // the arrow: head
                        g.DrawLine(pen, X(7), Y(17), X(2.5f), Y(17));
                    }
                }
                catch (Exception)
                {
                    // Drawing the icon must not break the menu.
                }
            }
        }

        public static void ShowContent(Point screen, IWin32Window owner, CommentEntry c, Action<CommentEntry> goTo)
        {
            try
            {
                float scale = DpiHelper.GetDpi(IntPtr.Zero) / 96f;
                Func<int, int> px = v => (int)Math.Round(v * scale);
                var form = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    StartPosition = FormStartPosition.Manual,
                    ShowInTaskbar = false,
                    KeyPreview = true,
                    BackColor = Color.FromArgb(255, 255, 225),
                    Size = new Size(px(320), px(240)),
                };
                form.Font = new Font("Yu Gothic UI", 12f * scale, FontStyle.Regular, GraphicsUnit.Pixel);

                string head = c.Sheet + " ! " + c.Address + (string.IsNullOrEmpty(c.Author) ? "" : "  (" + c.Author + ")")
                    + (c.Threaded && c.Done ? "  [解決済み]" : "");
                var header = new Label
                {
                    Text = head,
                    AutoEllipsis = true,
                    Dock = DockStyle.Top,
                    Height = px(26),
                    Padding = new Padding(px(8), px(4), px(8), 0),
                    ForeColor = Color.FromArgb(70, 70, 70),
                };
                string body = c.Text ?? "";
                for (int i = 0; i < c.Replies.Count; i++)
                {
                    if (i == 0) body += "\r\n\r\n── 返信 (" + c.Replies.Count + ") ──";
                    body += "\r\n" + c.Replies[i];
                }
                var box = new TextBox
                {
                    Multiline = true,
                    ReadOnly = true,
                    ScrollBars = ScrollBars.Vertical,
                    BorderStyle = BorderStyle.None,
                    BackColor = form.BackColor,
                    Dock = DockStyle.Fill,
                    Text = body.Replace("\r\n", "\n").Replace("\n", "\r\n"),
                    TabStop = false,
                };
                var bottom = new Panel { Dock = DockStyle.Bottom, Height = px(34), BackColor = form.BackColor };
                var go = new Button
                {
                    Text = "そのセルへ移動",
                    FlatStyle = FlatStyle.System,
                    Size = new Size(px(120), px(26)),
                    Location = new Point(px(8), px(4)),
                };
                go.Click += (s, e) =>
                {
                    form.Close();
                    if (goTo != null) goTo(c);
                };
                bottom.Controls.Add(go);
                var pad = new Panel { Dock = DockStyle.Fill, Padding = new Padding(px(8), 0, px(8), 0), BackColor = form.BackColor };
                pad.Controls.Add(box);
                form.Controls.Add(pad);
                form.Controls.Add(bottom);
                form.Controls.Add(header);
                form.Paint += (s, e) => e.Graphics.DrawRectangle(Pens.Gray, 0, 0, form.Width - 1, form.Height - 1);
                form.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) form.Close(); };
                form.Deactivate += (s, e) => { if (!form.IsDisposed) form.Close(); };
                form.FormClosed += (s, e) => form.Dispose();

                Rectangle area = Screen.FromPoint(screen).WorkingArea;
                int x = Math.Max(area.Left, Math.Min(screen.X, area.Right - form.Width));
                int y = Math.Max(area.Top, Math.Min(screen.Y, area.Bottom - form.Height));
                form.Location = new Point(x, y);
                form.Show(owner);
                form.Activate();
            }
            catch (Exception)
            {
                // The window could not be shown: nothing else to do.
            }
        }
    }
}
