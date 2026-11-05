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
            foreach (CommentEntry c in items)
            {
                CommentEntry item = c;
                var mi = new ToolStripMenuItem(c.ListText);
                mi.Click += (s, e) => picked = item;
                menu.Items.Add(mi);
            }
            // the content is shown when the menu has closed (a form opened while the menu is closing would be closed with it)
            menu.Closed += (s, e) =>
            {
                CommentEntry c = picked;
                var timer = new Timer { Interval = 30 };
                timer.Tick += (s2, e2) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    menu.Dispose();
                    if (c != null) ShowContent(screen, owner, c, goTo);
                };
                timer.Start();
            };
            menu.Show(screen);
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
