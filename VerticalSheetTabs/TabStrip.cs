using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // The vertical stack of sheet tabs, drawn like Excel's own horizontal tabs (measurements in
    // Agents.md 9.2). Each tab is one row of the height of Excel's tab bar. The control owns the
    // selection model (click, Ctrl+click, Shift+click), drag and drop, in-place renaming and
    // scrolling; the sheet commands themselves are reported to the pane through events.
    // One row of the list: a sheet, or a folder (FolderId != null).
    public class StripRow
    {
        public string Sheet;
        public bool Indented;
        public string FolderId;
        public string FolderName;
        public bool Expanded;
        public int ColorBgr = -1;
        // Inside the dashed outline of an open folder (the folder row and its sheet rows): these rows end
        // before the outline by the width of its cut corner.
        public bool InGroup;
        // Folder: its visible sheets (also while it is closed).
        public List<string> Members = new List<string>();

        public bool IsFolder { get { return FolderId != null; } }

        public string Signature()
        {
            return IsFolder ? "F|" + FolderId + "|" + FolderName + "|" + Expanded + "|" + ColorBgr + "|" + string.Join("|", Members) : "S|" + Sheet + "|" + Indented;
        }
    }

    public class TabStrip : Control
    {
        // ---- Excel's colors (measured, Officeテーマ: カラフル) ----
        // The same in every Office theme: the underline, the text on a white tab and on a colored tab.
        private static readonly Color Underline = Color.FromArgb(16, 124, 65);
        private static readonly Color TextColor = Color.FromArgb(36, 36, 36);

        private TabTheme _theme = TabTheme.Current;
        private TabTheme Theme { get { return _theme; } }

        // ---- Excel's metrics in DIP (physical px at 125% / 1.25) ----
        private const double RowHeight = 32;
        private const double InsetLeft = 3, InsetRight = 4;
        private const double PlainTop = 1, ColoredTop = 4, BottomInset = 4;
        private const double CornerRadius = 4;
        private const double UnderlineInset = 8, UnderlineThickness = 1.6, UnderlineGap = 3.2;
        private const double ActiveTextLift = 3.2;
        private const double SeparatorLength = 19;
        private const double TextPadLeft = 12, TextPadRight = 8;
        // Folders: sheets inside a folder are indented; the folder row has an open/closed mark and,
        // with a color, a bar at its left edge.
        private const double OutlineRightInset = 1.5, OutlineCut = 10;
        private const double FolderIndent = 14, FolderMarkWidth = 6, FolderMarkHeight = 10, FolderBarWidth = 4;
        // Excel's tab text is about 9pt "Yu Gothic" (the measured width of "Sheet1" matches it).
        private const string FontFamilyName = "Yu Gothic";
        private const double FontPixelsAt96 = 15.5;
        // The "+" (new sheet) button: a plain plus sign that gets a light rounded square on hover.
        private const double PlusArm = 10.5, PlusThickness = 1.6, PlusHoverSize = 24, PlusHoverRadius = 3.2;
        // The drop marker while dragging: Excel shows a small black triangle at the insertion point.
        private const double MarkerWidth = 4, MarkerHeight = 8;

        private List<string> _names = new List<string>();
        // What is on screen, row by row (equal to the sheets when the workbook has no folders).
        private List<StripRow> _rows = new List<StripRow>();
        private string _rowSignature = "";
        private HashSet<string> _selected = new HashSet<string>();
        private string _active;
        private readonly Dictionary<string, int> _colors = new Dictionary<string, int>();

        private readonly VScrollBar _scroll = new VScrollBar { Dock = DockStyle.Right, Visible = false };
        private int _scrollY;
        private Font _font;
        private int _fontDpi;

        private TextBox _editor;
        private string _editingName;

        // drag and drop
        private Point _dragOrigin;
        private int _dragIndex = -1;
        private bool _dragging;
        private List<string> _dragNames = new List<string>();
        private int _dropIndex = -1;
        // Activating a sheet moves the focus into Excel, which would cancel a drag that has just
        // started. So a left click only reaches Excel when the button is released.
        private bool _selectionPending;
        // Selected folders (a selected folder has all its sheets selected as well) and where Shift+click starts.
        private HashSet<string> _selFolders = new HashSet<string>();
        private string _anchorFolder;
        private List<string> _dragSheets = new List<string>();
        private List<string> _dragFolders = new List<string>();
        // Where the dragged things would land: inside a folder, or on the top level.
        private bool _dropInside;
        private string _dropInto, _dropBeforeSheet, _dropBeforeFolder;
        private string _editingFolderId;
        private string _rightFolder;
        private string _toggleOffFolder;
        private bool _plusHot, _plusDown;
        // The full name of a tab whose text is cut off is shown as a tool tip (Excel's own tabs
        // grow to fit the name, so they never need one).
        private readonly ToolTip _tip = new ToolTip { ShowAlways = true, InitialDelay = 500, ReshowDelay = 100 };
        private int _hoverIndex = -1;

        // (selected sheets in tab order, the active one) after the user changed the selection.
        public event Action<IList<string>, string> SelectionRequested;
        // (old name, new name) after the user finishes in-place editing.
        public event Action<string, string> RenameRequested;
        // (sheets, the sheet they are dropped in front of or null for the end, copy instead of move)
        public event Action<IList<string>, string, bool> MoveRequested;
        // The "+" button was clicked.
        public event Action AddRequested;
        // The mark of a folder row was clicked (folder id).
        public event Action<string> FolderToggleRequested;
        // (sheets, folders, the folder they are dropped into or null, the loose sheet or the folder
        // they are dropped in front of; all null = at the end, copy instead of move)
        public event Action<IList<string>, IList<string>, string, string, string, bool> FolderMoveRequested;
        // (folder id, new name) after the user finishes editing a folder name.
        public event Action<string, string> FolderRenameRequested;
        // Ctrl+S / F12 pressed while the pane has the focus: Excel does not get these keys then.
        public event Action<string> ExcelCommandRequested;
        // The Delete key: (selected sheets, selected folders).
        public event Action<IList<string>, IList<string>> DeleteKeyRequested;

        public ContextMenuStrip SheetMenu { get; set; }
        public ContextMenuStrip FolderMenu { get; set; }
        // The folder a right click opened the folder menu on.
        public string RightClickedFolder { get { return _rightFolder; } }
        public List<string> SelectedFolders() { return _selFolders.ToList(); }
        // With the workbook structure protected the tabs can be neither dragged nor renamed.
        public bool StructureProtected { get; set; }

        // With the workbook structure protected the "+" button is dimmed as well.
        public int SelectedCount { get { return _selected.Count; } }
        public string ActiveName { get { return _active; } }

        public TabStrip()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer
                | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            AllowDrop = true;
            BackColor = _theme.BarBack;
            _scroll.Dock = DockStyle.None;
            _scroll.Scroll += (s, e) => { _scrollY = _scroll.Value; Invalidate(); };
            Controls.Add(_scroll);
        }

        // ---- DPI: the pane's WinForms scaling is always 96 dpi, Excel itself runs at the real one ----

        private int Dpi
        {
            get { return DpiHelper.GetDpi(IsHandleCreated ? Handle : IntPtr.Zero); }
        }

        private int Px(double dip)
        {
            return (int)Math.Round(dip * Dpi / 96.0);
        }

        private int RowPx { get { return Px(RowHeight); } }

        private Font TabFont
        {
            get
            {
                int dpi = Dpi;
                if (_font == null || _fontDpi != dpi)
                {
                    if (_font != null) _font.Dispose();
                    FontFamily family;
                    try
                    {
                        // The family list uses localized names ("游ゴシック"), the constructor accepts the English one.
                        family = new FontFamily(FontFamilyName);
                    }
                    catch (ArgumentException)
                    {
                        family = SystemFonts.MessageBoxFont.FontFamily;
                    }
                    _font = new Font(family, (float)(FontPixelsAt96 * dpi / 96.0), FontStyle.Regular, GraphicsUnit.Pixel);
                    _fontDpi = dpi;
                }
                return _font;
            }
        }

        // First-open width: 12 half-width characters plus the tab's own paddings.
        public int DefaultContentWidth()
        {
            Size text = TextRenderer.MeasureText(new string('0', 12), TabFont, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding);
            return text.Width + Px(InsetLeft + InsetRight + TextPadLeft + TextPadRight);
        }

        // ---- data from the pane ----

        // Updates what changed only, so the polling timer does not cause flicker.
        public void SetSheets(IList<string> names, string activeName, ICollection<string> selected, IDictionary<string, int> colors, IList<StripRow> rows)
        {
            RefreshTheme();
            List<StripRow> newRows = rows != null
                ? new List<StripRow>(rows)
                : names.Select(n => new StripRow { Sheet = n }).ToList();
            MarkGroups(newRows);
            string newSignature = string.Join("\n", newRows.Select(r => r.Signature()));
            bool sameRows = newSignature == _rowSignature;
            bool sameColors = _colors.Count == colors.Count && _colors.All(kv => { int v; return colors.TryGetValue(kv.Key, out v) && v == kv.Value; });
            bool sameNames = _names.SequenceEqual(names);
            // While a mouse button is down the user is changing the selection; do not fight them.
            bool userBusy = Control.MouseButtons != MouseButtons.None;
            bool sameSelection = userBusy || (_active == activeName && _selected.SetEquals(selected));
            if (sameNames && sameRows && sameSelection && sameColors) return;

            if (!sameColors)
            {
                _colors.Clear();
                foreach (var kv in colors) _colors[kv.Key] = kv.Value;
            }
            if (!sameNames || !sameRows)
            {
                EndRename(false);
                _names = new List<string>(names);
                _rows = newRows;
                _rowSignature = newSignature;
                UpdateScrollBar();
            }
            if (!userBusy)
            {
                bool activeChanged = _active != activeName;
                _active = activeName;
                _selected = new HashSet<string>(selected);
                if (activeChanged) EnsureVisible(RowIndexOf(_active));
                PruneFolders();
            }
            Invalidate();
        }

        // The Office theme can be changed while Excel runs.
        private void RefreshTheme()
        {
            TabTheme current = TabTheme.Current;
            if (ReferenceEquals(current, _theme)) return;
            _theme = current;
            BackColor = _theme.BarBack;
            Invalidate();
        }

        public List<string> SelectedInOrder()
        {
            return _names.Where(n => _selected.Contains(n)).ToList();
        }

        // ---- geometry ----

        // The "+" sits after the last tab like in Excel; when the tabs overflow it stays at the bottom.
        private bool Pinned { get { return (_rows.Count + 1) * RowPx > ClientSize.Height; } }

        private int ViewHeight { get { return Pinned ? Math.Max(0, ClientSize.Height - RowPx) : ClientSize.Height; } }

        // The rows of an open folder that has sheets listed (they are enclosed by the dashed outline).
        private static void MarkGroups(List<StripRow> rows)
        {
            for (int i = 0; i < rows.Count; i++)
            {
                rows[i].InGroup = false;
                if (!rows[i].IsFolder || !rows[i].Expanded) continue;
                int last = i;
                while (last + 1 < rows.Count && rows[last + 1].Indented) last++;
                if (last == i) continue;
                for (int k = i; k <= last; k++) rows[k].InGroup = true;
                i = last;
            }
        }

        // The space right of a row's tab: the usual inset; in an outlined folder the tab ends as far from
        // the dashed line as the cut corner of the outline is wide.
        private int RightInset(StripRow r)
        {
            return r.InGroup ? Px(OutlineRightInset + OutlineCut) : Px(InsetRight);
        }

        private int RowIndexOf(string sheet)
        {
            return sheet == null ? -1 : _rows.FindIndex(r => !r.IsFolder && r.Sheet == sheet);
        }

        private bool HasFolders { get { return _rows.Any(r => r.IsFolder); } }

        // A folder stays selected only while all its sheets are.
        private void PruneFolders()
        {
            var gone = _selFolders.Where(id =>
            {
                StripRow r = _rows.FirstOrDefault(x => x.IsFolder && x.FolderId == id);
                return r == null;
            }).ToList();
            foreach (string id in gone) _selFolders.Remove(id);
            if (_anchorFolder != null && !_rows.Any(x => x.IsFolder && x.FolderId == _anchorFolder)) _anchorFolder = null;
        }

        private StripRow FolderRow(string id)
        {
            return _rows.FirstOrDefault(x => x.IsFolder && x.FolderId == id);
        }

        // Right of the open/closed mark: a click left of this opens or closes the folder.
        private int FolderMarkRight()
        {
            return Px(InsetLeft + TextPadLeft + FolderMarkWidth + 4);
        }

        public int IndexFromPoint(Point p)
        {
            if (p.Y < 0 || p.Y >= ViewHeight || RowPx <= 0) return -1;
            int i = (p.Y + _scrollY) / RowPx;
            return i >= 0 && i < _rows.Count ? i : -1;
        }

        private Rectangle RowRect(int index)
        {
            return new Rectangle(0, index * RowPx - _scrollY, ContentWidth, RowPx);
        }

        private int ContentWidth { get { return ClientSize.Width - (_scroll.Visible ? _scroll.Width : 0); } }

        private Rectangle PlusButtonRect()
        {
            int top = Pinned ? ClientSize.Height - RowPx : _rows.Count * RowPx;
            int size = Px(PlusHoverSize);
            return new Rectangle((ClientSize.Width - size) / 2, top + (RowPx - size) / 2, size, size);
        }

        // An exception here would break the in-place activation of the whole task pane (the pane
        // would then be shown without being attached to Excel's windows and get no mouse input),
        // so the scroll bar is only touched with values that are valid for it.
        private void UpdateScrollBar()
        {
            int view = ViewHeight;
            int total = _rows.Count * RowPx;
            bool need = ClientSize.Height > 0 && Pinned && view > 0;
            _scroll.Visible = need;
            if (!need)
            {
                _scrollY = 0;
                return;
            }
            try
            {
                _scroll.SetBounds(ClientSize.Width - _scroll.Width, 0, _scroll.Width, view);
                _scroll.Minimum = 0;
                _scroll.LargeChange = view;
                _scroll.SmallChange = RowPx;
                _scroll.Maximum = total - 1;
                _scrollY = Math.Max(0, Math.Min(_scrollY, total - view));
                _scroll.Value = Math.Max(0, Math.Min(_scrollY, _scroll.Maximum - _scroll.LargeChange + 1));
            }
            catch (ArgumentException)
            {
                // The sizes are still settling during the first layout; the next resize fixes it.
            }
        }

        private void SetScroll(int y)
        {
            if (!_scroll.Visible) return;
            int max = Math.Max(0, _rows.Count * RowPx - ViewHeight);
            _scrollY = Math.Max(0, Math.Min(y, max));
            try
            {
                _scroll.Value = Math.Max(0, Math.Min(_scrollY, Math.Max(0, _scroll.Maximum - _scroll.LargeChange + 1)));
            }
            catch (ArgumentException)
            {
                // Out of range while the scroll bar is being resized.
            }
            Invalidate();
        }

        private void EnsureVisible(int index)
        {
            if (index < 0 || !_scroll.Visible) return;
            Rectangle r = RowRect(index);
            if (r.Top < 0) SetScroll(_scrollY + r.Top);
            else if (r.Bottom > ViewHeight) SetScroll(_scrollY + r.Bottom - ViewHeight);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateScrollBar();
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            UpdateScrollBar();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            SetScroll(_scrollY - Math.Sign(e.Delta) * RowPx);
        }

        // ---- drawing ----

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Theme.BarBack);
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.PixelOffsetMode = PixelOffsetMode.Half;

            // the tabs scroll inside the area above the pinned "+" row
            g.SetClip(new Rectangle(0, 0, ClientSize.Width, ViewHeight));
            DrawFolderOutlines(g);
            int first = Math.Max(0, _scrollY / Math.Max(1, RowPx));
            for (int i = first; i < _rows.Count; i++)
            {
                Rectangle row = RowRect(i);
                if (row.Top > ViewHeight) break;
                if (row.Bottom < 0) continue;
                if (_rows[i].IsFolder) DrawFolder(g, _rows[i], row);
                else DrawTab(g, _rows[i], row);
            }

            if (_dropIndex >= 0 && _rows.Count > 0)
            {
                int y = _dropIndex < _rows.Count ? RowRect(_dropIndex).Top : RowRect(_rows.Count - 1).Bottom;
                DrawDropMarker(g, y, _dropInside ? Px(FolderIndent) : 0);
            }
            g.ResetClip();

            DrawPlus(g);
        }

        // An open folder and its sheets are enclosed by a dashed line, white like a selected tab, with the
        // upper right and the lower left corners cut off. The line goes when the folder is closed.
        private void DrawFolderOutlines(Graphics g)
        {
            for (int i = 0; i < _rows.Count; i++)
            {
                if (!_rows[i].IsFolder || !_rows[i].Expanded) continue;
                int last = i;
                while (last + 1 < _rows.Count && _rows[last + 1].Indented) last++;
                if (last == i) continue; // no sheet of the folder is listed

                Rectangle top = RowRect(i), bottom = RowRect(last);
                if (bottom.Bottom < 0 || top.Top > ViewHeight) { i = last; continue; }
                int l = Px(1), r = ContentWidth - Px(OutlineRightInset), t = top.Top + Px(1), b = bottom.Bottom - Px(1);
                int cut = Math.Min(Px(OutlineCut), Math.Min(r - l, b - t) / 2);
                var points = new[]
                {
                    // the upper right and the lower left corners are cut off
                    new Point(l, t), new Point(r - cut, t), new Point(r, t + cut), new Point(r, b),
                    new Point(l + cut, b), new Point(l, b - cut)
                };
                // The dashes are edged with a dark line: the bar color, 75% darker. The edge is a wider pen
                // whose dashes are one pixel longer at both ends; both pens use the same dash pattern in pixels.
                int w = Math.Max(2, Px(1.2)), edge = 1, on = Px(5), off = Px(4);
                Color dark = Color.FromArgb(Theme.BarBack.R / 4, Theme.BarBack.G / 4, Theme.BarBack.B / 4);
                float wide = w + 2 * edge;
                using (var edgePen = new Pen(dark, wide))
                using (var pen = new Pen(Color.White, w))
                {
                    edgePen.LineJoin = LineJoin.Miter;
                    edgePen.DashPattern = new[] { (on + 2 * edge) / wide, Math.Max(0.5f, (off - 2 * edge) / wide) };
                    edgePen.DashOffset = edge / wide;
                    pen.LineJoin = LineJoin.Miter;
                    pen.DashPattern = new[] { on / (float)w, off / (float)w };
                    g.DrawPolygon(edgePen, points);
                    g.DrawPolygon(pen, points);
                }
                i = last;
            }
        }

        // Excel marks the place a dragged tab will land with a small black triangle; here it points
        // into the stack from the left edge.
        private void DrawDropMarker(Graphics g, int y, int x)
        {
            int w = Px(MarkerWidth), h = Px(MarkerHeight);
            var points = new[] { new Point(x, y - h / 2), new Point(x + w, y), new Point(x, y + h / 2) };
            using (var brush = new SolidBrush(Theme.DropMarker)) g.FillPolygon(brush, points);
        }

        private void DrawPlus(Graphics g)
        {
            bool enabled = !StructureProtected;
            Rectangle box = PlusButtonRect();
            if (_plusHot && enabled)
            {
                using (GraphicsPath path = RoundedRect(box, Px(PlusHoverRadius)))
                using (var brush = new SolidBrush(Theme.PlusHoverBack))
                {
                    g.FillPath(brush, path);
                }
            }
            int arm = Px(PlusArm), thick = Math.Max(1, Px(PlusThickness));
            int cx = box.Left + box.Width / 2, cy = box.Top + box.Height / 2;
            g.SmoothingMode = SmoothingMode.None;
            using (var brush = new SolidBrush(enabled ? Theme.PlusColor : Color.FromArgb(166, 166, 166)))
            {
                g.FillRectangle(brush, cx - arm / 2, cy - thick / 2, arm, thick);
                g.FillRectangle(brush, cx - thick / 2, cy - arm / 2, thick, arm);
            }
            g.SmoothingMode = SmoothingMode.AntiAlias;
        }

        // A folder: the open/closed mark, then the name in bold; a folder color shows as a bar at the left edge.
        private void DrawFolder(Graphics g, StripRow folder, Rectangle row)
        {
            int left = row.Left + Px(InsetLeft);
            int right = row.Right - RightInset(folder);
            int radius = Px(CornerRadius);
            bool selected = _selFolders.Contains(folder.FolderId);
            bool colored = folder.ColorBgr >= 0;
            Color color = colored
                ? Color.FromArgb(folder.ColorBgr & 0xFF, (folder.ColorBgr >> 8) & 0xFF, (folder.ColorBgr >> 16) & 0xFF)
                : Color.Empty;
            Color textColor = selected ? TextColor : Theme.InactiveText;

            // like a sheet tab: selected = white (with the color fading in from the top), not selected
            // with a color = a box in that color
            if (selected)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                using (GraphicsPath path = RoundedRect(body, radius))
                {
                    using (var brush = new SolidBrush(Color.White)) g.FillPath(brush, path);
                    if (colored)
                    {
                        Rectangle inner = Rectangle.Inflate(body, -1, -1);
                        Color bottom = Mix(Color.White, color, 0.45);
                        using (var brush = new LinearGradientBrush(new Rectangle(inner.X, inner.Y - 1, inner.Width, inner.Height + 2), Color.White, bottom, LinearGradientMode.Vertical))
                        using (GraphicsPath innerPath = RoundedRect(inner, Math.Max(1, radius - 1)))
                        {
                            g.FillPath(brush, innerPath);
                        }
                    }
                    using (var pen = new Pen(Theme.TabBorder, 1)) g.DrawPath(pen, path);
                }
            }
            else if (colored)
            {
                var body = new Rectangle(left, row.Top + Px(ColoredTop), right - left, row.Height - Px(ColoredTop) - Px(BottomInset));
                using (GraphicsPath path = RoundedRect(body, radius))
                using (var brush = new SolidBrush(color))
                {
                    g.FillPath(brush, path);
                }
                double luma = (color.R * 299 + color.G * 587 + color.B * 114) / 1000.0;
                textColor = luma >= 150 ? TextColor : Color.White;
            }

            // closed: a right-pointing chevron; open: the same turned down
            int cx = left + Px(TextPadLeft) / 2 + Px(FolderMarkWidth) / 2;
            int cy = row.Top + row.Height / 2;
            int w = Px(FolderMarkWidth), h = Px(FolderMarkHeight);
            Point[] chevron = folder.Expanded
                ? new[] { new Point(cx - h / 2, cy - w / 2), new Point(cx, cy + w / 2), new Point(cx + h / 2, cy - w / 2) }
                : new[] { new Point(cx - w / 2, cy - h / 2), new Point(cx + w / 2, cy), new Point(cx - w / 2, cy + h / 2) };
            using (var pen = new Pen(textColor, Math.Max(1.5f, Px(1.4))))
            {
                pen.LineJoin = LineJoin.Miter;
                g.DrawLines(pen, chevron);
            }

            int textLeft = left + Px(TextPadLeft) + Px(FolderMarkWidth) + Px(4);
            var textRect = new Rectangle(textLeft, row.Top + TextShift(folder.FolderName), Math.Max(1, right - textLeft - Px(TextPadRight)), row.Height);
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            if (_editor == null || _editingFolderId != folder.FolderId)
            {
                using (var bold = new Font(TabFont, FontStyle.Bold))
                using (var brush = new SolidBrush(textColor))
                using (var format = new StringFormat(StringFormat.GenericDefault))
                {
                    format.Alignment = StringAlignment.Near;
                    format.LineAlignment = StringAlignment.Center;
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                    g.DrawString(folder.FolderName, bold, brush, textRect, format);
                }
            }
            DrawSeparator(g, row);
        }

        // The text of a row is centered with the line box of its font, whose glyphs sit above the middle
        // of the box (Yu Gothic has a tall descent; Japanese characters sit higher than Latin ones). The
        // text is moved down so that its glyphs are centered in the tab body (measured: Latin 1.2 DIP,
        // text with Japanese characters 3.2 DIP at the font size used here).
        private int TextShift(string text)
        {
            bool wide = false;
            if (text != null)
            {
                foreach (char c in text) if (c >= '⺀') { wide = true; break; }
            }
            return Px(wide ? 3.2 : 1.2);
        }

        private void DrawSeparator(Graphics g, Rectangle row)
        {
            int sepHalf = Px(SeparatorLength) / 2;
            int cx = row.Left + row.Width / 2;
            using (var pen = new Pen(Theme.Separator, 1))
            {
                g.SmoothingMode = SmoothingMode.None;
                g.DrawLine(pen, cx - sepHalf, row.Bottom - 1, cx + sepHalf, row.Bottom - 1);
                g.SmoothingMode = SmoothingMode.AntiAlias;
            }
        }

        private void DrawTab(Graphics g, StripRow tabRow, Rectangle row)
        {
            string name = tabRow.Sheet;
            bool selected = _selected.Contains(name);
            bool active = selected && name == _active;
            int bgr;
            bool colored = _colors.TryGetValue(name, out bgr);
            Color tabColor = colored ? Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF) : Color.Empty;

            int left = row.Left + Px(InsetLeft) + (tabRow.Indented ? Px(FolderIndent) : 0);
            int right = row.Right - RightInset(tabRow);
            int radius = Px(CornerRadius);
            Color textColor = selected ? TextColor : Theme.InactiveText;

            // A colored tab that is not selected is a filled, rounded box; every other tab is white
            // when selected and has no background otherwise.
            if (selected)
            {
                var body = new Rectangle(left, row.Top + Px(colored ? PlainTop : PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                using (GraphicsPath path = RoundedRect(body, radius))
                {
                    using (var brush = new SolidBrush(Color.White)) g.FillPath(brush, path);
                    if (colored)
                    {
                        // white at the top, a mix of the tab color (about 45%) at the bottom
                        Rectangle inner = Rectangle.Inflate(body, -1, -1);
                        Color bottom = Mix(Color.White, tabColor, 0.45);
                        using (var brush = new LinearGradientBrush(new Rectangle(inner.X, inner.Y - 1, inner.Width, inner.Height + 2), Color.White, bottom, LinearGradientMode.Vertical))
                        using (GraphicsPath innerPath = RoundedRect(inner, Math.Max(1, radius - 1)))
                        {
                            g.FillPath(brush, innerPath);
                        }
                    }
                    using (var pen = new Pen(Theme.TabBorder, 1)) g.DrawPath(pen, path);
                }
                if (active)
                {
                    int thick = Math.Max(1, Px(UnderlineThickness));
                    int ux = body.Left + Px(UnderlineInset);
                    int uw = body.Width - 2 * Px(UnderlineInset);
                    int uy = body.Bottom - 1 - Px(UnderlineGap) - thick;
                    using (var brush = new SolidBrush(Underline)) g.FillRectangle(brush, ux, uy, Math.Max(1, uw), thick);
                }
            }
            else if (colored)
            {
                var body = new Rectangle(left, row.Top + Px(ColoredTop), right - left, row.Height - Px(ColoredTop) - Px(BottomInset));
                using (GraphicsPath path = RoundedRect(body, radius))
                using (var brush = new SolidBrush(tabColor))
                {
                    g.FillPath(brush, path);
                }
                // dark text on light tab colors, white on dark ones
                double luma = (tabColor.R * 299 + tabColor.G * 587 + tabColor.B * 114) / 1000.0;
                textColor = luma >= 150 ? TextColor : Color.White;
            }

            // the separator between tabs: Excel's thin line, turned to lie across the stack
            DrawSeparator(g, row);

            // text: left aligned, the active tab's text sits higher (the underline needs room)
            int lift = active ? Px(ActiveTextLift) : 0;
            var textRect = new Rectangle(left + Px(TextPadLeft), row.Top - lift + TextShift(name),Math.Max(1, right - left - Px(TextPadLeft) - Px(TextPadRight)), row.Height);
            if (_editor == null || name != _editingName)
            {
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                using (var brush = new SolidBrush(textColor))
                using (var format = new StringFormat(StringFormat.GenericDefault))
                {
                    format.Alignment = StringAlignment.Near;
                    format.LineAlignment = StringAlignment.Center;
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                    g.DrawString(name, TabFont, brush, textRect, format);
                }
            }
        }

        private static Color Mix(Color a, Color b, double t)
        {
            return Color.FromArgb(
                (int)Math.Round(a.R + (b.R - a.R) * t),
                (int)Math.Round(a.G + (b.G - a.G) * t),
                (int)Math.Round(a.B + (b.B - a.B) * t));
        }

        private static GraphicsPath RoundedRect(Rectangle r, int radius)
        {
            var path = new GraphicsPath();
            int d = Math.Max(1, radius * 2);
            Rectangle arc = new Rectangle(r.X, r.Y, d, d);
            path.AddArc(arc, 180, 90);
            arc.X = r.Right - 1 - d;
            path.AddArc(arc, 270, 90);
            arc.Y = r.Bottom - 1 - d;
            path.AddArc(arc, 0, 90);
            arc.X = r.X;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
        }

        // ---- selection (the same rules as Excel's own tabs) ----

        private void Raise()
        {
            if (_active == null || SelectionRequested == null) return;
            SelectionRequested(SelectedInOrder(), _active);
        }

        // Left button down: update the selection on screen, remember where a drag could start.
        // Excel is told when the button is released, unless a drag starts first.
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            _toggleOffFolder = null; // a drag swallows the button release, so nothing may be left over
            if (e.Button == MouseButtons.Left && !StructureProtected && PlusButtonRect().Contains(e.Location))
            {
                _plusDown = true;
                return;
            }
            int index = IndexFromPoint(e.Location);
            if (index < 0) return;
            StripRow row = _rows[index];
            bool ctrl = (ModifierKeys & Keys.Control) != 0;
            bool shift = (ModifierKeys & Keys.Shift) != 0;

            if (row.IsFolder)
            {
                if (e.Button == MouseButtons.Left && e.X <= FolderMarkRight())
                {
                    // the mark opens and closes the folder
                    if (FolderToggleRequested != null) FolderToggleRequested(row.FolderId);
                    return;
                }
                if (e.Button == MouseButtons.Left)
                {
                    _dragIndex = index;
                    _dragOrigin = e.Location;
                    if (_selFolders.Contains(row.FolderId) && !ctrl && !shift)
                    {
                        // a click on a selected folder deselects it, when the button is released
                        // without a drag (dragging a selected folder must keep the selection)
                        _toggleOffFolder = row.FolderId;
                        return;
                    }
                    SelectFolderByClick(row, index, ctrl, shift);
                    Invalidate();
                    return;
                }
                if (e.Button == MouseButtons.Right)
                {
                    _rightFolder = row.FolderId;
                    if (SheetMenu != null || FolderMenu != null) ContextMenuStrip = FolderMenu;
                    if (!_selFolders.Contains(row.FolderId)) SelectFolderByClick(row, index, false, false);
                    Invalidate();
                }
                return;
            }
            string name = row.Sheet;
            if (SheetMenu != null) ContextMenuStrip = SheetMenu;

            if (e.Button == MouseButtons.Left)
            {
                _dragIndex = index;
                _dragOrigin = e.Location;

                if (ctrl)
                {
                    _anchorFolder = null;
                    if (_selected.Contains(name) && _selected.Count > 1)
                    {
                        _selected.Remove(name);
                        if (_active == name) _active = SelectedInOrder().First();
                    }
                    else
                    {
                        _selected.Add(name);
                        _active = name;
                    }
                }
                else if (shift && (_anchorFolder != null || _names.IndexOf(_active) >= 0))
                {
                    SelectRange(index);
                    // the anchor stays
                }
                else if (_selected.Contains(name) && _selected.Count > 1 && _selected.Count < _names.Count)
                {
                    _anchorFolder = null;
                    _active = name; // inside a group: switch the active sheet, keep the group
                }
                else
                {
                    _anchorFolder = null;
                    _selFolders.Clear();
                    _selected = new HashSet<string> { name };
                    _active = name;
                }
                PruneFolders();
                _selectionPending = true;
                Invalidate();
                return;
            }

            if (e.Button != MouseButtons.Right) return;
            // Like Excel, a right click on a sheet outside the group ungroups and activates it;
            // inside the group it only makes that sheet the active one. Excel needs to know now,
            // because the menu commands act on the active sheet(s).
            if (!_selected.Contains(name)) { _selected = new HashSet<string> { name }; _selFolders.Clear(); }
            _active = name;
            PruneFolders();
            Invalidate();
            Raise();
        }

        // A click on a folder row: the folder and all its sheets become selected (Ctrl adds or removes
        // it, Shift takes the rows from the anchor to here).
        private void SelectFolderByClick(StripRow row, int index, bool ctrl, bool shift)
        {
            if (shift && (_anchorFolder != null || _names.IndexOf(_active) >= 0))
            {
                SelectRange(index);
                return;
            }
            _anchorFolder = row.FolderId;
            if (ctrl)
            {
                if (!_selFolders.Remove(row.FolderId)) _selFolders.Add(row.FolderId);
                return;
            }
            // The sheets inside a folder are not selected with it: the folder commands (right click)
            // act on all of them. Selecting a folder does not touch the selection of sheets.
            _selFolders = new HashSet<string> { row.FolderId };
        }

        // Selects every row from the anchor (the last clicked folder, otherwise the active sheet) to this one.
        private void SelectRange(int index)
        {
            int a = _anchorFolder != null ? _rows.FindIndex(r => r.IsFolder && r.FolderId == _anchorFolder) : RowIndexOf(_active);
            if (a < 0) a = index;
            int lo = Math.Min(a, index), hi = Math.Max(a, index);
            var sheets = new HashSet<string>();
            var folders = new HashSet<string>();
            for (int i = lo; i <= hi; i++)
            {
                StripRow r = _rows[i];
                if (r.IsFolder) folders.Add(r.FolderId);
                else sheets.Add(r.Sheet);
            }
            _selFolders = folders;
            if (sheets.Count == 0) return;
            _selected = sheets;
            if (!_selected.Contains(_active)) _active = SelectedInOrder().First();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            bool noDrag = _dragIndex >= 0;
            _dragIndex = -1;
            string toggleOff = _toggleOffFolder;
            _toggleOffFolder = null;
            if (toggleOff != null && noDrag && e.Button == MouseButtons.Left)
            {
                _selFolders.Remove(toggleOff);
                Invalidate();
                return;
            }
            if (_plusDown)
            {
                _plusDown = false;
                if (e.Button == MouseButtons.Left && PlusButtonRect().Contains(e.Location) && AddRequested != null) AddRequested();
                return;
            }
            if (e.Button != MouseButtons.Left || !_selectionPending) return;
            _selectionPending = false;
            Raise();
        }

        // Keys that Excel would handle if the pane did not have the focus (save, save as).
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            string command = keyData == (Keys.Control | Keys.S) ? "FileSave" : keyData == Keys.F12 ? "FileSaveAs" : null;
            if (command != null && _editor == null && ExcelCommandRequested != null)
            {
                ExcelCommandRequested(command);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // Up/Down arrow keys move the active sheet (a single selection).
        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Up || keyData == Keys.Down || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode == Keys.Delete)
            {
                e.Handled = true;
                if (!StructureProtected && DeleteKeyRequested != null) DeleteKeyRequested(_selFolders.Count > 0 ? new List<string>() : SelectedInOrder(), _selFolders.ToList());
                return;
            }
            if (e.KeyCode == Keys.F2)
            {
                e.Handled = true;
                if (_selFolders.Count == 1) BeginRenameFolder(_selFolders.First());
                else BeginRenameActive();
                return;
            }
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down) return;
            _selFolders.Clear();
            _anchorFolder = null;
            int i = RowIndexOf(_active) + (e.KeyCode == Keys.Down ? 1 : -1);
            while (i >= 0 && i < _rows.Count && _rows[i].IsFolder) i += e.KeyCode == Keys.Down ? 1 : -1;
            if (i < 0 || i >= _rows.Count) return;
            e.Handled = true;
            _selected = new HashSet<string> { _rows[i].Sheet };
            _active = _rows[i].Sheet;
            EnsureVisible(i);
            Invalidate();
            Raise();
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            SetPlusHot(false);
            _hoverIndex = -1;
            _tip.Hide(this);
        }

        private void UpdateTip(Point p)
        {
            int index = IndexFromPoint(p);
            if (index == _hoverIndex) return;
            _hoverIndex = index;
            _tip.Hide(this);
            if (index < 0 || Control.MouseButtons != MouseButtons.None) return;
            string text = _rows[index].IsFolder ? _rows[index].FolderName : _rows[index].Sheet;
            if (!IsTextCut(text, _rows[index])) return;
            Rectangle row = RowRect(index);
            _tip.Show(text, this, row.Left + Px(InsetLeft + TextPadLeft), row.Bottom, 4000);
        }

        // Whether the name does not fit the tab (it is then drawn with a trailing "…").
        private bool IsTextCut(string name, StripRow row)
        {
            int available = ContentWidth - Px(InsetLeft + TextPadLeft + TextPadRight) - RightInset(row)
                - (row.Indented ? Px(FolderIndent) : 0) - (row.IsFolder ? Px(FolderMarkWidth + 4) : 0);
            using (Graphics g = CreateGraphics())
            using (var format = new StringFormat(StringFormat.GenericDefault))
            {
                format.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                return g.MeasureString(name, TabFont, int.MaxValue, format).Width > available;
            }
        }

        private void SetPlusHot(bool hot)
        {
            if (_plusHot == hot) return;
            _plusHot = hot;
            Invalidate();
        }

        // ---- drag and drop ----

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            SetPlusHot(!StructureProtected && PlusButtonRect().Contains(e.Location));
            UpdateTip(e.Location);
            if (e.Button != MouseButtons.Left || _dragIndex < 0 || _editor != null || StructureProtected) return;
            Size drag = SystemInformation.DragSize;
            if (Math.Abs(e.X - _dragOrigin.X) < drag.Width && Math.Abs(e.Y - _dragOrigin.Y) < drag.Height) return;

            int index = _dragIndex;
            _dragIndex = -1;
            _selectionPending = false; // a drag does not change the selection in Excel by itself
            _toggleOffFolder = null;
            if (index >= _rows.Count) return;

            // Dragging something selected takes the whole selection along (sheets and folders).
            StripRow row = _rows[index];
            // A folder row takes the selected folders (with their sheets); a sheet row takes the selected sheets.
            if (row.IsFolder)
            {
                _dragSheets = new List<string>();
                _dragFolders = _selFolders.ToList();
                if (!_dragFolders.Contains(row.FolderId)) _dragFolders.Add(row.FolderId);
            }
            else
            {
                _dragSheets = _selected.Contains(row.Sheet) ? SelectedInOrder() : new List<string> { row.Sheet };
                _dragFolders = new List<string>();
            }
            _dragNames = _dragSheets;
            _dragging = true;
            try
            {
                DoDragDrop(row.IsFolder ? row.FolderName : row.Sheet, DragDropEffects.Move | DragDropEffects.Copy);
            }
            finally
            {
                _dragging = false;
                SetDropIndex(-1);
            }
        }

        // Ctrl held: copy, otherwise move (the same keys as Excel's own tabs).
        private static bool IsCopy(DragEventArgs e)
        {
            return (e.KeyState & 8) != 0;
        }

        protected override void OnDragOver(DragEventArgs e)
        {
            base.OnDragOver(e);
            if (!_dragging || !e.Data.GetDataPresent(typeof(string)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            e.Effect = IsCopy(e) ? DragDropEffects.Copy : DragDropEffects.Move;

            Point p = PointToClient(new Point(e.X, e.Y));
            ScrollNearEdge(p);
            if (HasFolders) ResolveFolderDrop(p);
            else { _dropInside = false; _dropInto = _dropBeforeSheet = _dropBeforeFolder = null; SetDropIndex(InsertIndexAt(p)); }
        }

        protected override void OnDragLeave(EventArgs e)
        {
            base.OnDragLeave(e);
            SetDropIndex(-1);
        }

        protected override void OnDragDrop(DragEventArgs e)
        {
            base.OnDragDrop(e);
            int target = _dropIndex;
            string into = _dropInto, beforeSheet = _dropBeforeSheet, beforeFolder = _dropBeforeFolder;
            SetDropIndex(-1);
            if (!_dragging || target < 0 || _dragNames.Count + _dragFolders.Count == 0) return;

            if (HasFolders)
            {
                if (FolderMoveRequested != null) FolderMoveRequested(_dragSheets, _dragFolders, into, beforeSheet, beforeFolder, IsCopy(e));
                return;
            }
            bool copy = IsCopy(e);

            // Dropping the sheets where they already are changes nothing (when they sit together).
            List<int> at = _dragNames.Select(n => _names.IndexOf(n)).OrderBy(i => i).ToList();
            if (at[0] < 0) return;
            bool together = at[at.Count - 1] - at[0] + 1 == at.Count;
            if (!copy && together && target >= at[0] && target <= at[at.Count - 1] + 1) return;

            // The sheets cannot be dropped in front of one of themselves: take the next other one.
            while (target < _names.Count && _dragNames.Contains(_names[target])) target++;
            if (MoveRequested != null) MoveRequested(_dragNames, target < _names.Count ? _names[target] : null, copy);
        }

        private void SetDropIndex(int index)
        {
            if (_dropIndex == index) return;
            _dropIndex = index;
            Invalidate();
        }

        // Whether a dragged sheet comes from the folder of this folder row. Dropped below the last sheet
        // of its own folder a sheet leaves the folder (to put a sheet at the end of a folder, drag it
        // from outside the folder, or drop it between the sheets and move it).
        private bool FromFolderOfRow(int folderRow)
        {
            if (folderRow < 0 || folderRow >= _rows.Count || !_rows[folderRow].IsFolder) return false;
            return _dragSheets.Any(s => _rows[folderRow].Members.Contains(s));
        }

        // The row that owns an indented row: the folder row above it.
        private int OwnerRow(int index)
        {
            for (int i = index; i >= 0; i--) if (_rows[i].IsFolder) return i;
            return -1;
        }

        // Works out where the dragged things would land. Sheets may go into a folder (between its
        // sheets, or at its end when the pointer is right of the indent), folders only to the top level.
        private void ResolveFolderDrop(Point p)
        {
            bool foldersDragged = _dragFolders.Count > 0;
            int b = InsertIndexAt(p);
            StripRow prev = b > 0 ? _rows[b - 1] : null;
            StripRow next = b < _rows.Count ? _rows[b] : null;
            bool right = p.X >= Px(InsetLeft + FolderIndent / 2);
            bool inside = false;
            string into = null, beforeSheet = null, beforeFolder = null;

            if (next != null && next.Indented)
            {
                int owner = OwnerRow(b);
                if (!foldersDragged)
                {
                    inside = true;
                    into = _rows[owner].FolderId;
                    beforeSheet = next.Sheet;
                }
                else if (prev != null && prev.IsFolder)
                {
                    b = owner; // in front of the folder
                    beforeFolder = _rows[owner].FolderId;
                }
                else
                {
                    int after = b;
                    while (after < _rows.Count && _rows[after].Indented) after++;
                    b = after;
                    next = after < _rows.Count ? _rows[after] : null;
                    if (next != null) { if (next.IsFolder) beforeFolder = next.FolderId; else beforeSheet = next.Sheet; }
                }
            }
            else if (!foldersDragged && right && prev != null && (prev.Indented || (prev.IsFolder && prev.Expanded))
                && !FromFolderOfRow(prev.Indented ? OwnerRow(b - 1) : -1))
            {
                inside = true;
                into = _rows[prev.Indented ? OwnerRow(b - 1) : b - 1].FolderId;
            }
            else if (next != null)
            {
                if (next.IsFolder) beforeFolder = next.FolderId; else beforeSheet = next.Sheet;
            }

            _dropInside = inside;
            _dropInto = into;
            _dropBeforeSheet = beforeSheet;
            _dropBeforeFolder = beforeFolder;
            if (_dropIndex == b) Invalidate();
            SetDropIndex(b);
        }

        // The position (0..count) whose boundary is nearest: upper half of a row = before it.
        private int InsertIndexAt(Point p)
        {
            int index = IndexFromPoint(p);
            if (index < 0) return _rows.Count;
            Rectangle r = RowRect(index);
            return p.Y < r.Top + r.Height / 2 ? index : index + 1;
        }

        // DoDragDrop does not scroll the control, so do it near the top and bottom edges.
        private void ScrollNearEdge(Point p)
        {
            int edge = Math.Max(8, RowPx / 2);
            if (p.Y < edge) SetScroll(_scrollY - RowPx / 2);
            else if (p.Y > ViewHeight - edge) SetScroll(_scrollY + RowPx / 2);
        }

        // ---- renaming ----

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            base.OnMouseDoubleClick(e);
            if (e.Button != MouseButtons.Left) return;
            int index = IndexFromPoint(e.Location);
            if (index < 0) return;
            if (_rows[index].IsFolder)
            {
                if (e.X > FolderMarkRight()) BeginRenameFolder(_rows[index].FolderId);
                return;
            }
            BeginRename(index);
        }

        public void BeginRenameActive()
        {
            BeginRename(RowIndexOf(_active));
        }

        public void BeginRenameFolder(string folderId)
        {
            int index = _rows.FindIndex(r => r.IsFolder && r.FolderId == folderId);
            if (index < 0) return;
            EnsureVisible(index);
            BeginEdit(index, _rows[index].FolderName, folderId);
        }

        // In-place name editing: Enter or leaving the box commits, Esc cancels. The length limit
        // is Excel's 31 characters; the other rules are checked by PaneManager when it renames.
        private void BeginRename(int index)
        {
            if (index < 0 || index >= _rows.Count || _rows[index].IsFolder || StructureProtected) return;
            BeginEdit(index, _rows[index].Sheet, null);
        }

        private void BeginEdit(int index, string text, string folderId)
        {
            EndRename(false);
            Rectangle r = RowRect(index);
            _editingName = text;
            _editingFolderId = folderId;
            int left = r.Left + Px(InsetLeft) + (_rows[index].Indented ? Px(FolderIndent) : 0);
            int right = r.Right - RightInset(_rows[index]);
            int textLeft = folderId != null ? left + Px(TextPadLeft) + Px(FolderMarkWidth) + Px(4) - 2 : left + Px(TextPadLeft) - 2;
            _editor = new TextBox
            {
                Text = text,
                MaxLength = 31,
                BorderStyle = BorderStyle.FixedSingle,
                Font = TabFont
            };
            _editor.Bounds = new Rectangle(textLeft, r.Top + (r.Height - _editor.PreferredHeight) / 2,
                Math.Max(Px(40), right - textLeft), _editor.PreferredHeight);
            _editor.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; EndRename(true); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; EndRename(false); }
            };
            _editor.LostFocus += (s, e) => EndRename(true);
            Controls.Add(_editor);
            _editor.Focus();
            _editor.SelectAll();
            Invalidate();
        }

        private void EndRename(bool commit)
        {
            TextBox editor = _editor;
            if (editor == null) return;
            _editor = null; // first, so the LostFocus raised by removing the box does nothing
            string oldName = _editingName;
            string folderId = _editingFolderId;
            _editingFolderId = null;
            string newName = editor.Text;
            Controls.Remove(editor);
            editor.Dispose();
            Invalidate();
            if (!commit || newName == oldName) return;
            if (folderId != null) { if (FolderRenameRequested != null) FolderRenameRequested(folderId, newName); }
            else if (RenameRequested != null) RenameRequested(oldName, newName);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (_font != null) _font.Dispose();
                _tip.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
