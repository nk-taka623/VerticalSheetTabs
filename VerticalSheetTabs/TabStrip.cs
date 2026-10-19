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
        // Excel's tab text is about 9pt "Yu Gothic" (the measured width of "Sheet1" matches it).
        private const string FontFamilyName = "Yu Gothic";
        private const double FontPixelsAt96 = 15.5;
        // The "+" (new sheet) button: a plain plus sign that gets a light rounded square on hover.
        private const double PlusArm = 10.5, PlusThickness = 1.6, PlusHoverSize = 24, PlusHoverRadius = 3.2;
        // The drop marker while dragging: Excel shows a small black triangle at the insertion point.
        private const double MarkerWidth = 4, MarkerHeight = 8;

        private List<string> _names = new List<string>();
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

        // First-open width: 16 half-width characters plus the tab's own paddings.
        public int DefaultContentWidth()
        {
            Size text = TextRenderer.MeasureText(new string('0', 16), TabFont, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding);
            return text.Width + Px(InsetLeft + InsetRight + TextPadLeft + TextPadRight);
        }

        // ---- data from the pane ----

        // Updates what changed only, so the polling timer does not cause flicker.
        public void SetSheets(IList<string> names, string activeName, ICollection<string> selected, IDictionary<string, int> colors)
        {
            RefreshTheme();
            bool sameColors = _colors.Count == colors.Count && _colors.All(kv => { int v; return colors.TryGetValue(kv.Key, out v) && v == kv.Value; });
            bool sameNames = _names.SequenceEqual(names);
            // While a mouse button is down the user is changing the selection; do not fight them.
            bool userBusy = Control.MouseButtons != MouseButtons.None;
            bool sameSelection = userBusy || (_active == activeName && _selected.SetEquals(selected));
            if (sameNames && sameSelection && sameColors) return;

            if (!sameColors)
            {
                _colors.Clear();
                foreach (var kv in colors) _colors[kv.Key] = kv.Value;
            }
            if (!sameNames)
            {
                EndRename(false);
                _names = new List<string>(names);
                UpdateScrollBar();
            }
            if (!userBusy)
            {
                bool activeChanged = _active != activeName;
                _active = activeName;
                _selected = new HashSet<string>(selected);
                if (activeChanged) EnsureVisible(_names.IndexOf(_active));
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
        private bool Pinned { get { return (_names.Count + 1) * RowPx > ClientSize.Height; } }

        private int ViewHeight { get { return Pinned ? Math.Max(0, ClientSize.Height - RowPx) : ClientSize.Height; } }

        public int IndexFromPoint(Point p)
        {
            if (p.Y < 0 || p.Y >= ViewHeight || RowPx <= 0) return -1;
            int i = (p.Y + _scrollY) / RowPx;
            return i >= 0 && i < _names.Count ? i : -1;
        }

        private Rectangle RowRect(int index)
        {
            return new Rectangle(0, index * RowPx - _scrollY, ContentWidth, RowPx);
        }

        private int ContentWidth { get { return ClientSize.Width - (_scroll.Visible ? _scroll.Width : 0); } }

        private Rectangle PlusButtonRect()
        {
            int top = Pinned ? ClientSize.Height - RowPx : _names.Count * RowPx;
            int size = Px(PlusHoverSize);
            return new Rectangle((ClientSize.Width - size) / 2, top + (RowPx - size) / 2, size, size);
        }

        // An exception here would break the in-place activation of the whole task pane (the pane
        // would then be shown without being attached to Excel's windows and get no mouse input),
        // so the scroll bar is only touched with values that are valid for it.
        private void UpdateScrollBar()
        {
            int view = ViewHeight;
            int total = _names.Count * RowPx;
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
            int max = Math.Max(0, _names.Count * RowPx - ViewHeight);
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
            int first = Math.Max(0, _scrollY / Math.Max(1, RowPx));
            for (int i = first; i < _names.Count; i++)
            {
                Rectangle row = RowRect(i);
                if (row.Top > ViewHeight) break;
                if (row.Bottom < 0) continue;
                DrawTab(g, i, row);
            }

            if (_dropIndex >= 0 && _names.Count > 0)
            {
                int y = _dropIndex < _names.Count ? RowRect(_dropIndex).Top : RowRect(_names.Count - 1).Bottom;
                DrawDropMarker(g, y);
            }
            g.ResetClip();

            DrawPlus(g);
        }

        // Excel marks the place a dragged tab will land with a small black triangle; here it points
        // into the stack from the left edge.
        private void DrawDropMarker(Graphics g, int y)
        {
            int w = Px(MarkerWidth), h = Px(MarkerHeight);
            var points = new[] { new Point(0, y - h / 2), new Point(w, y), new Point(0, y + h / 2) };
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

        private void DrawTab(Graphics g, int index, Rectangle row)
        {
            string name = _names[index];
            bool selected = _selected.Contains(name);
            bool active = selected && name == _active;
            int bgr;
            bool colored = _colors.TryGetValue(name, out bgr);
            Color tabColor = colored ? Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF) : Color.Empty;

            int left = row.Left + Px(InsetLeft);
            int right = row.Right - Px(InsetRight);
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
            int sepHalf = Px(SeparatorLength) / 2;
            int cx = row.Left + row.Width / 2;
            using (var pen = new Pen(Theme.Separator, 1))
            {
                g.SmoothingMode = SmoothingMode.None;
                g.DrawLine(pen, cx - sepHalf, row.Bottom - 1, cx + sepHalf, row.Bottom - 1);
                g.SmoothingMode = SmoothingMode.AntiAlias;
            }

            // text: left aligned, the active tab's text sits higher (the underline needs room)
            int lift = active ? Px(ActiveTextLift) : 0;
            var textRect = new Rectangle(left + Px(TextPadLeft), row.Top - lift, Math.Max(1, right - left - Px(TextPadLeft) - Px(TextPadRight)), row.Height);
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
            if (e.Button == MouseButtons.Left && !StructureProtected && PlusButtonRect().Contains(e.Location))
            {
                _plusDown = true;
                return;
            }
            int index = IndexFromPoint(e.Location);
            if (index < 0) return;
            string name = _names[index];

            if (e.Button == MouseButtons.Left)
            {
                _dragIndex = index;
                _dragOrigin = e.Location;

                if ((ModifierKeys & Keys.Control) != 0)
                {
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
                else if ((ModifierKeys & Keys.Shift) != 0 && _names.IndexOf(_active) >= 0)
                {
                    int a = _names.IndexOf(_active);
                    int lo = Math.Min(a, index), hi = Math.Max(a, index);
                    _selected = new HashSet<string>(_names.Skip(lo).Take(hi - lo + 1));
                    // the anchor stays the active sheet
                }
                else if (_selected.Contains(name) && _selected.Count > 1 && _selected.Count < _names.Count)
                {
                    _active = name; // inside a group: switch the active sheet, keep the group
                }
                else
                {
                    _selected = new HashSet<string> { name };
                    _active = name;
                }
                _selectionPending = true;
                Invalidate();
                return;
            }

            if (e.Button != MouseButtons.Right) return;
            // Like Excel, a right click on a sheet outside the group ungroups and activates it;
            // inside the group it only makes that sheet the active one. Excel needs to know now,
            // because the menu commands act on the active sheet(s).
            if (!_selected.Contains(name)) _selected = new HashSet<string> { name };
            _active = name;
            Invalidate();
            Raise();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragIndex = -1;
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

        // Up/Down arrow keys move the active sheet (a single selection).
        protected override bool IsInputKey(Keys keyData)
        {
            return keyData == Keys.Up || keyData == Keys.Down || base.IsInputKey(keyData);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down) return;
            int i = _names.IndexOf(_active) + (e.KeyCode == Keys.Down ? 1 : -1);
            if (i < 0 || i >= _names.Count) return;
            e.Handled = true;
            _selected = new HashSet<string> { _names[i] };
            _active = _names[i];
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
            if (index < 0 || Control.MouseButtons != MouseButtons.None || !IsTextCut(_names[index])) return;
            Rectangle row = RowRect(index);
            _tip.Show(_names[index], this, row.Left + Px(InsetLeft + TextPadLeft), row.Bottom, 4000);
        }

        // Whether the name does not fit the tab (it is then drawn with a trailing "…").
        private bool IsTextCut(string name)
        {
            int available = ContentWidth - Px(InsetLeft + InsetRight + TextPadLeft + TextPadRight);
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
            if (index >= _names.Count) return;

            // Dragging a sheet of the group takes the whole group along.
            string name = _names[index];
            _dragNames = _selected.Contains(name) ? SelectedInOrder() : new List<string> { name };
            _dragging = true;
            try
            {
                DoDragDrop(name, DragDropEffects.Move | DragDropEffects.Copy);
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
            SetDropIndex(InsertIndexAt(p));
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
            SetDropIndex(-1);
            if (!_dragging || target < 0 || _dragNames.Count == 0) return;
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

        // The position (0..count) whose boundary is nearest: upper half of a row = before it.
        private int InsertIndexAt(Point p)
        {
            int index = IndexFromPoint(p);
            if (index < 0) return _names.Count;
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
            if (index >= 0) BeginRename(index);
        }

        public void BeginRenameActive()
        {
            BeginRename(_names.IndexOf(_active));
        }

        // In-place name editing: Enter or leaving the box commits, Esc cancels. The length limit
        // is Excel's 31 characters; the other rules are checked by PaneManager when it renames.
        private void BeginRename(int index)
        {
            if (index < 0 || index >= _names.Count || StructureProtected) return;
            EndRename(false);
            Rectangle r = RowRect(index);
            _editingName = _names[index];
            int left = r.Left + Px(InsetLeft);
            int right = r.Right - Px(InsetRight);
            _editor = new TextBox
            {
                Text = _editingName,
                MaxLength = 31,
                BorderStyle = BorderStyle.FixedSingle,
                Font = TabFont
            };
            _editor.Bounds = new Rectangle(left + Px(TextPadLeft) - 2, r.Top + (r.Height - _editor.PreferredHeight) / 2,
                Math.Max(Px(40), right - left - Px(TextPadLeft)), _editor.PreferredHeight);
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
            string newName = editor.Text;
            Controls.Remove(editor);
            editor.Dispose();
            Invalidate();
            if (commit && newName != oldName && RenameRequested != null) RenameRequested(oldName, newName);
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
