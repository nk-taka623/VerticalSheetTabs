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
        // The text drawn for a sheet when it differs from its name (the sheet-name rule: "B" for "A | B").
        public string Display;
        // The longest name that may be typed when renaming this row; 0 = Excel's 31.
        public int EditLimit;
        public bool Indented;
        public string FolderId;
        public string FolderName;
        public bool Expanded;
        public int ColorBgr = -1;
        // Folder: its sheets all have the folder's color (a check mark at the right end of the row).
        public bool SyncColor;
        // Inside the dashed outline of an open folder (the folder row and its sheet rows): these rows end
        // before the outline by the width of its cut corner.
        public bool InGroup;
        // Folder: its visible sheets (also while it is closed).
        public List<string> Members = new List<string>();

        public bool IsFolder { get { return FolderId != null; } }

        public string Text { get { return Display ?? Sheet; } }

        public string Signature()
        {
            return IsFolder ? "F|" + FolderId + "|" + FolderName + "|" + Expanded + "|" + ColorBgr + "|" + SyncColor + "|" + EditLimit + "|" + string.Join("|", Members)
                : "S|" + Sheet + "|" + Indented + "|" + Display + "|" + EditLimit;
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
        private const double ActiveTextLift = 1.6;
        // The standard type (Excel's own look): the underline of the active sheet, and its text sitting higher.
        private const double UnderlineInset = 8, UnderlineThickness = 1.6, UnderlineGap = 3.2, StandardTextLift = 3.2;
        // The triangles at the ends of a selected row: distance of their center from the edge of the body, size,
        // and the room the text keeps free at each end.
        private const double MarkerInset = 13, MarkerHalfHeight = 5, MarkerBack = 4, MarkerFront = 3, MarkerRoom = 15;
        // The bars at the ends of the active sheet: width, distance from the ends of the body and from its top and
        // bottom, and the room the text keeps free at each end.
        private const double BarWidth = 6, BarInset = 9, BarMargin = 4, BarRoom = 12, BarScale = 0.75; // 0.75 times the first size (8 wide); moved in by one bar width (Agents.md 12.132)
        // the check badge marker at the left of the name of the active sheet: the room it takes before the text (Agents.md 12.131)
        private const double BadgeRoom = 12, BadgeScale = 0.75; // the badge is 0.75 times the first size (Agents.md 12.132)
        private const double SeparatorLength = 19;
        private const double TextPadLeft = 12, TextPadRight = 8;
        // Folders: sheets inside a folder are indented; the folder row has an open/closed mark and,
        // with a color, a bar at its left edge.
        private const double OutlineRightInset = 1.5, OutlineCut = 10;
        // The check mark of a folder that keeps its sheets in its color: size, space to the right end, gap to the text.
        private const double CheckWidth = 10, CheckHeight = 7, CheckPadRight = 8, CheckGap = 4, CheckHalo = 3;
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
        // The full name of the sheet being renamed (the edited text may be only the part after "A | ").
        private string _editingSheet;
        private string _rightFolder;
        private string _toggleOffFolder;
        private bool _plusHot, _plusDown;
        // The full name of a tab whose text is cut off is shown as a tool tip (Excel's own tabs
        // grow to fit the name, so they never need one).
        private readonly ToolTip _tip = new ToolTip { ShowAlways = true, InitialDelay = 500, ReshowDelay = 100 };
        private int _hoverIndex = -1;

        // (selected sheets in tab order, the active one) after the user changed the selection.
        public event Action<IList<string>, string> SelectionRequested;
        // Raised after a click with the left button has selected sheets: Excel's own tabs hand the keyboard
        // focus back to the sheet, so that typing goes into the cell of the sheet just chosen.
        // The argument tells whether the pane had the keyboard focus when the click was made.
        public event Action<bool> GridFocusRequested;
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

        // The "+" is fixed at the bottom of the pane, so that it does not move down as sheets are added
        // and can be pressed repeatedly. The tabs scroll above it when they do not fit.
        private bool NeedScroll { get { return _rows.Count * RowPx > ViewHeight; } }

        private int ViewHeight { get { return IsSample ? ClientSize.Height : Math.Max(0, ClientSize.Height - RowPx); } }

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
        // The width the check mark takes at the right end of a folder row (the name stays clear of it).
        private int CheckReserve(StripRow r)
        {
            return r.IsFolder && r.SyncColor ? Px(CheckWidth + CheckPadRight + CheckGap) : 0;
        }

        private int RightInset(StripRow r)
        {
            return r.InGroup && _outlineVisible ? Px(OutlineRightInset + OutlineCut) : Px(InsetRight);
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
            int top = ClientSize.Height - RowPx;
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
            bool need = ClientSize.Height > 0 && NeedScroll && view > 0;
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
            if (IsSample) return;
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

            if (!IsSample) DrawPlus(g);
        }

        public const int OutlineLevels = 9;

        private bool _outlineVisible = true;
        private int _outlineLevel = 5;

        // Whether the dashed outline around an open folder is drawn, and its color level (1..9).
        // How selected folders and sheets look: Excel's own look (white, with the underline), or the double frame.
        private SelectStyle _selectStyle = SelectStyle.Standard;

        public SelectStyle SelectionStyle
        {
            get { return _selectStyle; }
            set { if (_selectStyle != value) { _selectStyle = value; Invalidate(); } }
        }

        private bool DoubleFrame { get { return _selectStyle == SelectStyle.DoubleFrame; } }

        private bool Inset { get { return _selectStyle == SelectStyle.Inset; } }

        private bool Stripes { get { return _selectStyle == SelectStyle.Stripes; } }

        private bool Gradient { get { return _selectStyle == SelectStyle.GradientFrame; } }

        // The mark of the active sheet, with every look (Agents.md 12.110).
        private ActiveMarker _activeMarker = ActiveMarker.Underline;

        public ActiveMarker ActiveMarker
        {
            get { return _activeMarker; }
            set { if (_activeMarker != value) { _activeMarker = value; Invalidate(); } }
        }

        private bool Glass { get { return _selectStyle == SelectStyle.FrostedGlass || _selectStyle == SelectStyle.GlossyGlass || _selectStyle == SelectStyle.FrostedGlass2; } }

        // A sample (in the settings dialog): the same drawing, three fixed rows, no "+" button, no operation.
        public bool IsSample { get; private set; }

        // The size of a sample of the given width (in DIP): three rows.
        public Size SampleSize(int widthDip)
        {
            return new Size(Px(widthDip), 3 * RowPx + Px(2));
        }

        // A click on a sample.
        public event Action SampleClicked;

        // Three rows: a folder, Sheet1, and Sheet2 in orange. What is selected: "none", "plain" (the folder, and
        // Sheet1 as the active sheet) or "colored" (Sheet2 as the active sheet).
        public void SetSample(SelectStyle style, string selection, ActiveMarker marker)
        {
            _activeMarker = marker;
            IsSample = true;
            TabStop = false;
            _selectStyle = style;
            var rows = new List<StripRow>
            {
                new StripRow { FolderId = "sample", FolderName = "資料", Expanded = true },
                new StripRow { Sheet = "Sheet1" },
                new StripRow { Sheet = "Sheet2" }
            };
            var names = new List<string> { "Sheet1", "Sheet2" };
            var colors = new Dictionary<string, int> { { "Sheet2", 0x3A80E9 } };
            string active = selection == "plain" ? "Sheet1" : selection == "colored" ? "Sheet2" : "";
            var selected = new List<string>();
            if (active != "") selected.Add(active);
            SetSheets(names, active, selected, colors, rows);
            _selFolders.Clear();
            if (selection == "plain") _selFolders.Add("sample");
            Invalidate();
        }

        public bool OutlineVisible
        {
            get { return _outlineVisible; }
            set { if (_outlineVisible != value) { _outlineVisible = value; Invalidate(); } }
        }

        public int OutlineLevel
        {
            get { return _outlineLevel; }
            set
            {
                int v = Math.Max(1, Math.Min(OutlineLevels, value));
                if (_outlineLevel != v) { _outlineLevel = v; Invalidate(); }
            }
        }

        // Level 9 is the strongest color, the opposite of the bar background on the hue circle; the
        // levels below it blend from the background (level 0) to that color in 9 equal steps.
        public static Color OutlineColorAt(Color bg, int level)
        {
            Color full = OppositeColor(bg);
            double t = Math.Max(0, Math.Min(OutlineLevels, level)) / (double)OutlineLevels;
            return Color.FromArgb(
                (int)Math.Round(bg.R + (full.R - bg.R) * t),
                (int)Math.Round(bg.G + (full.G - bg.G) * t),
                (int)Math.Round(bg.B + (full.B - bg.B) * t));
        }

        // The opposite of the bar background on the hue circle (hue + 180 degrees).
        // A gray background has no hue to turn, so the saturation is raised and the lightness is pushed
        // away from the background's, to keep the line visible.
        private static Color OppositeColor(Color bg)
        {
            double h = bg.GetHue(), s = bg.GetSaturation(), l = bg.GetBrightness();
            h = (h + 180) % 360;
            s = Math.Max(s, 0.85);
            l = l > 0.5 ? 0.35 : 0.65;
            return FromHsl(h, s, l);
        }

        // A color from hue (degrees), saturation and lightness (0..1).
        private static Color FromHsl(double h, double s, double l)
        {
            double q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
            Func<double, int> ch = t =>
            {
                t = ((t % 1) + 1) % 1;
                double v = t < 1.0 / 6 ? p + (q - p) * 6 * t : t < 0.5 ? q : t < 2.0 / 3 ? p + (q - p) * (2.0 / 3 - t) * 6 : p;
                return (int)Math.Round(v * 255);
            };
            return Color.FromArgb(ch(h / 360 + 1.0 / 3), ch(h / 360), ch(h / 360 - 1.0 / 3));
        }

        // The colors of a selected row's lines and glow come from the color of the row: the opposite on the hue
        // circle (hue + 180 degrees), which differs from it the most. A row without a color (white) or a gray one has
        // no hue to turn: Excel's green is used then, as before.
        private static bool HasHue(Color c, bool colored)
        {
            return colored && c.GetSaturation() >= 0.12;
        }

        // The color of the triangles: dark on a light color, light on a dark one (OppositeColor).
        private static Color LineColorFor(Color rowColor, bool colored)
        {
            return Visible(HasHue(rowColor, colored) ? OppositeColor(rowColor) : Underline, colored ? rowColor : Color.White);
        }

        // Marks must stand out from the row they are on (Agents.md 12.134): when the color of a mark is too near to the color
        // behind it (a color difference, CIE76 delta E, below MinMarkerDistance), the mark is made darker or lighter, by the smallest
        // step that gives enough difference (the direction that gives more); when no shade of it does, black or white.
        private const double MinMarkerDistance = 45;

        private static Color Visible(Color mark, Color behind)
        {
            if (ColorDistance(mark, behind) >= MinMarkerDistance) return mark;
            for (int i = 1; i <= 20; i++)
            {
                double t = i * 0.05;
                Color dark = Mix(mark, Color.Black, t), light = Mix(mark, Color.White, t);
                double dd = ColorDistance(dark, behind), dl = ColorDistance(light, behind);
                if (dd >= MinMarkerDistance || dl >= MinMarkerDistance) return dd >= dl ? dark : light;
            }
            return ColorDistance(Color.Black, behind) >= ColorDistance(Color.White, behind) ? Color.Black : Color.White;
        }

        // CIE76 color difference (distance in CIE L*a*b*, D65).
        private static double ColorDistance(Color a, Color b)
        {
            double[] la = Lab(a), lb = Lab(b);
            double dl = la[0] - lb[0], da = la[1] - lb[1], db = la[2] - lb[2];
            return Math.Sqrt(dl * dl + da * da + db * db);
        }

        private static double[] Lab(Color c)
        {
            Func<int, double> lin = v => { double x = v / 255.0; return x <= 0.04045 ? x / 12.92 : Math.Pow((x + 0.055) / 1.055, 2.4); };
            double r = lin(c.R), g = lin(c.G), b = lin(c.B);
            double x1 = (0.4124564 * r + 0.3575761 * g + 0.1804375 * b) / 0.95047;
            double y1 = 0.2126729 * r + 0.7151522 * g + 0.0721750 * b;
            double z1 = (0.0193339 * r + 0.1191920 * g + 0.9503041 * b) / 1.08883;
            Func<double, double> f = t => t > 0.008856 ? Math.Pow(t, 1.0 / 3.0) : 7.787 * t + 16.0 / 116.0;
            double fx = f(x1), fy = f(y1), fz = f(z1);
            return new[] { 116 * fy - 16, 500 * (fx - fy), 200 * (fy - fz) };
        }

        // An open folder and its sheets are enclosed by a dashed line, in the opposite color of the bar
        // background, with the upper right and the lower left corners cut off. The line goes when the folder is closed.
        private void DrawFolderOutlines(Graphics g)
        {
            if (!_outlineVisible) return;
            Color lineColor = OutlineColorAt(Theme.BarBack, _outlineLevel);
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
                using (var pen = new Pen(lineColor, Math.Max(1f, Px(1.2))))
                {
                    pen.DashStyle = DashStyle.Dash;
                    pen.LineJoin = LineJoin.Miter;
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

            // like a sheet tab: selected = white, or the color itself, with a shadow; not selected
            // with a color = a box in that color
            if (selected && DoubleFrame)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                DrawShadow(g, body, radius, 1.0);
                using (GraphicsPath path = RoundedRect(body, radius))
                using (var brush = new SolidBrush(colored ? color : Color.White))
                {
                    g.FillPath(brush, path);
                }
                DrawSelectionFrame(g, body, radius, LineColorFor(color, colored), false); // a folder has no triangles
                if (colored) textColor = TextOn(color);
            }
            else if (selected && Inset)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawInset(g, body, radius, colored, color);
            }
            else if (selected && Stripes)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawStripes(g, body, radius, colored, color);
            }
            else if (selected && Gradient)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawGradientFrame(g, body, colored, color, 1.0);
            }
            else if (selected && Glass)
            {
                if (colored) FillColoredBox(g, left, right, row, radius, color);
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                DrawGlass(g, body, radius, true, colored, color);
                textColor = GlassText(colored);
            }
            else if (selected)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                DrawStandardSelected(g, body, radius, colored, color);
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
            int markShift = 0; // (a selected folder has no marker, so its mark and its name keep their place)
            int cx = left + markShift + Px(TextPadLeft) / 2 + Px(FolderMarkWidth) / 2;
            int cy = row.Top + row.Height / 2 - (selected && DoubleFrame ? Px(ActiveTextLift) : 0);
            int w = Px(FolderMarkWidth), h = Px(FolderMarkHeight);
            Point[] chevron = folder.Expanded
                ? new[] { new Point(cx - h / 2, cy - w / 2), new Point(cx, cy + w / 2), new Point(cx + h / 2, cy - w / 2) }
                : new[] { new Point(cx - w / 2, cy - h / 2), new Point(cx + w / 2, cy), new Point(cx - w / 2, cy + h / 2) };
            using (var pen = new Pen(textColor, Math.Max(1.5f, Px(1.4))))
            {
                pen.LineJoin = LineJoin.Miter;
                g.DrawLines(pen, chevron);
            }

            int textLeft = left + markShift + Px(TextPadLeft) + Px(FolderMarkWidth) + Px(4);
            int roomRight = 0;
            var textRect = new Rectangle(textLeft, row.Top - (selected && DoubleFrame ? Px(ActiveTextLift) : 0) + TextShift(folder.FolderName), Math.Max(1, right - textLeft - Px(TextPadRight) - roomRight - CheckReserve(folder)), row.Height);
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
            if (folder.SyncColor) DrawCheck(g, row, right);
            DrawSeparator(g, row);
        }

        // The check mark at the right end of a folder row. It is dark and has a soft light halo (a few pixels
        // wide, white pens of growing width with a little opacity each, which looks like a blur), so that it
        // can be seen on the dark bar of the dark Office themes and on any folder color alike.
        private void DrawCheck(Graphics g, Rectangle row, int right)
        {
            int w = Px(CheckWidth), h = Px(CheckHeight);
            int x2 = right - Px(CheckPadRight), x1 = x2 - w;
            int top = row.Top + Px(PlainTop), bottom = row.Bottom - Px(BottomInset);
            int cy = (top + bottom) / 2;
            var points = new[]
            {
                new Point(x1, cy), new Point(x1 + (int)Math.Round(w * 0.38), cy + h / 2), new Point(x2, cy - h / 2)
            };
            float markWidth = Math.Max(1.5f, Px(1.6));
            float halo = Math.Max(2f, Px(CheckHalo));
            const int steps = 6;
            for (int i = steps; i >= 1; i--)
            {
                using (var pen = new Pen(Color.FromArgb(46, 255, 255, 255), markWidth + 2 * halo * i / steps))
                {
                    pen.LineJoin = LineJoin.Round;
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    g.DrawLines(pen, points);
                }
            }
            using (var pen = new Pen(TextColor, markWidth))
            {
                pen.LineJoin = LineJoin.Round;
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                g.DrawLines(pen, points);
            }
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
            string text = tabRow.Text;
            bool selected = _selected.Contains(name);
            bool active = selected && name == _active;
            int bgr;
            bool colored = _colors.TryGetValue(name, out bgr);
            Color tabColor = colored ? Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF) : Color.Empty;

            int left = row.Left + Px(InsetLeft) + (tabRow.Indented ? Px(FolderIndent) : 0);
            int right = row.Right - RightInset(tabRow);
            int radius = Px(CornerRadius);
            Color textColor = selected ? TextColor : Theme.InactiveText;
            bool underline = active && _activeMarker == ActiveMarker.Underline;
            bool triangles = active && _activeMarker == ActiveMarker.Triangle;
            bool bars = active && _activeMarker == ActiveMarker.EdgeBars;
            bool badge = active && _activeMarker == ActiveMarker.CheckBadge;

            // A colored tab that is not selected is a filled, rounded box; every other tab is white
            // when selected and has no background otherwise.
            if (selected && DoubleFrame)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                // the active sheet is raised most; the other sheets of a group a little less
                DrawShadow(g, body, radius, active ? 1.0 : 0.6);
                using (GraphicsPath path = RoundedRect(body, radius))
                using (var brush = new SolidBrush(colored ? tabColor : Color.White))
                {
                    g.FillPath(brush, path);
                }
                DrawSelectionFrame(g, body, radius, LineColorFor(tabColor, colored), triangles);
                if (underline) DrawUnderline(g, body, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
                if (colored) textColor = TextOn(tabColor);
            }
            else if (selected && Inset)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawInset(g, body, radius, colored, tabColor);
                if (underline) DrawUnderline(g, body, LineColorFor(tabColor, colored));
                if (triangles) DrawMarkers(g, body, 0, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
            }
            else if (selected && Stripes)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawStripes(g, body, radius, colored, tabColor);
                if (underline) DrawUnderline(g, body, LineColorFor(tabColor, colored));
                if (triangles) DrawMarkers(g, body, 0, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
            }
            else if (selected && Gradient)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawGradientFrame(g, body, colored, tabColor, active ? 1.0 : 0.6);
                if (underline) DrawUnderline(g, body, LineColorFor(tabColor, colored));
                if (triangles) DrawMarkers(g, body, 0, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
            }
            else if (selected && Glass)
            {
                if (colored) FillColoredBox(g, left, right, row, radius, tabColor);
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                // the marks of the active sheet lie under the glass (it covers them as well as the row)
                if (underline) DrawUnderline(g, body, LineColorFor(tabColor, colored));
                if (triangles) DrawMarkers(g, body, 0, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
                if (badge) DrawCheckBadge(g, body, colored, tabColor);
                // the active sheet behind thicker glass than the other sheets of a group
                DrawGlass(g, body, radius, active, colored, tabColor);
                textColor = GlassText(colored);
            }
            else if (selected)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                DrawStandardSelected(g, body, radius, colored, tabColor);
                // Excel's own look: its green underline, whatever the color of the tab
                if (underline) DrawUnderline(g, body, Visible(Underline, colored ? Mix(Color.White, tabColor, 0.45) : Color.White));
                if (triangles) DrawMarkers(g, body, 0, LineColorFor(tabColor, colored));
                if (bars) DrawEdgeBars(g, body, LineColorFor(tabColor, colored));
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

            // the check badge marker (Agents.md 12.131), over whichever look the row has
            if (badge && !(selected && Glass)) DrawCheckBadge(g, new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset)), colored, tabColor);

            // the separator between tabs: Excel's thin line, turned to lie across the stack
            DrawSeparator(g, row);

            // text: left aligned; the text of a selected tab sits a little higher (the lines above and below)
            // an underline needs room below the text, so the text sits higher
            int lift = underline ? Px(StandardTextLift) : DoubleFrame && selected ? Px(ActiveTextLift) : 0;
            int room = triangles ? Px(MarkerRoom) : bars ? Px(BarRoom) : 0; // room for the triangles or the bars at both ends of the active sheet
            int roomLeft = badge ? Px(BadgeRoom) : room; // (the check badge takes the left end; it is the only mark then)
            var textRect = new Rectangle(left + roomLeft + Px(TextPadLeft), row.Top - lift + TextShift(text), Math.Max(1, right - left - roomLeft - room - Px(TextPadLeft) - Px(TextPadRight)), row.Height);
            if (_editor == null || name != _editingSheet)
            {
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                using (var font = selected && (DoubleFrame || Glass || Inset || Stripes || Gradient) ? new Font(TabFont, FontStyle.Bold) : null)
                using (var brush = new SolidBrush(textColor))
                using (var format = new StringFormat(StringFormat.GenericDefault))
                {
                    format.Alignment = StringAlignment.Near;
                    format.LineAlignment = StringAlignment.Center;
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                    g.DrawString(text, font ?? TabFont, brush, textRect, format);
                }
            }
        }

        // The standard type, Excel's own look: white with a thin border; a color fades in from the top
        // (white at the top, about 45% of the color at the bottom).
        private void DrawStandardSelected(Graphics g, Rectangle body, int radius, bool colored, Color color)
        {
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

        // The inset type (Agents.md 12.115): the row looks pressed in. The body is a little darker than the bar
        // (a colored row a little darker than its color), with a shadow that fades inward from its upper and
        // left edges and a thin light line along its bottom. Returns the color of the text on it.
        private Color DrawInset(Graphics g, Rectangle body, int radius, bool colored, Color color)
        {
            bool dark = Theme.BarBack.GetBrightness() < 0.4f;
            Color fill = colored ? Mix(color, Color.Black, 0.12) : Mix(Theme.BarBack, Color.Black, dark ? 0.30 : 0.10);
            using (GraphicsPath path = RoundedRect(body, radius))
            {
                using (var brush = new SolidBrush(fill)) g.FillPath(brush, path);
                Region clip = g.Clip;
                g.SetClip(path, CombineMode.Intersect);
                var mode = g.SmoothingMode;
                g.SmoothingMode = SmoothingMode.None;
                int depth = Math.Max(3, Px(4));
                for (int i = 0; i < depth; i++)
                {
                    using (var brush = new SolidBrush(Color.FromArgb(55 * (depth - i) / depth, 0, 0, 0)))
                    {
                        g.FillRectangle(brush, body.X, body.Y + i, body.Width, 1);
                        g.FillRectangle(brush, body.X + i, body.Y, 1, body.Height);
                    }
                }
                g.SmoothingMode = mode;
                g.Clip = clip;
                using (var pen = new Pen(Color.FromArgb(dark ? 70 : 200, 255, 255, 255), 1f))
                {
                    g.DrawLine(pen, body.X + radius, body.Bottom, body.Right - radius, body.Bottom);
                }
            }
            return TextOn(fill);
        }

        // The stripes look (Agents.md 12.131): a white body (a colored row: its color) with fine diagonal stripes
        // (green on white, white on a color) and an outline in the green (a colored row: a dark shade of its color).
        // Returns the color of the text on it.
        private Color DrawStripes(Graphics g, Rectangle body, int radius, bool colored, Color color)
        {
            Color fill = colored ? color : Color.White;
            Color line = colored ? Mix(color, Color.Black, 0.4) : Underline;
            using (GraphicsPath path = RoundedRect(body, radius))
            {
                using (var brush = new SolidBrush(fill)) g.FillPath(brush, path);
                Region clip = g.Clip;
                g.SetClip(path, CombineMode.Intersect);
                using (var pen = new Pen(colored ? Color.FromArgb(70, 255, 255, 255) : Color.FromArgb(55, Underline.R, Underline.G, Underline.B), Math.Max(2f, Px(2.6))))
                {
                    int step = Math.Max(4, Px(7));
                    for (int x = body.Left - body.Height; x < body.Right; x += step)
                    {
                        g.DrawLine(pen, x, body.Bottom, x + body.Height, body.Top);
                    }
                }
                g.Clip = clip;
                using (var pen = new Pen(line, Math.Max(1.5f, Px(1.4)))) g.DrawPath(pen, path);
            }
            return TextOn(fill);
        }

        // The gradient-frame look (Agents.md 12.131): a pale chip (a tint of the row's color; a row without color: of Excel's green)
        // that is a little raised (half of the shadow of the neumorphic sketch: a light shadow at the lower right, a white glow at
        // the upper left; GDI+ has no blur, so a few translucent shapes of growing size stand for each), inside a ring whose color
        // runs from left to right: a dark shade of the row's color to a light one (Excel's green to a teal for a row without color).
        // Returns the color of the text on it.
        private Color DrawGradientFrame(Graphics g, Rectangle body, bool colored, Color color, double strength)
        {
            bool dark = Theme.BarBack.GetBrightness() < 0.4f;
            int radius = Px(8);
            Color tint = colored ? Mix(color, Color.White, 0.72) : Mix(Underline, Color.White, 0.88);
            Color left = colored ? Mix(color, Color.Black, 0.35) : Underline;
            Color right = colored ? Mix(color, Color.White, 0.15) : Color.FromArgb(0, 176, 160);
            for (int k = 3; k >= 1; k--)
            {
                var r = Rectangle.Inflate(body, Px(0.5 * k), Px(0.5 * k));
                r.Offset(Px(1), Px(1.25));
                using (GraphicsPath path = RoundedRect(r, radius + Px(0.5 * k)))
                using (var brush = new SolidBrush(Color.FromArgb((int)Math.Round((dark ? 40 : 22) * strength), dark ? Color.Black : Color.FromArgb(120, 112, 128))))
                {
                    g.FillPath(brush, path);
                }
            }
            for (int k = 2; k >= 1; k--)
            {
                var r = Rectangle.Inflate(body, Px(0.5 * k), Px(0.5 * k));
                r.Offset(-Px(1), -Px(1));
                using (GraphicsPath path = RoundedRect(r, radius + Px(0.5 * k)))
                using (var brush = new SolidBrush(Color.FromArgb((int)Math.Round((dark ? 10 : 70) * strength), Color.White)))
                {
                    g.FillPath(brush, path);
                }
            }
            using (GraphicsPath path = RoundedRect(body, radius))
            using (var brush = new SolidBrush(tint))
            {
                g.FillPath(brush, path);
            }
            float ring = Math.Max(2f, Px(2.6));
            Rectangle ringRect = Rectangle.Inflate(body, -(int)Math.Ceiling(ring / 2), -(int)Math.Ceiling(ring / 2));
            using (GraphicsPath path = RoundedRect(ringRect, Math.Max(1, radius - (int)Math.Ceiling(ring / 2))))
            using (var brush = new LinearGradientBrush(new Rectangle(body.X, body.Y, Math.Max(2, body.Width), Math.Max(2, body.Height)), left, right, LinearGradientMode.Horizontal))
            using (var pen = new Pen(brush, ring))
            {
                g.DrawPath(pen, path);
            }
            return TextColor;
        }

        // The check badge marker (Agents.md 12.131): a round badge at the left end of the active sheet, with a check mark: a green
        // badge with a white check; on a colored row a white badge with a check in the color of the row.
        private void DrawCheckBadge(Graphics g, Rectangle body, bool colored, Color color)
        {
            int full = Math.Max(6, body.Height - Px(8)); // the first size; the badge keeps its left edge and is 0.75 times as large
            int d = Math.Max(5, (int)Math.Round(full * BadgeScale));
            var badge = new Rectangle(body.Left + (body.Height - full) / 2 + Px(1), body.Top + (body.Height - d) / 2, d, d);
            // a colored row: a white badge with a check in the row's color (a darker or lighter one when it is too near to white);
            // when the row itself is too near to white for the badge to show, a dark badge with a white check
            bool whiteBadge = colored && ColorDistance(Color.White, color) >= MinMarkerDistance;
            Color fill = whiteBadge ? Color.White : colored ? Visible(Underline, color) : Underline;
            Color mark = whiteBadge ? Visible(color, Color.White) : Color.White;
            using (var brush = new SolidBrush(fill)) g.FillEllipse(brush, badge);
            int cx = badge.Left + d / 2, cy = badge.Top + d / 2;
            var check = new[]
            {
                new Point(cx - (int)(d * 0.25), cy), new Point(cx - (int)(d * 0.05), cy + (int)(d * 0.2)), new Point(cx + (int)(d * 0.28), cy - (int)(d * 0.2))
            };
            using (var pen = new Pen(mark, Math.Max(1.5f, Px(1.7))))
            {
                pen.LineJoin = LineJoin.Round;
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                g.DrawLines(pen, check);
            }
        }

        // The colored box of a row that is not selected (the glass of a selected row lies over it).
        private void FillColoredBox(Graphics g, int left, int right, Rectangle row, int radius, Color color)
        {
            var box = new Rectangle(left, row.Top + Px(ColoredTop), right - left, row.Height - Px(ColoredTop) - Px(BottomInset));
            using (GraphicsPath path = RoundedRect(box, radius))
            using (var brush = new SolidBrush(color))
            {
                g.FillPath(brush, path);
            }
        }

        // Text on glass: the glass makes a colored row lighter, so dark text; on the dark bars of the dark
        // Office themes the glass stays dark enough for white text.
        private Color GlassText(bool colored)
        {
            if (colored) return TextColor;
            return Theme.BarBack.GetBrightness() < 0.4f ? Color.White : TextColor;
        }

        // The glass types (Agents.md 12.105): a pane of glass over the row, through which the row (its color,
        // the separator line) shows. Frosted: a fine grain, and edges that glow from inside. Glossy: a reflection
        // on the upper half. GDI+ has no blur: the glass is translucent white, the grain a fixed texture and the
        // glow a few rings of growing inset. The shadow is drawn outside the glass only (it would show through).
        // On the dark bars of the dark Office themes the glass is less white.
        private void DrawGlass(Graphics g, Rectangle body, int radius, bool strong, bool colored, Color color)
        {
            double k = strong ? 1.0 : 0.7;
            double wk = Theme.BarBack.GetBrightness() < 0.4f ? 0.55 : 1.0;
            int rad = radius + Px(1.6);
            using (GraphicsPath path = RoundedRect(body, rad))
            {
                Region clip = g.Clip;
                g.SetClip(path, CombineMode.Exclude);
                DrawShadow(g, body, rad, strong ? 1.0 : 0.6);
                g.Clip = clip;

                if (_selectStyle == SelectStyle.GlossyGlass)
                {
                    using (var brush = new SolidBrush(Color.FromArgb((int)(95 * k * wk), 255, 255, 255))) g.FillPath(brush, path);
                    var top = new Rectangle(body.X, body.Y, body.Width, (int)(body.Height * 0.48));
                    using (GraphicsPath topPath = RoundedRect(top, rad))
                    using (var brush = new LinearGradientBrush(new Rectangle(top.X, top.Y - 1, top.Width, top.Height + 2),
                        Color.FromArgb((int)(200 * k * wk), 255, 255, 255), Color.FromArgb((int)(90 * k * wk), 255, 255, 255), LinearGradientMode.Vertical))
                    {
                        g.FillPath(brush, topPath);
                    }
                    using (var pen = new Pen(Color.FromArgb(240, 255, 255, 255), Math.Max(1f, (float)(1.1 * Dpi / 96.0)))) g.DrawPath(pen, path);
                    return;
                }

                if (_selectStyle == SelectStyle.FrostedGlass2)
                {
                    // frosted glass tinted with the row's color (Excel's green on a row without color), with an edge
                    // in that color and a thin white line inside it
                    Color tint = colored ? color : Underline;
                    using (var brush = new SolidBrush(Color.FromArgb((int)(110 * k * wk), 255, 255, 255))) g.FillPath(brush, path);
                    using (var brush = new SolidBrush(Color.FromArgb((int)(60 * k), tint))) g.FillPath(brush, path);
                    using (var brush = new TextureBrush(GrainTexture((int)(60 * wk), 9))) g.FillPath(brush, path);
                    using (var pen = new Pen(Color.FromArgb(170, tint), Math.Max(1f, (float)(1.1 * Dpi / 96.0)))) g.DrawPath(pen, path);
                    int inset = Math.Max(1, (int)Math.Round(1.3 * Dpi / 96.0));
                    var inner = Rectangle.Inflate(body, -inset, -inset);
                    using (GraphicsPath innerPath = RoundedRect(inner, Math.Max(1, rad - inset)))
                    using (var pen = new Pen(Color.FromArgb(200, 255, 255, 255), 1f))
                    {
                        g.DrawPath(pen, innerPath);
                    }
                    return;
                }

                using (var brush = new SolidBrush(Color.FromArgb((int)(105 * k * wk), 255, 255, 255))) g.FillPath(brush, path);
                using (var brush = new TextureBrush(GrainTexture((int)(55 * wk), 9)))
                {
                    g.FillPath(brush, path);
                }
                float step = (float)(1.0 * Dpi / 96.0);
                for (int i = 0; i < 6; i++)
                {
                    int inset = (int)Math.Round(1 + i * step);
                    var ring = Rectangle.Inflate(body, -inset, -inset);
                    if (ring.Height < 4) break;
                    using (GraphicsPath ringPath = RoundedRect(ring, Math.Max(1, rad - inset)))
                    using (var pen = new Pen(Color.FromArgb((int)((150 - i * 24) * wk), 255, 255, 255), Math.Max(1f, (float)(1.1 * Dpi / 96.0))))
                    {
                        g.DrawPath(pen, ringPath);
                    }
                }
                using (var pen = new Pen(Color.FromArgb((int)(70 * k), 0, 0, 0), 1f)) g.DrawPath(pen, path);
            }
        }

        // The grain of the frosted glass: white and a few dark dots, always the same pattern (a fixed seed), so
        // that it does not flicker when the list is drawn again.
        // One per strength (the two frosted types and the light and dark Office themes differ).
        private static readonly Dictionary<int, Bitmap> _grains = new Dictionary<int, Bitmap>();

        private static Bitmap GrainTexture(int whiteAlpha, int darkAlpha)
        {
            int key = whiteAlpha * 1000 + darkAlpha;
            Bitmap cached;
            if (_grains.TryGetValue(key, out cached)) return cached;
            var bmp = new Bitmap(97, 97, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            var rnd = new Random(20261022);
            Color white = Color.FromArgb(whiteAlpha, 255, 255, 255), dark = Color.FromArgb(darkAlpha, 0, 0, 0);
            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    double v = rnd.NextDouble();
                    bmp.SetPixel(x, y, v < 0.30 ? white : v > 0.93 ? dark : Color.Transparent);
                }
            }
            _grains[key] = bmp;
            return bmp;
        }

        // A selected row looks raised: a soft shadow below and at the right of its body, made of a few
        // translucent black rounded rectangles of growing size (GDI+ has no blur), drawn before the body.
        private void DrawShadow(Graphics g, Rectangle body, int radius, double strength)
        {
            bool dark = Theme.BarBack.GetBrightness() < 0.4f;
            int alpha = (int)Math.Round((dark ? 22 : 11) * strength);
            for (int k = 4; k >= 1; k--)
            {
                var r = new Rectangle(body.X - Px(k - 1), body.Y + Px(1 + k), body.Width + Px(1.2 * k), body.Height);
                using (GraphicsPath path = RoundedRect(r, radius + Px(k)))
                using (var brush = new SolidBrush(Color.FromArgb(alpha, 0, 0, 0)))
                {
                    g.FillPath(brush, path);
                }
            }
        }

        // A selected row has a double frame (a dark outer one and a white inner one) and, for the active sheet only, a small triangle at each end
        // that points to the middle, in the opposite color of the row; the triangles have a white edge and a soft
        // white halo of a few pixels (white shapes of growing width with a little opacity each: GDI+ has no blur),
        // so that they can be seen on dark colors and on the dark bar of the dark Office themes as well.
        private void DrawSelectionFrame(Graphics g, Rectangle body, int radius, Color marker, bool markers)
        {
            bool dark = Theme.BarBack.GetBrightness() < 0.4f;
            float outerWidth = Math.Max(2f, Px(1.6));
            float innerWidth = Math.Max(1.5f, Px(1.3));
            var outerRect = Rectangle.Inflate(body, -(int)Math.Ceiling(outerWidth / 2), -(int)Math.Ceiling(outerWidth / 2));
            using (GraphicsPath path = RoundedRect(outerRect, radius))
            using (var pen = new Pen(dark ? Color.FromArgb(16, 16, 16) : Color.FromArgb(43, 43, 43), outerWidth))
            {
                g.DrawPath(pen, path);
            }
            int inset = Px(2.6);
            var innerRect = Rectangle.Inflate(body, -inset, -inset);
            if (innerRect.Width > 4 && innerRect.Height > 4)
            {
                using (GraphicsPath path = RoundedRect(innerRect, Math.Max(1, radius - Px(1.5))))
                using (var pen = new Pen(Color.White, innerWidth))
                {
                    g.DrawPath(pen, path);
                }
            }
            if (markers) DrawMarkers(g, body, 0, marker);
        }

        // The bars at both ends of the active sheet (Agents.md 12.116): thick rounded bars in the color of the other
        // marks, with a white edge so that they can be seen on any color of the row and on the dark Office themes.
        private void DrawEdgeBars(Graphics g, Rectangle body, Color color)
        {
            int width = Px(BarWidth), inset = Px(BarInset);
            foreach (int x in new[] { body.Left + inset, body.Right - inset - width }) DrawEdgeBar(g, body, x, color);
        }

        // One bar of DrawEdgeBars, with its left edge at x.
        private void DrawEdgeBar(Graphics g, Rectangle body, int x, Color color)
        {
            int width = Px(BarWidth), margin = Px(BarMargin);
            float edge = Math.Max(1f, Px(1.2));
            int height = Math.Max(2, (int)Math.Round((body.Height - 2 * margin) * BarScale));
            var bar = new Rectangle(x, body.Top + (body.Height - height) / 2, width, height);
            using (GraphicsPath path = RoundedRect(bar, Math.Max(1, Px(2.5))))
            {
                using (var pen = new Pen(Color.White, edge * 2)) g.DrawPath(pen, path);
                using (var brush = new SolidBrush(color)) g.FillPath(brush, path);
            }
        }

        // The underline of the active sheet (Excel's measures: 8 DIP in from both ends, 1.6 DIP thick, 3.2 DIP above
        // the bottom of the body).
        private void DrawUnderline(Graphics g, Rectangle body, Color color)
        {
            int thick = Math.Max(1, Px(UnderlineThickness));
            int ux = body.Left + Px(UnderlineInset);
            int uw = body.Width - 2 * Px(UnderlineInset);
            int uy = body.Bottom - 1 - Px(UnderlineGap) - thick;
            using (var brush = new SolidBrush(color)) g.FillRectangle(brush, ux, uy, Math.Max(1, uw), thick);
        }

        // The triangles at both ends of the active sheet (Agents.md 12.109: every type has them), raise pixels above
        // the middle of the body.
        private void DrawMarkers(Graphics g, Rectangle body, int raise, Color marker)
        {
            int cy = body.Top + body.Height / 2 - raise;
            DrawMarker(g, body.Left + Px(MarkerInset), cy, 1, marker);
            DrawMarker(g, body.Right - Px(MarkerInset), cy, -1, marker);
        }

        // One triangle pointing right (dir 1) or left (dir -1), centered at (cx, cy).
        private void DrawMarker(Graphics g, int cx, int cy, int dir, Color color)
        {
            int halfH = Px(MarkerHalfHeight), back = Px(MarkerBack), front = Px(MarkerFront);
            var points = new[]
            {
                new Point(cx - dir * back, cy - halfH), new Point(cx + dir * front, cy), new Point(cx - dir * back, cy + halfH)
            };
            const int steps = 6;
            for (int i = steps; i >= 1; i--)
            {
                using (var pen = new Pen(Color.FromArgb(34, 255, 255, 255), Math.Max(1f, Px(1.2)) + i * Math.Max(1f, Px(1.3)) / 2f)) // half of the first halo (12.109)
                {
                    pen.LineJoin = LineJoin.Round;
                    g.DrawPolygon(pen, points);
                }
            }
            using (var brush = new SolidBrush(color)) g.FillPolygon(brush, points);
            using (var pen = new Pen(Color.White, Math.Max(1f, Px(1.2))))
            {
                pen.LineJoin = LineJoin.Round;
                g.DrawPolygon(pen, points);
            }
        }

        // The text color on a sheet or folder color: dark on light colors, white on dark ones.
        private static Color TextOn(Color c)
        {
            double luma = (c.R * 299 + c.G * 587 + c.B * 114) / 1000.0;
            return luma >= 150 ? TextColor : Color.White;
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

        // A click that chooses another sheet must not give the keyboard focus to the pane: while the pane has
        // it, Excel does not pass the keys on to the sheet, and typing right after choosing a sheet (which
        // works after a click on Excel's own tabs) would be lost. So such a click leaves the focus where it
        // is. A click that chooses nothing new (the active tab, a folder) still focuses the pane, for its
        // keys (F2, Del, up/down).
        // Whether Excel is editing a cell: a click on the tabs must then leave the focus in the edit window too.
        public Func<bool> IsExcelEditing;

        private bool _noFocusClick;
        private bool _hadFocusAtClick;

        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_LBUTTONDBLCLK = 0x0203;

        protected override void WndProc(ref Message m)
        {
            if (IsSample) { base.WndProc(ref m); return; }
            if ((m.Msg == WM_LBUTTONDOWN || m.Msg == WM_LBUTTONDBLCLK) && _editor == null)
            {
                long l = m.LParam.ToInt64();
                var point = new Point((short)(l & 0xFFFF), (short)((l >> 16) & 0xFFFF));
                if (ClickChangesSheet(point) || (IsExcelEditing != null && IsExcelEditing()))
                {
                    _hadFocusAtClick = ContainsFocus;
                    _noFocusClick = true;
                    SetStyle(ControlStyles.Selectable, false);
                    try { base.WndProc(ref m); }
                    finally
                    {
                        SetStyle(ControlStyles.Selectable, true);
                        _noFocusClick = false;
                    }
                    return;
                }
            }
            base.WndProc(ref m);
        }

        // Whether a left click here would choose other sheets than the ones Excel has selected now.
        private bool ClickChangesSheet(Point p)
        {
            int index = IndexFromPoint(p);
            if (index < 0 || index >= _rows.Count || _rows[index].IsFolder) return false;
            if (ModifierKeys != Keys.None) return true; // Ctrl / Shift: the group changes
            string name = _rows[index].Sheet;
            return !(name == _active && _selected.Count == 1 && _selected.Contains(name));
        }

        // Left button down: update the selection on screen, remember where a drag could start.
        // Excel is told when the button is released, unless a drag starts first.
        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (IsSample) { if (SampleClicked != null) SampleClicked(); return; }
            if (!_noFocusClick) Focus();
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
            if (IsSample) return;
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
            if (GridFocusRequested != null) GridFocusRequested(_hadFocusAtClick);
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
            if (IsSample) return;
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
            string text = _rows[index].IsFolder ? _rows[index].FolderName : _rows[index].Text;
            if (!IsTextCut(text, _rows[index])) return;
            Rectangle row = RowRect(index);
            _tip.Show(text, this, row.Left + Px(InsetLeft + TextPadLeft), row.Bottom, 4000);
        }

        // Whether the name does not fit the tab (it is then drawn with a trailing "…").
        private bool IsTextCut(string name, StripRow row)
        {
            int available = ContentWidth - Px(InsetLeft + TextPadLeft + TextPadRight) - RightInset(row)
                - (row.Indented ? Px(FolderIndent) : 0) - (row.IsFolder ? Px(FolderMarkWidth + 4) : 0) - CheckReserve(row);
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
            if (IsSample) return;
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
            if (IsSample) return;
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
            BeginEdit(index, _rows[index].FolderName, folderId, null);
        }

        // In-place name editing: Enter or leaving the box commits, Esc cancels. The length limit
        // is Excel's 31 characters; the other rules are checked by PaneManager when it renames.
        private void BeginRename(int index)
        {
            if (index < 0 || index >= _rows.Count || _rows[index].IsFolder || StructureProtected) return;
            BeginEdit(index, _rows[index].Text, null, _rows[index].Sheet);
        }

        private void BeginEdit(int index, string text, string folderId, string sheet)
        {
            EndRename(false);
            Rectangle r = RowRect(index);
            _editingName = text;
            _editingFolderId = folderId;
            _editingSheet = sheet;
            int left = r.Left + Px(InsetLeft) + (_rows[index].Indented ? Px(FolderIndent) : 0);
            int right = r.Right - RightInset(_rows[index]) - CheckReserve(_rows[index]);
            int textLeft = folderId != null ? left + Px(TextPadLeft) + Px(FolderMarkWidth) + Px(4) - 2 : left + Px(TextPadLeft) - 2;
            _editor = new TextBox
            {
                Text = text,
                MaxLength = _rows[index].EditLimit > 0 ? _rows[index].EditLimit : 31,
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
            string sheetName = _editingSheet;
            _editingFolderId = null;
            _editingSheet = null;
            string newName = editor.Text;
            Controls.Remove(editor);
            editor.Dispose();
            Invalidate();
            if (!commit || newName == oldName) return;
            if (folderId != null) { if (FolderRenameRequested != null) FolderRenameRequested(folderId, newName); }
            else if (RenameRequested != null) RenameRequested(sheetName ?? oldName, newName);
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
