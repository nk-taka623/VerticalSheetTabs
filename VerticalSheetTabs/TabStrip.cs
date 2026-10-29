using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // The vertical stack of sheet tabs, drawn like Excel's own horizontal tabs (measured on a real Excel).
    // Each tab is one row of the height of Excel's tab bar. The control owns the
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
        // Folder: its sheet rows, also while the folder is closed (the search shows them without opening the folder).
        public List<StripRow> Children = new List<StripRow>();

        public bool IsFolder { get { return FolderId != null; } }

        public string Text { get { return Display ?? Sheet; } }

        public string Signature()
        {
            return IsFolder ? "F|" + FolderId + "|" + FolderName + "|" + Expanded + "|" + ColorBgr + "|" + SyncColor + "|" + EditLimit + "|" + string.Join("|", Members)
                : "S|" + Sheet + "|" + Indented + "|" + Display + "|" + EditLimit;
        }
    }

    // What a sheet holds that a badge tells of: the number of cells with an error, and of notes and comments.
    public class BadgeInfo
    {
        public int Errors;
        public int Memos;
        public bool Protected;
        public int Progress; // 0 = none, 1 = not started, 2 = in progress, 3 = done
    }

    // A line of the menu of the sheets seen last.
    public class HistoryItem
    {
        public string Key;
        public string Name;
        public bool Current;
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
        private const double BarWidth = 6, BarInset = 9, BarMargin = 4, BarRoom = 12, BarScale = 0.75; // 0.75 times the first size (8 wide); moved in by one bar width
        // the check badge marker at the left of the name of the active sheet: the room it takes before the text
        // The badges at the right end of a row: each takes a step of 18 DIP, the first one stands 6 DIP from the end.
        private const double RowBadgeStep = 18, RowBadgeRadius = 7.5, RowBadgeEdge = 6;
        private const double BadgeRoom = 12, BadgeScale = 0.75; // the badge is 0.75 times the first size
        private const double SeparatorLength = 19;
        private const double TextPadLeft = 12, TextPadRight = 8;
        // Folders: sheets inside a folder are indented; the folder row has an open/closed mark and,
        // with a color, a bar at its left edge.
        private const double OutlineRightInset = 1.5, OutlineCut = 10;
        // The mark of a folder that keeps its sheets in its color (a tipped paint can with a blue drop): size, space to the right end, gap to the text.
        private const double CheckWidth = 16, CheckHeight = 14, CanRadius = 7.6, CheckPadRight = 8, CheckGap = 4, CheckHalo = 3;
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
        // The paint can mark of a folder: a click asks for the palette of the folder color (the folder, and the point
        // on the screen below the mark), a double click for the color dialog of "その他の色".
        public event Action<string, Point> FolderColorMarkClicked;
        public event Action<string> FolderColorMarkDoubleClicked;
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
            CreateSearchBox();
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
            // While a mouse button is down the user is changing the selection; do not fight them. A sample of the settings form takes no
            // clicks for that, and it is set up while the button that chose a tab of the form is still down.
            bool userBusy = !IsSample && Control.MouseButtons != MouseButtons.None;
            bool sameSelection = userBusy || (_active == activeName && _selected.SetEquals(selected));
            if (sameNames && sameRows && sameSelection && sameColors) return;

            // a sheet or a folder added or deleted: the row appears or shrinks away (looked at while the old state is still here)
            RowAnimPending pending = sameNames && sameRows ? null : PrepareRowAnim(names, newRows);

            if (!sameColors)
            {
                _colors.Clear();
                foreach (var kv in colors) _colors[kv.Key] = kv.Value;
            }
            if (!sameNames || !sameRows)
            {
                EndRename(false);
                _names = new List<string>(names);
                _baseRows = newRows;
                _folderCount = newRows.Count(r => r.IsFolder);
                _rowSignature = newSignature;
                ApplyFilter();
                LayoutSearch();
            }
            FinishRowAnim(pending);
            if (!userBusy)
            {
                bool activeChanged = _active != activeName;
                _active = activeName;
                _selected = new HashSet<string>(selected);
                if (activeChanged) EnsureVisible(RowIndexOf(_active));
                PruneFolders();
            }
            DropRowAnimPictures(); // (a row that is coming in looks as the rows look now)
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
        private bool NeedScroll { get { return _rows.Count * RowPx > AreaHeight; } }

        // With one or more folders, a band at the top holds the "open all" and "close all" buttons; the rows start below it.
        private const double HeaderHeight = 24, HeaderGap = 1.6, HeaderRadius = 3.2, HeaderTextShift = 1.5;
        private int _folderCount;
        private int _headerDown, _headerHot; // 0 none, 1 open all, 2 close all
        private bool ShowAllButtons { get { return !IsSample && _folderCount >= 1; } }
        private int BandPx { get { return ShowAllButtons ? Px(HeaderHeight) : 0; } }
        // The rows start below the tool row (back, forward, search) and the band.
        private int HeaderPx { get { return ToolPx + BandPx; } }
        private int AreaHeight { get { return Math.Max(0, ViewHeight - HeaderPx); } }

        public event Action<bool> AllFoldersRequested;

        // The two buttons fill the band with a gap of about 2 px between them and to the edges of the pane.
        private Rectangle AllButtonRect(bool open)
        {
            int gap = Math.Max(1, Px(HeaderGap));
            int w = Math.Max(1, (ClientSize.Width - 3 * gap) / 2);
            return new Rectangle(open ? gap : 2 * gap + w, ToolPx + gap, w, Math.Max(1, BandPx - 2 * gap));
        }

        private int AllButtonAt(Point p)
        {
            if (!ShowAllButtons) return 0;
            return AllButtonRect(true).Contains(p) ? 1 : AllButtonRect(false).Contains(p) ? 2 : 0;
        }


        // ---- the tool row: back, forward, search box, folder switch ----

        private const double ToolHeight = 38, ToolGap = 2.4, ToolButton = 22, ToolBoxHeight = 24, ToolRadius = 4, ToolSmallGap = 1.6;
        private List<StripRow> _baseRows = new List<StripRow>(); // the rows as the workbook has them
        private TextBox _search;
        private string _query = "";
        private bool _folderSearch;                 // the folder switch: search the folder names instead of the sheet names
        private int _searchSel = -1;                // the row Up/Down chose among the results
        private int _toolHot, _toolDown;            // 1 back, 2 forward, 3 clear, 4 folder switch
        private bool _canBack, _canForward;
        private System.Windows.Forms.Timer _longPress;
        private bool _longPressShown;

        public event Action BackRequested;
        public event Action ForwardRequested;
        // The search is over (Esc, or a sheet was chosen with Enter): the keyboard focus goes back to the sheet.
        public event Action GridFocusForced;
        // Asks for the sheets seen last (the menu of the back / forward buttons); the list is made by the pane.
        public Func<List<HistoryItem>> HistoryProvider;
        public event Action<string> HistoryJumpRequested;
        public event Action HistoryClearRequested;

        private int ToolPx { get { return IsSample ? 0 : Px(ToolHeight); } }
        private bool Filtering { get { return !IsSample && _query.Length > 0; } }

        public void SetNavigation(bool canBack, bool canForward)
        {
            if (_canBack == canBack && _canForward == canForward) return;
            _canBack = canBack;
            _canForward = canForward;
            Invalidate();
        }

        private Rectangle ToolBackRect()
        {
            int s = Px(ToolButton);
            return new Rectangle(Px(ToolGap), (ToolPx - s) / 2, s, s);
        }

        private Rectangle ToolForwardRect()
        {
            Rectangle b = ToolBackRect();
            return new Rectangle(b.Right + Px(ToolSmallGap), b.Top, b.Width, b.Height);
        }

        private Rectangle ToolFolderRect()
        {
            int s = Px(ToolButton);
            return new Rectangle(ClientSize.Width - Px(ToolGap) - s, (ToolPx - s) / 2, s, s);
        }

        private Rectangle ToolBoxRect()
        {
            int left = ToolForwardRect().Right + Px(ToolGap);
            int right = ToolFolderRect().Left - Px(ToolGap);
            int h = Px(ToolBoxHeight);
            return new Rectangle(left, (ToolPx - h) / 2, Math.Max(Px(36), right - left), h);
        }

        private Rectangle ToolClearRect()
        {
            Rectangle box = ToolBoxRect();
            int s = Px(16);
            return new Rectangle(box.Right - s - Px(5), box.Top + (box.Height - s) / 2, s, s);
        }

        // The area of the text in the box, between the magnifier and the clear button.
        private Rectangle ToolTextRect()
        {
            Rectangle box = ToolBoxRect();
            int left = box.Left + Px(26);
            int right = box.Right - Px(_query.Length > 0 ? 26 : 6);
            return new Rectangle(left, box.Top + Px(3), Math.Max(Px(10), right - left), box.Height - Px(6));
        }

        private int ToolAt(Point p)
        {
            if (IsSample || p.Y >= ToolPx) return 0;
            if (ToolBackRect().Contains(p)) return 1;
            if (ToolForwardRect().Contains(p)) return 2;
            if (_query.Length > 0 && ToolClearRect().Contains(p)) return 3;
            if (ToolFolderRect().Contains(p)) return 4;
            return ToolBoxRect().Contains(p) ? 5 : 0;
        }

        private void CreateSearchBox()
        {
            _search = new TextBox { BorderStyle = BorderStyle.None, Visible = false, TabStop = false, MaxLength = 64, AutoSize = false };
            _search.TextChanged += (s, e) => OnQueryChanged();
            _search.GotFocus += (s, e) => Invalidate();
            _search.LostFocus += (s, e) => { UpdateSearchVisible(); Invalidate(); };
            _search.KeyDown += SearchKeyDown;
            Controls.Add(_search);
        }

        // The text box exists only while the box has the focus or a text (otherwise the box shows its hint).
        private void UpdateSearchVisible()
        {
            if (_search == null) return;
            bool show = !IsSample && (_search.Focused || _search.Text.Length > 0);
            if (_search.Visible != show) _search.Visible = show;
        }

        private void LayoutSearch()
        {
            if (_search == null) return;
            try
            {
                Rectangle t = ToolTextRect();
                _search.Font = TabFont;
                _search.BackColor = SearchFill();
                _search.ForeColor = SearchText();
                _search.Height = Math.Min(t.Height, _search.PreferredHeight);
                _search.SetBounds(t.Left, t.Top + (t.Height - _search.Height) / 2, t.Width, _search.Height);
                UpdateSearchVisible();
            }
            catch (Exception)
            {
                // The sizes are still settling during the first layout; the next resize fixes it.
            }
        }

        private Color SearchFill() { return Theme.PlusHoverBack; }
        private Color SearchText() { return Theme == TabTheme.Light ? TextColor : Color.FromArgb(235, 235, 235); }

        // Ctrl+F, a click in the box, or typing in the pane.
        public void FocusSearch(string first = null)
        {
            if (IsSample || _search == null) return;
            _search.Visible = true;
            LayoutSearch();
            _search.Focus();
            // (a click that activates the pane may give the focus back to the pane itself afterwards)
            BeginInvoke(new Action(() =>
            {
                if (_search == null || _search.IsDisposed || _search.Focused) return;
                _search.Visible = true;
                _search.Focus();
            }));
            if (first != null)
            {
                _search.Text = first;
                _search.SelectionStart = _search.Text.Length;
            }
            else _search.SelectAll();
        }

        private void OnQueryChanged()
        {
            string q = _search.Text.Trim();
            _query = q;
            ApplyFilter();
            LayoutSearch();
            UpdateSearchVisible();
            Invalidate();
        }

        private void ClearSearch()
        {
            if (_search != null) _search.Text = "";
        }

        // Up / Down choose a result, Enter goes to it, Esc ends the search.
        private void SearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.Handled = e.SuppressKeyPress = true;
                if (_search.Text.Length > 0) ClearSearch();
                else if (GridFocusForced != null) GridFocusForced();
                return;
            }
            if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                e.Handled = e.SuppressKeyPress = true;
                MoveSearchSelection(e.KeyCode == Keys.Down ? 1 : -1);
                return;
            }
            if (e.KeyCode == Keys.Enter)
            {
                e.Handled = e.SuppressKeyPress = true;
                if (_searchSel < 0 || _searchSel >= _rows.Count) return;
                StripRow row = _rows[_searchSel];
                if (row.IsFolder)
                {
                    if (FolderToggleRequested != null) FolderToggleRequested(row.FolderId);
                    return;
                }
                _selFolders.Clear();
                _selected = new HashSet<string> { row.Sheet };
                _active = row.Sheet;
                ClearSearch();
                Raise();
                if (GridFocusForced != null) GridFocusForced();
                return;
            }
            if (ModifierKeys == Keys.Control && e.KeyCode == Keys.S && ExcelCommandRequested != null)
            {
                e.Handled = e.SuppressKeyPress = true;
                ExcelCommandRequested("FileSave");
            }
        }

        private bool IsResultRow(StripRow r)
        {
            return _folderSearch ? r.IsFolder : !r.IsFolder;
        }

        private void MoveSearchSelection(int step)
        {
            if (_rows.Count == 0) return;
            int i = _searchSel;
            do { i += step; } while (i >= 0 && i < _rows.Count && !IsResultRow(_rows[i]));
            if (i < 0 || i >= _rows.Count) return;
            _searchSel = i;
            EnsureVisible(i);
            Invalidate();
        }

        // The number of results (sheets, or folders with the folder switch on).
        private int ResultCount { get { return _rows.Count(IsResultRow); } }

        // ---- the search itself ----

        // Case and width are not told apart: letters are folded to lower case, full-width ASCII to half-width,
        // half-width katakana (with their voiced marks) to full-width. The maps give, for each character of the
        // folded text, where it came from in the original (so that the match can be marked in the name).
        private static string Fold(string text, List<int> starts, List<int> ends)
        {
            const string kana = "ヲァィゥェォャュョッーアイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワン";
            var sb = new System.Text.StringBuilder();
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                int len = 1;
                char r;
                if (c >= '\uFF01' && c <= '\uFF5E') r = char.ToLowerInvariant((char)(c - 0xFEE0));
                else if (c == '\u3000') r = ' ';
                else if (c >= '\uFF66' && c <= '\uFF9D')
                {
                    r = kana[c - 0xFF66];
                    if (i + 1 < text.Length && text[i + 1] == '\uFF9E')
                    {
                        if ("カキクケコサシスセソタチツテトハヒフヘホ".IndexOf(r) >= 0) { r = (char)(r + 1); len = 2; }
                        else if (r == 'ウ') { r = 'ヴ'; len = 2; }
                    }
                    else if (i + 1 < text.Length && text[i + 1] == '\uFF9F' && "ハヒフヘホ".IndexOf(r) >= 0) { r = (char)(r + 2); len = 2; }
                }
                else r = char.ToLowerInvariant(c);
                sb.Append(r);
                if (starts != null) { starts.Add(i); ends.Add(i + len); }
                i += len - 1;
            }
            return sb.ToString();
        }

        private static bool FindMatch(string text, string query, out int start, out int length)
        {
            start = length = 0;
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(query)) return false;
            var starts = new List<int>();
            var ends = new List<int>();
            string folded = Fold(text, starts, ends);
            string q = Fold(query, null, null);
            if (q.Length == 0) return false;
            int at = folded.IndexOf(q, StringComparison.Ordinal);
            if (at < 0) return false;
            start = starts[at];
            length = ends[at + q.Length - 1] - start;
            return true;
        }

        private static bool Matches(string text, string query)
        {
            int s, l;
            return FindMatch(text, query, out s, out l);
        }

        // The rows that are shown: all of them, or the results of the search. The sheets (with the folders around
        // them, opened) whose names contain the text; with the folder switch on, the folders whose names do (each
        // as it is, closed or open: opening it shows all its sheets).
        private void ApplyFilter()
        {
            if (!Filtering)
            {
                MarkGroups(_baseRows); // (a search may have marked them for its own list)
                _rows = _baseRows;
                _searchSel = -1;
            }
            else
            {
                var result = new List<StripRow>();
                foreach (StripRow row in _baseRows)
                {
                    if (row.Indented) continue; // the sheets of a folder are looked at through the folder
                    if (row.IsFolder)
                    {
                        List<StripRow> kids = row.Children;
                        if (_folderSearch)
                        {
                            if (!Matches(row.FolderName, _query)) continue;
                            result.Add(row);
                            if (row.Expanded) result.AddRange(kids);
                        }
                        else
                        {
                            List<StripRow> hits = kids.Where(k => Matches(k.Text, _query)).ToList();
                            if (hits.Count == 0) continue;
                            result.Add(new StripRow
                            {
                                FolderId = row.FolderId,
                                FolderName = row.FolderName,
                                Expanded = true,
                                ColorBgr = row.ColorBgr,
                                SyncColor = row.SyncColor,
                                EditLimit = row.EditLimit,
                                Members = row.Members
                            });
                            result.AddRange(hits);
                        }
                    }
                    else if (!_folderSearch && Matches(row.Text, _query)) result.Add(row);
                }
                MarkGroups(result);
                _rows = result;
                int first = result.FindIndex(IsResultRow);
                if (_searchSel < 0 || _searchSel >= result.Count || !IsResultRow(result[_searchSel])) _searchSel = first;
            }
            UpdateScrollBar();
            if (_scrollY > 0 && !_scroll.Visible) _scrollY = 0;
            Invalidate();
        }

        // The mark of the row Up / Down have chosen.
        private void DrawSearchChoice(Graphics g)
        {
            if (!Filtering || _searchSel < 0 || _searchSel >= _rows.Count) return;
            Rectangle row = RowRect(_searchSel);
            var area = new Rectangle(row.Left + Px(1), row.Top + Px(PlainTop) - 1, row.Width - Px(2), row.Height - Px(PlainTop) - Px(BottomInset) + 2);
            using (GraphicsPath path = RoundedRect(area, Px(7)))
            using (var pen = new Pen(Underline, Math.Max(1.5f, Px(1.6))))
                g.DrawPath(pen, path);
        }

        // Marks the part of a name the search found.
        private void DrawMatch(Graphics g, string text, Font font, Rectangle textRect, StringFormat format, bool folder)
        {
            if (!Filtering || folder != _folderSearch) return;
            int s, l;
            if (!FindMatch(text, _query, out s, out l)) return;
            try
            {
                using (var f = (StringFormat)format.Clone())
                {
                    f.SetMeasurableCharacterRanges(new[] { new CharacterRange(s, l) });
                    Region[] regions = g.MeasureCharacterRanges(text, font, textRect, f);
                    RectangleF b = regions[0].GetBounds(g);
                    regions[0].Dispose();
                    if (b.Width < 1) return;
                    var box = Rectangle.Round(new RectangleF(b.Left - 1, b.Top + Px(3), b.Width + 2, Math.Max(1, b.Height - Px(5))));
                    box.Intersect(textRect);
                    using (GraphicsPath path = RoundedRect(box, Px(3)))
                    using (var brush = new SolidBrush(Color.FromArgb(150, 255, 214, 0)))
                        g.FillPath(brush, path);
                }
            }
            catch (Exception)
            {
                // A mark that cannot be measured is left out.
            }
        }

        // ---- drawing of the tool row ----

        private void DrawToolRow(Graphics g)
        {
            DrawNavButton(g, ToolBackRect(), true, _canBack, _toolHot == 1, _toolDown == 1);
            DrawNavButton(g, ToolForwardRect(), false, _canForward, _toolHot == 2, _toolDown == 2);

            Rectangle box = ToolBoxRect();
            bool focus = _search != null && _search.Focused;
            using (GraphicsPath path = RoundedRect(box, Px(ToolRadius + 1)))
            {
                using (var brush = new SolidBrush(SearchFill())) g.FillPath(brush, path);
                int count = Filtering ? ResultCount : -1;
                using (var pen = new Pen(Filtering && count == 0 ? Color.FromArgb(200, 40, 60) : focus ? Underline : Theme.Separator, focus || (Filtering && count == 0) ? Math.Max(1.5f, Px(1.5)) : 1f))
                    g.DrawPath(pen, path);
            }
            // the magnifier
            int mx = box.Left + Px(12), my = box.Top + box.Height / 2 - Px(1);
            using (var pen = new Pen(Theme.PlusColor, Math.Max(1.4f, Px(1.4))))
            {
                int r = Px(4.5);
                g.DrawEllipse(pen, mx - r, my - r, 2 * r, 2 * r);
                g.DrawLine(pen, mx + r - Px(1), my + r - Px(1), mx + r + Px(3), my + r + Px(3));
            }
            if (_search == null || !_search.Visible)
            {
                // the hint
                using (var font = new Font(TabFont.FontFamily, TabFont.Size * 0.8f, FontStyle.Regular, TabFont.Unit))
                using (var brush = new SolidBrush(Theme.InactiveText))
                using (var format = new StringFormat { LineAlignment = StringAlignment.Center, Trimming = StringTrimming.EllipsisCharacter, FormatFlags = StringFormatFlags.NoWrap })
                {
                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    Rectangle t = ToolTextRect();
                    g.DrawString("検索", font, brush, new Rectangle(t.Left, t.Top, t.Width, t.Height), format);
                }
            }
            if (Filtering)
            {
                // the clear button, and how many were found
                Rectangle c = ToolClearRect();
                using (var brush = new SolidBrush(_toolHot == 3 ? Theme.PlusColor : Theme.Separator)) g.FillEllipse(brush, c);
                using (var pen = new Pen(Theme.PlusHoverBack, Math.Max(1.4f, Px(1.4))))
                {
                    int q = Px(3.2);
                    g.DrawLine(pen, c.Left + c.Width / 2 - q, c.Top + c.Height / 2 - q, c.Left + c.Width / 2 + q, c.Top + c.Height / 2 + q);
                    g.DrawLine(pen, c.Left + c.Width / 2 - q, c.Top + c.Height / 2 + q, c.Left + c.Width / 2 + q, c.Top + c.Height / 2 - q);
                }
                int n = ResultCount;
                string label = n + "件";
                using (var font = new Font(TabFont.FontFamily, TabFont.Size * 0.72f, FontStyle.Bold, TabFont.Unit))
                {
                    SizeF size = g.MeasureString(label, font);
                    Rectangle t = ToolTextRect();
                    // only when there is room next to the typed text
                    int textWidth = _search != null ? TextRenderer.MeasureText(_search.Text, TabFont, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding).Width : 0;
                    if (t.Width > textWidth + size.Width + Px(14))
                    {
                        var badge = new Rectangle((int)(c.Left - size.Width - Px(14)), box.Top + Px(4), (int)size.Width + Px(10), box.Height - Px(8));
                        using (GraphicsPath path = RoundedRect(badge, badge.Height / 2))
                        using (var brush = new SolidBrush(n == 0 ? Color.FromArgb(250, 225, 228) : Color.FromArgb(220, 238, 227)))
                            g.FillPath(brush, path);
                        using (var brush = new SolidBrush(n == 0 ? Color.FromArgb(170, 30, 50) : Underline))
                        using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                            g.DrawString(label, font, brush, badge, format);
                        }
                    }
                }
            }
            // the folder switch
            Rectangle f = ToolFolderRect();
            bool hot = _toolHot == 4, down = _toolDown == 4;
            using (GraphicsPath path = RoundedRect(f, Px(ToolRadius)))
            {
                using (var brush = new SolidBrush(_folderSearch ? Color.FromArgb(70, Underline) : down ? Theme.TabBorder : Theme.PlusHoverBack)) g.FillPath(brush, path);
                using (var pen = new Pen(_folderSearch ? Underline : hot || down ? Theme.PlusColor : Theme.Separator, _folderSearch ? Math.Max(1.5f, Px(1.5)) : 1f)) g.DrawPath(pen, path);
            }
            DrawFolderIcon(g, f, _folderSearch ? Underline : Theme.PlusColor);
        }

        private void DrawFolderIcon(Graphics g, Rectangle r, Color color)
        {
            int w = Px(14), h = Px(10);
            int x = r.Left + (r.Width - w) / 2, y = r.Top + (r.Height - h) / 2 + Px(1);
            var points = new[]
            {
                new Point(x, y + h), new Point(x, y + Px(1)), new Point(x + Px(5), y + Px(1)), new Point(x + Px(7), y + Px(3)),
                new Point(x + w, y + Px(3)), new Point(x + w, y + h)
            };
            using (var pen = new Pen(color, Math.Max(1.4f, Px(1.4))) { LineJoin = LineJoin.Round })
            {
                g.DrawPolygon(pen, points);
            }
        }

        private void DrawNavButton(Graphics g, Rectangle r, bool back, bool enabled, bool hot, bool down)
        {
            using (GraphicsPath path = RoundedRect(r, Px(ToolRadius)))
            {
                using (var brush = new SolidBrush(enabled && down ? Theme.TabBorder : Theme.PlusHoverBack)) g.FillPath(brush, path);
                using (var pen = new Pen(enabled && (hot || down) ? Theme.PlusColor : Theme.Separator)) g.DrawPath(pen, path);
            }
            int cx = r.Left + r.Width / 2, cy = r.Top + r.Height / 2;
            int a = Px(3.5), b = Px(6.5);
            Color color = enabled ? Theme.PlusColor : Color.FromArgb(90, Theme.PlusColor);
            using (var pen = new Pen(color, Math.Max(1.8f, Px(1.8))) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                Point[] chevron = back
                    ? new[] { new Point(cx + a, cy - b), new Point(cx - a, cy), new Point(cx + a, cy + b) }
                    : new[] { new Point(cx - a, cy - b), new Point(cx + a, cy), new Point(cx - a, cy + b) };
                g.DrawLines(pen, chevron);
            }
        }

        // The sheets seen last, as a menu under the back button.
        private void ShowHistoryMenu()
        {
            List<HistoryItem> items = null;
            try { items = HistoryProvider != null ? HistoryProvider() : null; } catch (Exception) { }
            var menu = new ContextMenuStrip();
            menu.Items.Add(new ToolStripMenuItem("最近見たシート") { Enabled = false });
            if (items == null || items.Count == 0) menu.Items.Add(new ToolStripMenuItem("(まだありません)") { Enabled = false });
            else
            {
                foreach (HistoryItem item in items)
                {
                    HistoryItem it = item;
                    var mi = new ToolStripMenuItem((it.Current ? "✓ " : "　") + it.Name);
                    mi.Click += (s, e) => { if (HistoryJumpRequested != null) HistoryJumpRequested(it.Key); };
                    menu.Items.Add(mi);
                }
            }
            menu.Items.Add(new ToolStripSeparator());
            var clear = new ToolStripMenuItem("履歴をクリア") { Enabled = items != null && items.Count > 0 };
            clear.Click += (s, e) => { if (HistoryClearRequested != null) HistoryClearRequested(); };
            menu.Items.Add(clear);
            menu.Closed += (s, e) => BeginInvoke(new Action(menu.Dispose));
            Rectangle b = ToolBackRect();
            menu.Show(this, new Point(b.Left, b.Bottom + Px(2)));
        }

        private void StartLongPress()
        {
            if (_longPress == null)
            {
                _longPress = new System.Windows.Forms.Timer { Interval = 450 };
                _longPress.Tick += (s, e) =>
                {
                    _longPress.Stop();
                    if (_toolDown != 1 && _toolDown != 2) return;
                    _longPressShown = true;
                    _toolDown = 0;
                    Invalidate();
                    ShowHistoryMenu();
                };
            }
            _longPressShown = false;
            _longPress.Stop();
            _longPress.Start();
        }

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
                StripRow r = _baseRows.FirstOrDefault(x => x.IsFolder && x.FolderId == id);
                return r == null;
            }).ToList();
            foreach (string id in gone) _selFolders.Remove(id);
            if (_anchorFolder != null && !_baseRows.Any(x => x.IsFolder && x.FolderId == _anchorFolder)) _anchorFolder = null;
        }

        // Right of the open/closed mark: a click left of this opens or closes the folder.
        private int FolderMarkRight()
        {
            return Px(InsetLeft + TextPadLeft + FolderMarkWidth + 4);
        }

        public int IndexFromPoint(Point p)
        {
            if (p.Y < HeaderPx || p.Y >= ViewHeight || RowPx <= 0) return -1;
            int i = (p.Y - HeaderPx + _scrollY) / RowPx;
            return i >= 0 && i < _rows.Count ? i : -1;
        }

        private Rectangle RowRect(int index)
        {
            return new Rectangle(0, HeaderPx + index * RowPx - _scrollY, ContentWidth, RowPx);
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
            int view = AreaHeight;
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
                _scroll.SetBounds(ClientSize.Width - _scroll.Width, HeaderPx, _scroll.Width, view);
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
            int max = Math.Max(0, _rows.Count * RowPx - AreaHeight);
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
            if (r.Top < HeaderPx) SetScroll(_scrollY + r.Top - HeaderPx);
            else if (r.Bottom > ViewHeight) SetScroll(_scrollY + r.Bottom - ViewHeight);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            UpdateScrollBar();
            LayoutSearch();
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

            // a newly selected sheet, and the folders that have just been opened or closed, start the wobble (the first ones seen do not)
            if (!IsSample)
            {
                bool sheetChanged = false;
                if (_active != _lastActive)
                {
                    _lastActive = _active;
                    if (_active != null)
                    {
                        sheetChanged = _activeSeen;
                        _activeSeen = true;
                    }
                }
                HashSet<string> toggled = null;
                HashSet<string> opened = null; // the sheets of the folders that have just been opened wobble with them
                var seen = new Dictionary<string, bool>();
                foreach (StripRow r in _rows)
                {
                    if (!r.IsFolder) continue;
                    seen[r.FolderId] = r.Expanded;
                    bool before;
                    if (_folderOpen.TryGetValue(r.FolderId, out before) && before != r.Expanded)
                    {
                        if (toggled == null) toggled = new HashSet<string>();
                        toggled.Add(r.FolderId);
                        if (r.Expanded)
                        {
                            if (opened == null) opened = new HashSet<string>();
                            foreach (string m in r.Members) opened.Add(m);
                        }
                    }
                }
                _folderOpen = seen;
                // a folder that is selected or unselected wobbles as well (the folders seen at first, and the ones that are gone, do not)
                var selNow = new HashSet<string>(_selFolders.Where(seen.ContainsKey));
                if (_folderSelSeen)
                {
                    foreach (string id in selNow.Where(id => !_folderSel.Contains(id)).Concat(_folderSel.Where(id => !selNow.Contains(id) && seen.ContainsKey(id))))
                    {
                        if (toggled == null) toggled = new HashSet<string>();
                        toggled.Add(id);
                    }
                }
                _folderSel = selNow;
                _folderSelSeen = true;
                // (the row of a sheet that has just been added comes in with its own motion)
                bool wobbleSheet = sheetChanged && !(_raAdded != null && _active != null && _raAdded.Contains("S:" + _active));
                if ((wobbleSheet || toggled != null) && !_ghostOn) StartWobble(wobbleSheet ? _active : null, toggled, opened);
            }

            // the tabs scroll inside the area above the pinned "+" row
            g.SetClip(new Rectangle(0, HeaderPx, ClientSize.Width, Math.Max(0, ViewHeight - HeaderPx)));
            DrawFolderOutlines(g);
            int first = Math.Max(0, _scrollY / Math.Max(1, RowPx));
            var wobbling = new List<int>();
            var comingIn = new List<int>();
            for (int i = first; i < _rows.Count; i++)
            {
                Rectangle row = RowRect(i);
                row.Offset(0, RowAnimDy(i));
                if (row.Top > ViewHeight) break;
                if (row.Bottom < 0) continue;
                row = RowAnimIndent(row, _rows[i]);
                if (_raAdded != null && _raAdded.Contains(RowKey(_rows[i]))) comingIn.Add(i);
                else if (_animClock != null && (_rows[i].IsFolder ? (_animFolders != null && _animFolders.Contains(_rows[i].FolderId)) : (_rows[i].Sheet == _animSheet || (_animSheets != null && _rows[i].Sheet != null && _animSheets.Contains(_rows[i].Sheet))))) wobbling.Add(i);
                else
                {
                    if (_rows[i].IsFolder) DrawFolder(g, _rows[i], row); else DrawTab(g, _rows[i], row);
                    if (IsFadedRow(_rows[i])) DrawFadedPlace(g, row);
                }
            }
            DrawRowAnim(g, comingIn);
            DrawDoneAnim(g);
            foreach (int index in wobbling)
            {
                if (_animClock == null) break;
                Rectangle row = RowRect(index);
                row.Offset(0, RowAnimDy(index));
                WobbleFrame f = WobbleMotion.At(_animType, _animClock.ElapsedMilliseconds / (double)_wobbleDur, _wobbleAmp / 100.0);
                var pivot = f.PivotTopLeft
                    ? new PointF(row.Left + Px(InsetLeft), row.Top + Px(PlainTop))
                    : new PointF(row.Left + row.Width / 2f, row.Top + row.Height / 2f);
                GraphicsState state = g.Save();
                g.TranslateTransform(pivot.X + Px(f.Dx), pivot.Y + Px(f.Dy));
                g.RotateTransform((float)f.Angle);
                g.ScaleTransform((float)f.ScaleX, (float)f.ScaleY);
                g.TranslateTransform(-pivot.X, -pivot.Y);
                _transforming = true;
                try { if (_rows[index].IsFolder) DrawFolder(g, _rows[index], row); else DrawTab(g, _rows[index], row); }
                finally { _transforming = false; g.Restore(state); }
            }

            DrawCopyIcon(g, wobbling);

            if (_dropIndex >= 0 && _rows.Count > 0)
            {
                int y = _dropIndex < _rows.Count ? RowRect(_dropIndex).Top : RowRect(_rows.Count - 1).Bottom;
                DrawDropMarker(g, y, _dropInside ? Px(FolderIndent) : 0);
            }
            DrawGhost(g);
            DrawSearchChoice(g);
            g.ResetClip();
            if (!IsSample) DrawToolRow(g);

            if (ShowAllButtons) DrawAllButtons(g);
            if (!IsSample) DrawPlus(g);
        }

        // ---- the burst of a progress mark that turns to "done" ----
        // The mark swells a little and falls back, rays and dots fly out of it, and a soft green light crosses the row. 0.5 seconds, no sound.

        private const int DoneAnimMs = 500;
        private bool _doneAnimation = true;
        private System.Diagnostics.Stopwatch _dnClock;
        private HashSet<string> _dnSheets;
        private System.Windows.Forms.Timer _dnTimer;

        public bool DoneAnimation
        {
            get { return _doneAnimation; }
            set { _doneAnimation = value; if (!value) EndDoneAnim(); }
        }

        private void StartDoneAnim(IEnumerable<string> sheets)
        {
            if (!IsHandleCreated || !base.Visible) return;
            _dnSheets = new HashSet<string>(sheets);
            _dnClock = System.Diagnostics.Stopwatch.StartNew();
            if (_dnTimer == null)
            {
                _dnTimer = new System.Windows.Forms.Timer { Interval = 15 };
                _dnTimer.Tick += (s, e) =>
                {
                    if (_dnClock == null || _dnClock.ElapsedMilliseconds >= DoneAnimMs) EndDoneAnim();
                    Invalidate();
                };
            }
            _dnTimer.Start();
        }

        private void EndDoneAnim()
        {
            if (_dnTimer != null) _dnTimer.Stop();
            bool was = _dnClock != null;
            _dnClock = null;
            _dnSheets = null;
            if (was) Invalidate();
        }

        private void DrawDoneAnim(Graphics g)
        {
            if (_dnClock == null || _dnSheets == null) return;
            double t = Math.Min(1.0, _dnClock.ElapsedMilliseconds / (double)DoneAnimMs);
            int first = Math.Max(0, _scrollY / Math.Max(1, RowPx));
            for (int i = first; i < _rows.Count; i++)
            {
                StripRow r = _rows[i];
                Rectangle row = RowRect(i);
                row.Offset(0, RowAnimDy(i));
                if (row.Top > ViewHeight) break;
                if (row.Bottom < 0 || r.IsFolder || r.Sheet == null || !_dnSheets.Contains(r.Sheet)) continue;
                RowBadgeSpec spec = RowBadges(r, row).FirstOrDefault(s => s.Kind == RowBadgeKind.Progress);
                if (spec == null) continue;
                int doneState = spec.State == 4 ? 4 : 3;
                var mode = g.SmoothingMode;
                g.SmoothingMode = SmoothingMode.AntiAlias;
                // a soft green light over the row
                double u = Clamp01((t - 0.1) / 0.7);
                int left = row.Left + Px(InsetLeft) + (r.Indented ? Px(FolderIndent) : 0), right = row.Right - RightInset(r);
                var body = new Rectangle(left, row.Top + Px(PlainTop), Math.Max(2, right - left), row.Height - Px(PlainTop) - Px(BottomInset));
                using (GraphicsPath path = RoundedRect(body, Px(CornerRadius)))
                using (var fill = new SolidBrush(Color.FromArgb((int)(60 * (1 - u)), BadgeProgressGreen)))
                {
                    g.FillPath(fill, path);
                }
                // the mark swells and falls back
                double pop = t < 0.1 ? 1.0 : 1 + 0.55 * Math.Sin(Math.PI * Clamp01((t - 0.1) / 0.45));
                int grow = (int)Math.Round(spec.Bounds.Width * (pop - 1) / 2);
                DrawProgressBadge(g, Rectangle.Inflate(spec.Bounds, grow, grow), doneState);
                // rays and dots
                if (t > 0.1 && t < 0.85)
                {
                    double k = (t - 0.1) / 0.75;
                    float cx = spec.Bounds.Left + spec.Bounds.Width / 2f, cy = spec.Bounds.Top + spec.Bounds.Height / 2f;
                    int alpha = (int)(255 * (1 - k));
                    for (int d = 0; d < 8; d++)
                    {
                        double ang = d * Math.PI / 4;
                        double r0 = Px(10 + 12 * k), r1 = Px(14 + 16 * k);
                        using (var pen = new Pen(Color.FromArgb(alpha, BadgeProgressGreen), Math.Max(1.5f, Px(1.6))) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                        {
                            g.DrawLine(pen, cx + (float)(Math.Cos(ang) * r0), cy + (float)(Math.Sin(ang) * r0), cx + (float)(Math.Cos(ang) * r1), cy + (float)(Math.Sin(ang) * r1));
                        }
                        double ang2 = ang + Math.PI / 8, rr = Px(8 + 20 * k);
                        float dot = Math.Max(1.5f, Px(1.8));
                        using (var b = new SolidBrush(Color.FromArgb(alpha, 245, 190, 40)))
                        {
                            g.FillEllipse(b, cx + (float)(Math.Cos(ang2) * rr) - dot, cy + (float)(Math.Sin(ang2) * rr) - dot, dot * 2, dot * 2);
                        }
                    }
                }
                g.SmoothingMode = mode;
            }
        }

        // ---- the rows of added and deleted sheets and folders ----
        // An added row: the rows below make room (they slide down), and the new row grows in with a little spring, fading in, with a
        // ring of light around it. A deleted row: it swells a little, shrinks and fades away, a few dots fly out, and the rows below
        // slide up. The sheets that move into a folder, or out of a deleted one, slide sideways to their new place. It runs 0.4 seconds
        // (the length is fixed) and makes no sound. It is off unless it is chosen in the settings.
        // Any number of rows may come or go: only the rows that are in view are drawn (and pictured, once); the rows out of view take
        // their part in the sliding of the rows below without being drawn.

        private const int RowAnimMs = 400;
        private bool _rowAnimation;
        private System.Windows.Forms.Timer _raTimer;
        private System.Diagnostics.Stopwatch _raClock;
        private HashSet<string> _raAdded;                 // the keys ("S:name", "F:id") of the rows that come in
        private List<RowAnimGhost> _raGhosts;
        private Dictionary<string, bool> _raIndent;       // sheet name -> whether it was indented before
        private int[] _raAddedAt;                         // the indexes of the rows that come in, in order (for the rows below them)
        private int[] _raGhostAt;                         // the At of the rows that go, in order
        private Dictionary<string, Bitmap> _raPictures;   // the pictures of the rows that come in (made when first drawn)

        private class RowAnimGhost
        {
            public Bitmap Bmp;  // null for a row that was out of view (it is not drawn)
            public int At;      // the number of rows of the new list that stand before it
            public Color Color;
        }

        private class RowAnimPending
        {
            public HashSet<string> Added = new HashSet<string>();
            public List<RowAnimGhost> Ghosts = new List<RowAnimGhost>();
            public Dictionary<string, bool> OldIndented = new Dictionary<string, bool>();
        }

        public bool RowAnimation
        {
            get { return _rowAnimation; }
            set { _rowAnimation = value; if (!value) EndRowAnim(); }
        }

        private static string RowKey(StripRow r) { return r.IsFolder ? "F:" + r.FolderId : "S:" + r.Sheet; }

        private static double EaseOut(double t) { t = Math.Max(0, Math.Min(1, t)); return 1 - Math.Pow(1 - t, 3); }

        private static double EaseOutBack(double t)
        {
            t = Math.Max(0, Math.Min(1, t)) - 1;
            const double k = 2.2;
            return 1 + (k + 1) * t * t * t + k * t * t;
        }

        private static double Clamp01(double t) { return Math.Max(0, Math.Min(1, t)); }

        private double RowAnimT { get { return _raClock == null ? 1 : Math.Min(1.0, _raClock.ElapsedMilliseconds / (double)RowAnimMs); } }

        // How far the rows below (and the sheets that change their indent) have come on their way, 0 to 1.
        private double RowAnimSlide()
        {
            double t = RowAnimT;
            return _raAdded != null ? EaseOut(t / 0.5) : EaseOut((t - 0.35) / 0.6);
        }

        private Color RowAnimColor(StripRow r)
        {
            if (r.IsFolder) return r.ColorBgr >= 0 ? Color.FromArgb(r.ColorBgr & 0xFF, (r.ColorBgr >> 8) & 0xFF, (r.ColorBgr >> 16) & 0xFF) : Underline;
            int bgr;
            return r.Sheet != null && _colors.TryGetValue(r.Sheet, out bgr) ? Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF) : Underline;
        }

        // Called with the old state still in place: which rows came and went (sheets by name, folders by id), and (of those that went) a
        // picture of their rows.
        private RowAnimPending PrepareRowAnim(IList<string> names, IList<StripRow> newRows)
        {
            if (!_rowAnimation || IsSample || _names.Count == 0 || Filtering || !IsHandleCreated || !base.Visible || ClientSize.Width < 20 || RowPx <= 0) return null;
            var oldFolders = new HashSet<string>(_rows.Where(r => r.IsFolder).Select(r => r.FolderId));
            var newFolders = new HashSet<string>(newRows.Where(r => r.IsFolder).Select(r => r.FolderId));
            var added = names.Where(n => !_names.Contains(n)).Select(n => "S:" + n).Concat(newFolders.Where(f => !oldFolders.Contains(f)).Select(f => "F:" + f)).ToList();
            var removed = new HashSet<string>(_names.Where(n => !names.Contains(n)).Select(n => "S:" + n).Concat(oldFolders.Where(f => !newFolders.Contains(f)).Select(f => "F:" + f)));
            // only rows that came, or only rows that went (a rename or a regrouping does both)
            if ((added.Count == 0) == (removed.Count == 0)) return null;
            var p = new RowAnimPending { Added = new HashSet<string>(added) };
            foreach (StripRow r in _rows) if (!r.IsFolder && r.Sheet != null) p.OldIndented[r.Sheet] = r.Indented;
            int before = 0;
            for (int i = 0; i < _rows.Count; i++)
            {
                StripRow r = _rows[i];
                if (!removed.Contains(RowKey(r))) continue;
                // a picture only of a row the user sees now; the others go without being drawn
                Rectangle rect = RowRect(i);
                bool inView = rect.Bottom > HeaderPx && rect.Top < ViewHeight;
                p.Ghosts.Add(new RowAnimGhost { Bmp = inView ? RenderRowBitmap(r, ContentWidth) : null, At = i - before, Color = RowAnimColor(r) });
                before++;
            }
            if (added.Count == 0 && p.Ghosts.Count == 0) return null;
            return p;
        }

        // Called with the new rows in place: starts the motion.
        private void FinishRowAnim(RowAnimPending p)
        {
            if (p == null) return;
            EndRowAnim();
            if (p.Added.Count > 0)
            {
                var set = new HashSet<string>(_rows.Select(RowKey).Where(k => p.Added.Contains(k)));
                if (set.Count == 0) return; // (the sheet is in a closed folder: no row)
                _raAdded = set;
                var at = new List<int>();
                for (int i = 0; i < _rows.Count; i++) if (set.Contains(RowKey(_rows[i]))) at.Add(i);
                _raAddedAt = at.ToArray();
            }
            else
            {
                foreach (RowAnimGhost g in p.Ghosts) g.At = Math.Min(g.At, _rows.Count);
                _raGhosts = p.Ghosts;
                _raGhostAt = p.Ghosts.Select(g => g.At).ToArray(); // (in order: the rows that go were taken from the top down)
            }
            // the sheets whose indent changes (a folder came or went around them) slide to their new place
            var indent = new Dictionary<string, bool>();
            foreach (StripRow r in _rows)
            {
                bool was;
                if (!r.IsFolder && r.Sheet != null && p.OldIndented.TryGetValue(r.Sheet, out was) && was != r.Indented) indent[r.Sheet] = was;
            }
            _raIndent = indent.Count > 0 ? indent : null;
            _raClock = System.Diagnostics.Stopwatch.StartNew();
            if (_raTimer == null)
            {
                _raTimer = new System.Windows.Forms.Timer { Interval = 15 };
                _raTimer.Tick += (s, e) =>
                {
                    if (_raClock == null || _raClock.ElapsedMilliseconds >= RowAnimMs) EndRowAnim();
                    Invalidate();
                };
            }
            _raTimer.Start();
        }

        private void EndRowAnim()
        {
            if (_raTimer != null) _raTimer.Stop();
            bool was = _raClock != null;
            _raClock = null;
            _raAdded = null;
            _raIndent = null;
            _raAddedAt = null;
            _raGhostAt = null;
            DropRowAnimPictures();
            if (_raGhosts != null)
            {
                foreach (RowAnimGhost g in _raGhosts) if (g.Bmp != null) g.Bmp.Dispose();
                _raGhosts = null;
            }
            if (was) Invalidate();
        }

        // The pictures of the rows that come in are made again when the rows' look changes (a selection, a color, a badge).
        private void DropRowAnimPictures()
        {
            if (_raPictures == null) return;
            foreach (Bitmap bmp in _raPictures.Values) bmp.Dispose();
            _raPictures = null;
        }

        // The number of values in a sorted array that are below (or, with orEqual, not above) the value.
        private static int CountBelow(int[] sorted, int value, bool orEqual)
        {
            int lo = 0, hi = sorted.Length;
            while (lo < hi)
            {
                int mid = (lo + hi) / 2;
                if (sorted[mid] < value || (orEqual && sorted[mid] == value)) lo = mid + 1; else hi = mid;
            }
            return lo;
        }

        // How far a row is from its place (the rows below an added or a deleted one are on their way).
        private int RowAnimDy(int index)
        {
            if (_raClock == null) return 0;
            if (_raAddedAt != null)
            {
                int before = CountBelow(_raAddedAt, index, false);
                return -(int)Math.Round(RowPx * before * (1 - RowAnimSlide()));
            }
            if (_raGhostAt != null)
            {
                int before = CountBelow(_raGhostAt, index, true);
                return (int)Math.Round(RowPx * before * (1 - RowAnimSlide()));
            }
            return 0;
        }

        // A sheet that moves into a folder, or out of one: its left edge slides from the old indent to the new one.
        private Rectangle RowAnimIndent(Rectangle row, StripRow r)
        {
            bool was;
            if (_raClock == null || _raIndent == null || r.IsFolder || r.Sheet == null || !_raIndent.TryGetValue(r.Sheet, out was)) return row;
            int dx = (was ? Px(FolderIndent) : 0) - (r.Indented ? Px(FolderIndent) : 0);
            dx = (int)Math.Round(dx * (1 - RowAnimSlide()));
            return new Rectangle(row.Left + dx, row.Top, row.Width - dx, row.Height);
        }

        // A picture of a row (transparent around it), as it looks now.
        private Bitmap RenderRowBitmap(StripRow r, int width)
        {
            var bmp = new Bitmap(Math.Max(1, width), Math.Max(1, RowPx), System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Color.Transparent);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                _transforming = true; // (text without ClearType: it cannot be drawn on a transparent picture)
                try
                {
                    var local = new Rectangle(0, 0, bmp.Width, bmp.Height);
                    if (r.IsFolder) DrawFolder(g, r, local); else DrawTab(g, r, local);
                }
                finally { _transforming = false; }
            }
            return bmp;
        }

        private void DrawRowImage(Graphics g, Bitmap bmp, Rectangle row, double sx, double sy, double alpha)
        {
            if (sx < 0.03 || sy < 0.03 || alpha < 0.02) return;
            int w = Math.Max(1, (int)(row.Width * sx)), h = Math.Max(1, (int)(row.Height * sy));
            var dest = new Rectangle(row.Left + (row.Width - w) / 2, row.Top + (row.Height - h) / 2, w, h);
            using (var attr = new System.Drawing.Imaging.ImageAttributes())
            {
                attr.SetColorMatrix(new System.Drawing.Imaging.ColorMatrix { Matrix33 = (float)Clamp01(alpha) });
                var mode = g.InterpolationMode;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.DrawImage(bmp, dest, 0, 0, bmp.Width, bmp.Height, GraphicsUnit.Pixel, attr);
                g.InterpolationMode = mode;
            }
        }

        // The rows that come in, and the rows that go.
        private void DrawRowAnim(Graphics g, List<int> comingIn)
        {
            if (_raClock == null) return;
            double t = RowAnimT;
            foreach (int index in comingIn)
            {
                Rectangle row = RowRect(index);
                double s = EaseOutBack((t - 0.12) / 0.6), a = (t - 0.12) / 0.2;
                if (_raPictures == null) _raPictures = new Dictionary<string, Bitmap>();
                string key = RowKey(_rows[index]);
                Bitmap bmp;
                if (!_raPictures.TryGetValue(key, out bmp) || bmp.Width != Math.Max(1, row.Width))
                {
                    if (bmp != null) bmp.Dispose();
                    bmp = RenderRowBitmap(_rows[index], row.Width);
                    _raPictures[key] = bmp;
                }
                DrawRowImage(g, bmp, row, Math.Min(1.0, s), s, a);
                // a ring of light that spreads and fades
                if (t > 0.1 && t < 0.7)
                {
                    double k = (t - 0.1) / 0.6;
                    int grow = (int)(Px(10) * k);
                    Color c = RowAnimColor(_rows[index]);
                    var ring = new Rectangle(row.Left + Px(InsetLeft) - grow, row.Top + Px(PlainTop) - grow, row.Width - Px(InsetLeft + InsetRight) + grow * 2, row.Height - Px(PlainTop + BottomInset) + grow * 2);
                    using (GraphicsPath path = RoundedRect(ring, Px(CornerRadius) + grow))
                    using (var pen = new Pen(Color.FromArgb((int)(130 * (1 - k)), c), Math.Max(1.5f, Px(2))))
                    {
                        g.DrawPath(pen, path);
                    }
                }
            }
            if (_raGhosts == null) return;
            for (int j = 0; j < _raGhosts.Count; j++)
            {
                RowAnimGhost gh = _raGhosts[j];
                if (gh.Bmp == null) continue; // (out of view when it went)
                var row = new Rectangle(0, HeaderPx + (gh.At + j) * RowPx - _scrollY, gh.Bmp.Width, gh.Bmp.Height);
                if (row.Bottom < HeaderPx - RowPx || row.Top > ViewHeight + RowPx) continue; // (out of view now: the particles too)
                double shrink = EaseOut((t - 0.1) / 0.55);
                double sy = t < 0.1 ? 1 + 0.06 * (t / 0.1) : 1 - shrink;
                double sx = t < 0.1 ? sy : 1 - 0.25 * shrink;
                DrawRowImage(g, gh.Bmp, row, sx, sy, 1 - shrink * 1.1);
                if (t > 0.2 && t < 0.85)
                {
                    double k = (t - 0.2) / 0.65;
                    float cx = row.Left + row.Width / 2f, cy = row.Top + row.Height / 2f;
                    var mode = g.SmoothingMode;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    for (int d = 0; d < 8; d++)
                    {
                        double ang = d * Math.PI / 4 + 0.3;
                        double dist = Px(5) + Px(24) * k;
                        float r = (float)(Px(2.2) * (1 - 0.75 * k));
                        float px = cx + (float)(Math.Cos(ang) * dist * 2.4), py = cy + (float)(Math.Sin(ang) * dist);
                        using (var b = new SolidBrush(Color.FromArgb((int)(220 * (1 - k)), gh.Color))) g.FillEllipse(b, px - r, py - r, r * 2, r * 2);
                    }
                    g.SmoothingMode = mode;
                }
            }
        }

        // ---- the drag ghost: while rows are dragged, a half transparent copy of the first one follows the pointer ----

        private const double GhostAlpha = 0.62;
        private bool _ghostOn;
        private bool _ghostSettling;
        private Bitmap _ghostBmp;
        private PointF _ghostPos, _ghostTarget;
        private PointF _ghostGrab;
        private int _ghostRowLeft, _ghostRowTop;
        private int _ghostCount;
        private bool _dragCopy;
        // what wobbles after the drop (the rows change a moment later, and a sheet may get a new name)
        private HashSet<string> _wobbleSheets, _wobbleFolders, _dropNames;
        private string _dropSignature;
        private HashSet<string> _fadeSheets = new HashSet<string>();
        private HashSet<string> _fadeFolders = new HashSet<string>();
        private System.Diagnostics.Stopwatch _ghostClock;
        private long _ghostLast, _settleStart;

        private int GhostPad { get { return Px(10); } }

        // The first dragged row is drawn once into a bitmap (with room around it for the shadow).
        private void BeginGhost(int index)
        {
            EndGhost();
            if (IsSample || index < 0 || index >= _rows.Count) return;
            Rectangle row = RowRect(index);
            if (row.Width < 4 || row.Height < 4) return;
            int pad = GhostPad;
            var bmp = new Bitmap(row.Width, row.Height, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.Clear(Theme.BarBack);
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.PixelOffsetMode = PixelOffsetMode.Half;
                var local = new Rectangle(0, 0, row.Width, row.Height);
                if (_rows[index].IsFolder) DrawFolder(g, _rows[index], local); else DrawTab(g, _rows[index], local);
            }
            _ghostBmp = bmp;
            _ghostRowLeft = row.Left;
            _ghostRowTop = row.Top;
            _ghostGrab = new PointF(_dragOrigin.X - row.Left, _dragOrigin.Y - row.Top);
            _ghostPos = _ghostTarget = new PointF(row.Left, row.Top);
            _ghostCount = _dragSheets.Count + _dragFolders.Count;
            _ghostSettling = false;
            _ghostOn = true;
            _ghostClock = System.Diagnostics.Stopwatch.StartNew();
            _ghostLast = 0;
            // the places the rows are dragged from stay faint (not while copying)
            _fadeSheets = new HashSet<string>(_dragSheets);
            _fadeFolders = new HashSet<string>(_dragFolders);
            foreach (StripRow r in _rows)
                if (r.IsFolder && _dragFolders.Contains(r.FolderId)) foreach (string m in r.Members) _fadeSheets.Add(m);
        }

        private void EndGhost()
        {
            _ghostOn = false;
            _ghostSettling = false;
            if (_ghostBmp != null) { _ghostBmp.Dispose(); _ghostBmp = null; }
            _fadeSheets.Clear();
            _fadeFolders.Clear();
            _wobbleSheets = _wobbleFolders = _dropNames = null;
            _ghostClock = null;
            Invalidate();
        }

        // The ghost follows the pointer (a little sideways, the rows are as wide as the pane).
        private void MoveGhost(Point client)
        {
            if (!_ghostOn || _ghostSettling) return;
            float dx = client.X - _dragOrigin.X;
            float limit = Px(8);
            dx = Math.Max(-limit, Math.Min(limit, dx));
            _ghostTarget = new PointF(_ghostRowLeft + dx, client.Y - _ghostGrab.Y);
            Invalidate();
        }

        // The drop: the ghost glides into the place and turns solid.
        private void SettleGhost(int markerY, bool movedDown, int rowCount)
        {
            if (!_ghostOn) return;
            _ghostSettling = true;
            _ghostTarget = new PointF(_ghostRowLeft, movedDown ? markerY - Math.Max(1, rowCount) * RowPx : markerY);
            _settleStart = _ghostClock.ElapsedMilliseconds;
            Invalidate();
        }

        private bool IsFadedRow(StripRow r)
        {
            if (!_ghostOn || _dragCopy) return false;
            return r.IsFolder ? _fadeFolders.Contains(r.FolderId) : (r.Sheet != null && _fadeSheets.Contains(r.Sheet));
        }

        // Drawn on top of the rows, inside the clip of the rows.
        private void DrawGhost(Graphics g)
        {
            if (!_ghostOn || _ghostBmp == null) return;
            long now = _ghostClock.ElapsedMilliseconds;
            double dt = Math.Max(1, now - _ghostLast);
            _ghostLast = now;
            double tau = _ghostSettling ? 38 : 55; // smoothing time, in milliseconds
            double k = 1 - Math.Exp(-dt / tau);
            float dxp = _ghostTarget.X - _ghostPos.X, dyp = _ghostTarget.Y - _ghostPos.Y;
            _ghostPos = new PointF(_ghostPos.X + (float)(dxp * k), _ghostPos.Y + (float)(dyp * k));
            bool arrived = Math.Abs(dxp) < 1.2 && Math.Abs(dyp) < 1.2;

            double alpha = GhostAlpha;
            if (_ghostSettling)
            {
                double t = Math.Min(1.0, (now - _settleStart) / 140.0);
                alpha = GhostAlpha + (1 - GhostAlpha) * t;
                // the ghost stays on its place until the rows show the new order (or for a short while at most)
                if (arrived && (_rowSignature != _dropSignature || now - _settleStart > 700))
                {
                    var sheets = _wobbleSheets ?? new HashSet<string>();
                    var folders = _wobbleFolders ?? new HashSet<string>();
                    if (_dropNames != null) foreach (string n in _names) if (!_dropNames.Contains(n)) sheets.Add(n); // renamed or copied sheets
                    EndGhost();
                    if (sheets.Count > 0 || folders.Count > 0) StartWobble(null, folders.Count > 0 ? folders : null, sheets.Count > 0 ? sheets : null);
                    return;
                }
            }

            int w = _ghostBmp.Width, h = _ghostBmp.Height;
            var dest = new Rectangle((int)Math.Round(_ghostPos.X), (int)Math.Round(_ghostPos.Y), w, h);
            int radius = Px(8);
            // soft shadow around and below the ghost
            for (int i = 4; i >= 1; i--)
            {
                int grow = Px(i * 1.4);
                var sr = new Rectangle(dest.Left + Px(3) - grow, dest.Top + Px(6) - grow / 2, w - Px(6) + grow * 2, h - Px(10) + grow);
                using (GraphicsPath path = RoundedRect(sr, radius + grow))
                using (var brush = new SolidBrush(Color.FromArgb((int)(10 * (_ghostSettling ? 1 - 0.5 * (alpha - GhostAlpha) / (1 - GhostAlpha) : 1)), 0, 0, 0)))
                    g.FillPath(brush, path);
            }
            var shape = new Rectangle(dest.Left + Px(InsetLeft) - Px(2), dest.Top + Px(PlainTop), w - Px(InsetLeft) * 2 + Px(4), h - Px(PlainTop) - Px(BottomInset));
            GraphicsState state = g.Save();
            using (GraphicsPath clip = RoundedRect(shape, radius))
            {
                g.SetClip(clip, CombineMode.Intersect);
                using (var attr = new System.Drawing.Imaging.ImageAttributes())
                {
                    var matrix = new System.Drawing.Imaging.ColorMatrix { Matrix33 = (float)alpha };
                    attr.SetColorMatrix(matrix, System.Drawing.Imaging.ColorMatrixFlag.Default, System.Drawing.Imaging.ColorAdjustType.Bitmap);
                    g.DrawImage(_ghostBmp, dest, 0, 0, w, h, GraphicsUnit.Pixel, attr);
                }
            }
            g.Restore(state);
            using (GraphicsPath edge = RoundedRect(shape, radius))
            using (var pen = new Pen(Color.FromArgb((int)(70 * alpha), 0, 0, 0), 1f)) g.DrawPath(pen, edge);

            // badges: the number of dragged rows, and "+" while copying
            if (_ghostCount > 1) DrawGhostBadge(g, _ghostCount.ToString(), new Point(shape.Right - Px(4), shape.Top), alpha);
            if (_dragCopy) DrawGhostBadge(g, "+", new Point(shape.Left + Px(4), shape.Top), alpha);

            if (!arrived || _ghostSettling) Invalidate();
        }

        private void DrawGhostBadge(Graphics g, string text, Point center, double alpha)
        {
            int d = Px(18);
            var box = new Rectangle(center.X - d / 2, center.Y - d / 2, d, d);
            using (var brush = new SolidBrush(Color.FromArgb(235, 16, 124, 65))) g.FillEllipse(brush, box);
            using (var pen = new Pen(Color.White, Math.Max(1f, Px(1.2)))) g.DrawEllipse(pen, box);
            using (var font = new Font(TabFont.FontFamily, TabFont.Size * 0.8f, FontStyle.Bold, TabFont.Unit))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoWrap })
            {
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.DrawString(text, font, Brushes.White, box, format);
            }
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

        private bool Crystal { get { return _selectStyle == SelectStyle.Crystal; } }

        // ---- the badges at the right end of the rows ----

        private BadgeKinds _shownBadges = BadgeKinds.None;
        private Dictionary<string, BadgeInfo> _rowBadges = new Dictionary<string, BadgeInfo>();
        private string _badgeTip;

        public BadgeKinds ShownBadges
        {
            get { return _shownBadges; }
            set { if (_shownBadges != value) { _shownBadges = value; Invalidate(); } }
        }

        // What the sheets hold, by sheet name (null = nothing is known).
        public void SetBadges(IDictionary<string, BadgeInfo> info)
        {
            var copy = new Dictionary<string, BadgeInfo>();
            if (info != null) foreach (var kv in info) copy[kv.Key] = new BadgeInfo { Errors = kv.Value.Errors, Memos = kv.Value.Memos, Protected = kv.Value.Protected, Progress = kv.Value.Progress }; // (copies: the caller changes its own)
            bool same = copy.Count == _rowBadges.Count && copy.All(kv => { BadgeInfo o; return _rowBadges.TryGetValue(kv.Key, out o) && o.Errors == kv.Value.Errors && o.Memos == kv.Value.Memos && o.Protected == kv.Value.Protected && o.Progress == kv.Value.Progress; });
            if (same) return;
            // the sheets that have just been marked as done burst (the ones seen first, at the start, do not)
            List<string> becameDone = null;
            if (_doneAnimation && !IsSample && (_shownBadges & BadgeKinds.Progress) != 0)
            {
                foreach (var kv in copy)
                {
                    BadgeInfo o;
                    if (kv.Value.Progress == 3 && _rowBadges.TryGetValue(kv.Key, out o) && o.Progress != 3)
                    {
                        if (becameDone == null) becameDone = new List<string>();
                        becameDone.Add(kv.Key);
                    }
                }
            }
            _rowBadges = copy;
            if (becameDone != null) StartDoneAnim(becameDone);
            DropRowAnimPictures();
            Invalidate();
        }

        private static readonly Color BadgeErrorColor = Color.FromArgb(214, 69, 65), BadgeMemoColor = Color.FromArgb(232, 168, 28), BadgeProtectColor = Color.FromArgb(78, 110, 150);

        private enum RowBadgeKind { Progress, ProgressPill, Error, Memo, Protect, Count }

        private class RowBadgeSpec
        {
            public RowBadgeKind Kind;
            public int Count;
            public int Total;
            public int State; // the progress (1 to 3)
            public Color Color;
            public Rectangle Bounds;
            public string Tip;
        }

        private const double RowBadgePillWidth = 34, RowBadgePillHeight = 17;
        // a little more room to the left of a folder's progress pill, before the badge with the number of sheets
        private const double RowBadgePillGap = 4;

        // the done + lock mark and the count badge of a folder are a little larger than the others
        private const double BadgeBigScale = 1.2;
        private static readonly Color BadgeProgressGray = Color.FromArgb(150, 150, 155), BadgeProgressBlue = Color.FromArgb(58, 128, 233), BadgeProgressGreen = Color.FromArgb(16, 124, 65);
        private static readonly string[] ProgressNames = { "なし", "未着手", "作業中", "完了" };

        private bool IsActiveRow(StripRow r) { return !r.IsFolder && _selected.Contains(r.Sheet) && r.Sheet == _active; }

        // The badges of a row from the right, left to right: the progress comes first, at the right end, then what each
        // other badge tells. A folder shows a pill with the number of its sheets that are done out of those that have a progress; a closed
        // folder also shows one round badge with the number of its sheets that have something else to tell (red when one of them has an
        // error, else amber).
        private List<RowBadgeSpec> RowBadges(StripRow r, Rectangle row)
        {
            var list = new List<RowBadgeSpec>();
            if (_shownBadges == BadgeKinds.None || r == null) return list;
            int radius = Px(RowBadgeRadius);
            int gap = Px(3);
            int cy = BodyCenterY(r, row);
            int right = BadgeRightEdge(r, row);
            var found = new List<RowBadgeSpec>();
            if (r.IsFolder)
            {
                if ((_shownBadges & BadgeKinds.Progress) != 0)
                {
                    int done = 0, total = 0;
                    foreach (string m in r.Members)
                    {
                        BadgeInfo b;
                        if (!_rowBadges.TryGetValue(m, out b) || b.Progress == 0) continue;
                        total++;
                        if (b.Progress == 3) done++;
                    }
                    if (total > 0)
                        found.Add(new RowBadgeSpec { Kind = RowBadgeKind.ProgressPill, Count = done, Total = total, Tip = "進捗: 完了 " + done + " / " + total });
                }
                if (!r.Expanded)
                {
                    int errors = 0, memos = 0, locks = 0, sheets = 0;
                    foreach (string m in r.Members)
                    {
                        BadgeInfo b;
                        if (!_rowBadges.TryGetValue(m, out b)) continue;
                        bool e = (_shownBadges & BadgeKinds.Error) != 0 && b.Errors > 0, n = (_shownBadges & BadgeKinds.Memo) != 0 && b.Memos > 0, p = (_shownBadges & BadgeKinds.Protect) != 0 && b.Protected;
                        if (e) errors++;
                        if (n) memos++;
                        if (p) locks++;
                        if (e || n || p) sheets++;
                    }
                    if (sheets > 0)
                    {
                        var lines = new List<string>();
                        if (errors > 0) lines.Add("数式エラーのあるシート " + errors + " 枚");
                        if (memos > 0) lines.Add("メモのあるシート " + memos + " 枚");
                        if (locks > 0) lines.Add("保護されたシート " + locks + " 枚");
                        found.Add(new RowBadgeSpec
                        {
                            Kind = RowBadgeKind.Count,
                            Count = sheets,
                            Color = errors > 0 ? BadgeErrorColor : memos > 0 ? BadgeMemoColor : BadgeProtectColor,
                            Tip = string.Join("\n", lines)
                        });
                    }
                }
            }
            else
            {
                BadgeInfo b;
                if (!_rowBadges.TryGetValue(r.Sheet, out b)) return list;
                // done and protected: one mark, a padlock that carries the check, and no lock of its own
                bool doneLocked = (_shownBadges & BadgeKinds.Progress) != 0 && b.Progress == 3 && b.Protected;
                if ((_shownBadges & BadgeKinds.Progress) != 0 && b.Progress > 0)
                    found.Add(new RowBadgeSpec
                    {
                        Kind = RowBadgeKind.Progress,
                        State = doneLocked ? 4 : b.Progress,
                        Tip = doneLocked ? "進捗: 完了 (シートは保護されています。解除は、右クリックのメニューか、リボンから)" : b.Progress == 3 ? "進捗: 完了 (クリックで、シートの保護)" : "進捗: " + ProgressNames[b.Progress] + " (クリックで、次へ)"
                    });
                if ((_shownBadges & BadgeKinds.Error) != 0 && b.Errors > 0)
                    found.Add(new RowBadgeSpec { Kind = RowBadgeKind.Error, Color = BadgeErrorColor, Tip = "数式エラーが " + b.Errors + " か所" });
                if ((_shownBadges & BadgeKinds.Memo) != 0 && b.Memos > 0)
                    found.Add(new RowBadgeSpec { Kind = RowBadgeKind.Memo, Color = BadgeMemoColor, Tip = "メモ・コメントが " + b.Memos + " 件" });
                if ((_shownBadges & BadgeKinds.Protect) != 0 && b.Protected && !doneLocked)
                    found.Add(new RowBadgeSpec { Kind = RowBadgeKind.Protect, Color = BadgeProtectColor, Tip = "シートは保護されています" });
            }
            // the first one stands at the right end, the next ones to its left
            int x = right;
            foreach (RowBadgeSpec s in found)
            {
                bool big = s.Kind == RowBadgeKind.Count || (s.Kind == RowBadgeKind.Progress && s.State == 4);
                int side = big ? (int)Math.Round(radius * 2 * BadgeBigScale) : radius * 2;
                int w = s.Kind == RowBadgeKind.ProgressPill ? Px(RowBadgePillWidth) : side;
                int h = s.Kind == RowBadgeKind.ProgressPill ? Px(RowBadgePillHeight) : side;
                s.Bounds = new Rectangle(x - w, cy - h / 2, w, h);
                list.Add(s);
                x -= w + gap + (s.Kind == RowBadgeKind.ProgressPill ? Px(RowBadgePillGap) : 0);
            }
            return list;
        }

        // The middle of the body of a row (a selected row is a little higher than the others).
        private int BodyCenterY(StripRow r, Rectangle row)
        {
            bool selected = r.IsFolder ? _selFolders.Contains(r.FolderId) : _selected.Contains(r.Sheet);
            int top = row.Top + Px(selected ? PlainTop : ColoredTop), bottom = row.Bottom - Px(BottomInset);
            return (top + bottom) / 2;
        }

        // Where the badges of a row start, from the right.
        private int BadgeRightEdge(StripRow r, Rectangle row)
        {
            int right = row.Right - RightInset(r) - Px(RowBadgeEdge) - CheckReserve(r);
            if (!r.IsFolder)
            {
                // the triangles and the bars of the active sheet stand at the right end
                if (IsActiveRow(r) && _activeMarker == ActiveMarker.Triangle) right -= Px(MarkerRoom + 3);
                else if (IsActiveRow(r) && _activeMarker == ActiveMarker.EdgeBars) right -= Px(BarRoom + 3);
            }
            return right;
        }

        // The width the badges of a row take (the name stays clear of them).
        private int RowBadgeReserve(StripRow r)
        {
            if (_shownBadges == BadgeKinds.None) return 0;
            List<RowBadgeSpec> specs = RowBadges(r, new Rectangle(0, 0, ContentWidth, RowPx));
            if (specs.Count == 0) return 0;
            return specs.Sum(s => s.Bounds.Width + Px(3) + (s.Kind == RowBadgeKind.ProgressPill && specs.Count > 1 ? Px(RowBadgePillGap) : 0)) + Px(RowBadgeEdge - 2);
        }

        // ---- the copy icon ----
        // With the left button held down on a row, a press of the right button shows an icon under the pointer. When both buttons are
        // let go (with the pointer still on the row), the name of the sheet (the name Excel's own tab shows, with the sheet-name rule
        // "folder | sheet" too) or of the folder is copied to the clipboard, and nothing else happens: the sheet does not become the
        // active one, the selection is as it was, no menu opens. A pointer moved off the row before the buttons are let go cancels it.

        private const double CopyScale = 0.8; // (the icon is 0.8 times as large as when it was first made)
        private const double CopyChipWidth = 24 * CopyScale, CopyChipHeight = 22 * CopyScale, CopyIconSize = 16 * CopyScale;
        private bool _chordArmed;              // the left button is held down on a row
        private string _chordKey;              // that row ("S:name" / "F:id")
        private HashSet<string> _chordSelected, _chordSelFolders; // what the left press has changed on screen (taken back when the right button joins)
        private string _chordActive, _chordAnchor;
        private bool _chord;                   // both buttons are held: the icon shows
        private bool _chordOnRow;              // the pointer is still on the row
        private Point _chordAt;                // where the right button was pressed: the middle of the icon
        private string _copiedKey;             // the row whose name has just been copied (a check mark, for a second)
        private System.Windows.Forms.Timer _copiedTimer;

        private bool CopyIconApplies(StripRow r)
        {
            return r != null && !IsSample && _editor == null && !_dragging && !_ghostOn && _raClock == null && (r.IsFolder ? r.FolderName != null : r.Sheet != null);
        }

        private static string NameToCopy(StripRow r) { return r.IsFolder ? r.FolderName : r.Sheet; }

        // The left button goes down on a row: remember what it is about to change on screen.
        private void ArmChord(StripRow row)
        {
            _chordArmed = CopyIconApplies(row);
            if (!_chordArmed) return;
            _chordKey = RowKey(row);
            _chordSelected = new HashSet<string>(_selected);
            _chordSelFolders = new HashSet<string>(_selFolders);
            _chordActive = _active;
            _chordAnchor = _anchorFolder;
        }

        // The right button goes down while the left one is held on a row.
        private bool TryStartChord(Point p)
        {
            if (!_chordArmed || (Control.MouseButtons & MouseButtons.Left) == 0 || _dragging || _ghostOn) return false;
            int index = IndexFromPoint(p);
            if (index < 0 || RowKey(_rows[index]) != _chordKey || !CopyIconApplies(_rows[index])) return false;
            // what the left press has changed on screen is taken back: nothing is chosen, and Excel is told nothing
            _selected = _chordSelected;
            _selFolders = _chordSelFolders;
            _active = _chordActive;
            _anchorFolder = _chordAnchor;
            _selectionPending = false;
            _toggleOffFolder = null;
            _dragIndex = -1;
            ContextMenuStrip = null; // (the menu of the row must not open when the right button is let go)
            _chord = true;
            _chordOnRow = true;
            _chordAt = p;
            _copiedKey = null;
            _badgeTip = null;
            _tip.Hide(this);
            Invalidate();
            return true;
        }

        // A button is let go during the copy: when both are, the name is copied (or not, with the pointer off the row).
        private void ChordMouseUp()
        {
            if ((Control.MouseButtons & (MouseButtons.Left | MouseButtons.Right)) != 0) return;
            bool copy = _chordOnRow;
            _chord = false;
            _chordArmed = false;
            int index = _chordKey == null ? -1 : _rows.FindIndex(r => RowKey(r) == _chordKey);
            if (copy && index >= 0 && CopyIconApplies(_rows[index])) CopyRowName(index);
            else Invalidate();
        }

        // The pointer moves during the copy: off the row, the icon goes for good (and letting go copies nothing, even if the pointer comes back).
        private void UpdateChordPointer(Point p)
        {
            if (!_chord || !_chordOnRow) return; // (once off the row, the copy stays cancelled until the buttons are let go)
            int index = IndexFromPoint(p);
            bool on = index >= 0 && RowKey(_rows[index]) == _chordKey;
            if (on == _chordOnRow) return;
            _chordOnRow = on;
            Invalidate();
        }

        // The square of the icon: its middle where the right button was pressed, kept inside the row.
        private Rectangle CopyChipRect(int index)
        {
            Rectangle row = RowRect(index);
            int w = Px(CopyChipWidth), h = Px(CopyChipHeight);
            int x = Math.Max(row.Left + 2, Math.Min(row.Right - w - 2, _chordAt.X - w / 2));
            int y = Math.Max(row.Top + 1, Math.Min(row.Bottom - h - 1, _chordAt.Y - h / 2));
            return new Rectangle(x, y, w, h);
        }

        private void CopyRowName(int index)
        {
            StripRow r = _rows[index];
            string name = NameToCopy(r);
            Rectangle chip = CopyChipRect(index);
            int tipY = RowRect(index).Bottom;
            _tip.Hide(this);
            try
            {
                Clipboard.SetDataObject(name, true, 10, 50);
            }
            catch (Exception)
            {
                _badgeTip = "コピーできませんでした";
                _tip.Show(_badgeTip, this, chip.Left, tipY, 2500);
                return;
            }
            _copiedKey = RowKey(r);
            if (_copiedTimer == null)
            {
                _copiedTimer = new System.Windows.Forms.Timer { Interval = 1000 };
                _copiedTimer.Tick += (s, e) =>
                {
                    _copiedTimer.Stop();
                    _copiedKey = null; // (the icon has done its work)
                    _badgeTip = null;
                    _tip.Hide(this);
                    Invalidate();
                };
            }
            _copiedTimer.Stop();
            _copiedTimer.Start();
            _badgeTip = "コピーしました: " + name;
            _tip.Show(_badgeTip, this, chip.Left, tipY, 2500);
            Invalidate();
        }

        private void DrawCopyIcon(Graphics g, List<int> wobbling)
        {
            bool done = _copiedKey != null && _copiedKey == _chordKey;
            if (!done && !(_chord && _chordOnRow)) return;
            int index = _chordKey == null ? -1 : _rows.FindIndex(r => RowKey(r) == _chordKey);
            if (index < 0 || !CopyIconApplies(_rows[index]) || wobbling.Contains(index)) return;
            Rectangle chip = CopyChipRect(index);
            if (chip.Bottom < HeaderPx || chip.Top > ViewHeight) return;
            GraphicsState state = g.Save();
            g.SmoothingMode = SmoothingMode.AntiAlias;
            // the chip behind the icon: it shows on any color of a row
            using (GraphicsPath path = RoundedRectF(new RectangleF(chip.Left + 0.5f, chip.Top + 0.5f, chip.Width - 1f, chip.Height - 1f), Px(4 * CopyScale)))
            {
                using (var fill = new SolidBrush(Theme.PlusHoverBack)) g.FillPath(fill, path);
                using (var pen = new Pen(Theme.PlusColor)) g.DrawPath(pen, path);
            }
            float s = Px(CopyIconSize);
            float ix = chip.Left + (chip.Width - s) / 2f, iy = chip.Top + (chip.Height - s) / 2f;
            if (done)
            {
                using (var brush = new SolidBrush(BadgeProgressGreen)) g.FillEllipse(brush, ix - 0.5f, iy - 0.5f, s + 1f, s + 1f);
                using (var pen = new Pen(Color.White, Math.Max(1.6f, s * 0.12f)) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round })
                    g.DrawLines(pen, new[] { new PointF(ix + s * 0.26f, iy + s * 0.52f), new PointF(ix + s * 0.43f, iy + s * 0.69f), new PointF(ix + s * 0.75f, iy + s * 0.32f) });
            }
            else
            {
                float w = Math.Max(1.2f, s * 0.085f);
                using (var pen = new Pen(Theme.PlusColor, w) { LineJoin = LineJoin.Round, StartCap = LineCap.Round, EndCap = LineCap.Round })
                {
                    // the sheet behind (an outline), the sheet in front (with lines of text)
                    using (GraphicsPath back = RoundedRectF(new RectangleF(ix + s * 0.30f, iy + s * 0.08f, s * 0.52f, s * 0.58f), s * 0.09f)) g.DrawPath(pen, back);
                    using (GraphicsPath front = RoundedRectF(new RectangleF(ix + s * 0.16f, iy + s * 0.30f, s * 0.52f, s * 0.60f), s * 0.09f))
                    {
                        using (var fill = new SolidBrush(Theme.PlusHoverBack)) g.FillPath(fill, front);
                        g.DrawPath(pen, front);
                    }
                    pen.Width = Math.Max(1f, s * 0.05f);
                    foreach (float ly in new[] { 0.46f, 0.60f, 0.74f }) g.DrawLine(pen, ix + s * 0.26f, iy + s * ly, ix + s * 0.58f, iy + s * ly);
                }
            }
            g.Restore(state);
        }

        private bool IsProgressDone(string sheet)
        {
            BadgeInfo b;
            return sheet != null && (_shownBadges & BadgeKinds.Progress) != 0 && _rowBadges.TryGetValue(sheet, out b) && b.Progress == 3;
        }

        public int ActiveProgress()
        {
            BadgeInfo b;
            return _active != null && _rowBadges.TryGetValue(_active, out b) ? b.Progress : 0;
        }

        // Asks for a progress for sheets (0 = none, 1 = not started, 2 = in progress, 3 = done).
        public event Action<List<string>, int> ProgressRequested;

        // A click on the mark of a sheet that is done asks to protect it, or to take the protection off (Excel's own dialog), see PaneManager.
        public event Action<string> ProgressProtectRequested;

        // A click on the progress mark of a sheet: the next progress (none, not started, in progress, done); a click on "done" protects the sheet.
        private bool ProgressBadgeHit(int index, Point p)
        {
            if (index < 0 || index >= _rows.Count || _rows[index].IsFolder || (_shownBadges & BadgeKinds.Progress) == 0) return false;
            foreach (RowBadgeSpec s in RowBadges(_rows[index], RowRect(index)))
            {
                if (s.Kind == RowBadgeKind.Progress && Rectangle.Inflate(s.Bounds, Px(3), Px(5)).Contains(p)) return true;
            }
            return false;
        }

        // The paint can mark at the right end of a folder row that keeps its sheets in the folder's color (where DrawCheck draws it).
        private Rectangle PaintCanRect(int index)
        {
            if (index < 0 || index >= _rows.Count || !_rows[index].IsFolder || !_rows[index].SyncColor) return Rectangle.Empty;
            Rectangle row = RowRect(index);
            int right = row.Right - RightInset(_rows[index]);
            int width = Px(CheckWidth);
            int cx = right - Px(CheckPadRight) - width / 2;
            int top = row.Top + Px(PlainTop), bottom = row.Bottom - Px(BottomInset);
            int height = Px(CanRadius * 2);
            return new Rectangle(cx - width / 2, (top + bottom) / 2 - height / 2, width, height);
        }

        private bool PaintCanHit(int index, Point p)
        {
            Rectangle r = PaintCanRect(index);
            return !r.IsEmpty && Rectangle.Inflate(r, Px(3), Px(3)).Contains(p);
        }

        // A click on the paint can waits for the time of a double click: a second click within it is a double click (the color dialog),
        // else the palette opens.
        private System.Windows.Forms.Timer _canTimer;
        private string _canFolder;
        private Point _canScreen;

        private void PaintCanDown(int index, int clicks)
        {
            _tip.Hide(this); // (the tip of the mark would stay behind the palette; it is shown again only when the pointer comes back to the mark)
            string id = _rows[index].FolderId;
            if (clicks >= 2 && _canFolder == id && _canTimer != null && _canTimer.Enabled)
            {
                _canTimer.Stop();
                _canFolder = null;
                if (FolderColorMarkDoubleClicked != null) FolderColorMarkDoubleClicked(id);
                return;
            }
            Rectangle r = PaintCanRect(index);
            _canFolder = id;
            _canScreen = PointToScreen(new Point(r.Left, r.Bottom + Px(2)));
            if (_canTimer == null)
            {
                _canTimer = new System.Windows.Forms.Timer();
                _canTimer.Tick += (s, e) =>
                {
                    _canTimer.Stop();
                    string folder = _canFolder;
                    _canFolder = null;
                    if (folder != null && FolderColorMarkClicked != null) FolderColorMarkClicked(folder, _canScreen);
                };
            }
            _canTimer.Interval = Math.Max(100, SystemInformation.DoubleClickTime);
            _canTimer.Stop();
            _canTimer.Start();
        }

        private bool _protectAsking;

        private void CycleProgress(string sheet)
        {
            BadgeInfo b;
            int now = _rowBadges.TryGetValue(sheet, out b) ? b.Progress : 0;
            if (now == 3)
            {
                // "done": the dialog of Excel to protect the sheet. A sheet that is already protected stays so: the mark does not take the
                // protection off (that is done by "Unprotect Sheet" of the menu or of the ribbon)
                if (b == null || b.Protected || ProgressProtectRequested == null || _protectAsking) return;
                _protectAsking = true; // (the request opens Excel's dialog and does not return until it is closed: clicks that come meanwhile are ignored)
                try { ProgressProtectRequested(sheet); }
                finally { _protectAsking = false; }
                return;
            }
            if (ProgressRequested != null) ProgressRequested(new List<string> { sheet }, now + 1);
        }

        private void DrawRowBadges(Graphics g, StripRow r, Rectangle row)
        {
            List<RowBadgeSpec> specs = RowBadges(r, row);
            if (specs.Count == 0) return;
            DrawBadgeGlows(g, specs);
            foreach (RowBadgeSpec s in specs) DrawBadgeBody(g, s, false);
        }

        private void DrawBadgeBody(Graphics g, RowBadgeSpec s, bool outline)
        {
            if (s.Kind == RowBadgeKind.Progress) DrawProgressBadge(g, s.Bounds, s.State);
            else if (s.Kind == RowBadgeKind.ProgressPill) DrawProgressPill(g, s.Bounds, s.Count, s.Total);
            else if (s.Kind == RowBadgeKind.Error) DrawErrorBadge(g, s.Bounds, s.Color, outline);
            else if (s.Kind == RowBadgeKind.Memo) DrawMemoBadge(g, s.Bounds, s.Color, outline);
            else if (s.Kind == RowBadgeKind.Protect) DrawLockBadge(g, s.Bounds, s.Color, outline);
            else DrawCountBadge(g, s.Bounds, s.Color, s.Count, outline);
        }

        // The white rim and the soft glow around every badge: the badges stand out on a dark row. The glows of the badges of a
        // row are added up first (where they overlap they sum, up to white), the badges are drawn on top.
        private sealed class GlowMask { public byte[] A; public int W, H, Pad; }

        private static readonly Dictionary<string, GlowMask> GlowCache = new Dictionary<string, GlowMask>();
        private const float GlowRim = 0.8f, GlowBlur = 2.0f;

        private void DrawBadgeGlows(Graphics g, List<RowBadgeSpec> specs)
        {
            try
            {
                var masks = new List<GlowMask>();
                int l = int.MaxValue, t = int.MaxValue, rt = int.MinValue, bt = int.MinValue;
                for (int i = 0; i < specs.Count; i++)
                {
                    GlowMask m = GlowFor(specs[i]);
                    masks.Add(m);
                    int x = specs[i].Bounds.Left - m.Pad, y = specs[i].Bounds.Top - m.Pad;
                    l = Math.Min(l, x);
                    t = Math.Min(t, y);
                    rt = Math.Max(rt, x + m.W);
                    bt = Math.Max(bt, y + m.H);
                }
                int w = rt - l, h = bt - t;
                if (w <= 0 || h <= 0) return;
                var acc = new int[w * h];
                for (int i = 0; i < specs.Count; i++)
                {
                    GlowMask m = masks[i];
                    int ox = specs[i].Bounds.Left - m.Pad - l, oy = specs[i].Bounds.Top - m.Pad - t;
                    for (int y = 0; y < m.H; y++)
                        for (int x = 0; x < m.W; x++)
                            acc[(oy + y) * w + ox + x] += m.A[y * m.W + x];
                }
                for (int i = 0; i < acc.Length; i++) acc[i] = (Math.Min(255, acc[i]) << 24) | 0xFFFFFF;
                using (var bmp = new Bitmap(w, h, PixelFormat.Format32bppArgb))
                {
                    BitmapData data = bmp.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, PixelFormat.Format32bppArgb);
                    try { Marshal.Copy(acc, 0, data.Scan0, acc.Length); }
                    finally { bmp.UnlockBits(data); }
                    g.DrawImage(bmp, new Rectangle(l, t, w, h));
                }
            }
            catch
            {
                // no glow is better than no badge
            }
        }

        private GlowMask GlowFor(RowBadgeSpec s)
        {
            int bw = s.Bounds.Width, bh = s.Bounds.Height;
            string key = s.Kind + "|" + s.State + "|" + bw + "|" + bh + "|" + Px(100);
            GlowMask cached;
            if (GlowCache.TryGetValue(key, out cached)) return cached;
            float sc = Px(100) / 125f;
            float rim = GlowRim * sc, grow = rim + 1.2f * sc, sigma = Math.Max(0.6f, GlowBlur * sc * 0.5f);
            int pad = (int)Math.Ceiling(grow + 3 * sigma) + 1;
            int W = bw + 2 * pad, H = bh + 2 * pad;
            var alpha = new byte[W * H];
            using (var bmp = new Bitmap(W, H, PixelFormat.Format32bppArgb))
            {
                using (Graphics gg = Graphics.FromImage(bmp))
                {
                    gg.Clear(Color.Transparent);
                    var body = new RowBadgeSpec { Kind = s.Kind, State = s.State, Color = s.Color, Count = s.Count, Total = s.Total, Bounds = new Rectangle(pad, pad, bw, bh) };
                    DrawBadgeBody(gg, body, false);
                }
                BitmapData data = bmp.LockBits(new Rectangle(0, 0, W, H), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    var raw = new byte[Math.Abs(data.Stride) * H];
                    Marshal.Copy(data.Scan0, raw, 0, raw.Length);
                    for (int y = 0; y < H; y++)
                        for (int x = 0; x < W; x++) alpha[y * W + x] = raw[y * data.Stride + x * 4 + 3];
                }
                finally { bmp.UnlockBits(data); }
            }
            // the hole inside a shape (the shackle of a padlock) stays open: what the outside cannot reach
            var outside = new bool[W * H];
            var stack = new Stack<int>();
            stack.Push(0);
            outside[0] = true;
            while (stack.Count > 0)
            {
                int p = stack.Pop(), px = p % W, py = p / W;
                if (px > 0 && !outside[p - 1] && alpha[p - 1] <= 16) { outside[p - 1] = true; stack.Push(p - 1); }
                if (px < W - 1 && !outside[p + 1] && alpha[p + 1] <= 16) { outside[p + 1] = true; stack.Push(p + 1); }
                if (py > 0 && !outside[p - W] && alpha[p - W] <= 16) { outside[p - W] = true; stack.Push(p - W); }
                if (py < H - 1 && !outside[p + W] && alpha[p + W] <= 16) { outside[p + W] = true; stack.Push(p + W); }
            }
            float[] rimA = Dilate(alpha, W, H, rim);
            float[] halo = Blur(Dilate(alpha, W, H, grow), W, H, sigma);
            var result = new byte[W * H];
            for (int i = 0; i < result.Length; i++)
            {
                bool hole = alpha[i] <= 16 && !outside[i];
                float v = hole ? 0 : Math.Max(rimA[i], halo[i] * 0.6f);
                result[i] = (byte)Math.Max(0, Math.Min(255, (int)Math.Round(v)));
            }
            if (GlowCache.Count > 80) GlowCache.Clear();
            var mask = new GlowMask { A = result, W = W, H = H, Pad = pad };
            GlowCache[key] = mask;
            return mask;
        }

        // the strongest alpha within a radius (a soft edge of one pixel)
        private static float[] Dilate(byte[] a, int W, int H, float radius)
        {
            var o = new float[W * H];
            int n = (int)Math.Ceiling(radius) + 1;
            for (int y = 0; y < H; y++)
                for (int x = 0; x < W; x++)
                {
                    float best = 0;
                    for (int dy = -n; dy <= n; dy++)
                    {
                        int yy = y + dy;
                        if (yy < 0 || yy >= H) continue;
                        for (int dx = -n; dx <= n; dx++)
                        {
                            int xx = x + dx;
                            if (xx < 0 || xx >= W) continue;
                            float d = (float)Math.Sqrt(dx * dx + dy * dy);
                            float f = d <= radius ? 1f : d < radius + 1 ? radius + 1 - d : 0f;
                            if (f <= 0) continue;
                            float v = a[yy * W + xx] * f;
                            if (v > best) best = v;
                        }
                    }
                    o[y * W + x] = best;
                }
            return o;
        }

        private static float[] Blur(float[] src, int W, int H, float sigma)
        {
            int rad = (int)Math.Ceiling(sigma * 3);
            var k = new float[rad * 2 + 1];
            float sum = 0;
            for (int i = -rad; i <= rad; i++) { k[i + rad] = (float)Math.Exp(-(i * i) / (2 * sigma * sigma)); sum += k[i + rad]; }
            for (int i = 0; i < k.Length; i++) k[i] /= sum;
            var tmp = new float[W * H];
            var o = new float[W * H];
            for (int y = 0; y < H; y++)
                for (int x = 0; x < W; x++)
                {
                    float v = 0;
                    for (int i = -rad; i <= rad; i++) { int xx = x + i; if (xx >= 0 && xx < W) v += src[y * W + xx] * k[i + rad]; }
                    tmp[y * W + x] = v;
                }
            for (int y = 0; y < H; y++)
                for (int x = 0; x < W; x++)
                {
                    float v = 0;
                    for (int i = -rad; i <= rad; i++) { int yy = y + i; if (yy >= 0 && yy < H) v += tmp[yy * W + x] * k[i + rad]; }
                    o[y * W + x] = v;
                }
            return o;
        }

        // the progress of a sheet: an empty ring (not started), a half filled ring (in progress), a green circle with a check (done)
        private static void DrawProgressBadge(Graphics g, Rectangle b, int state)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = b.Left + b.Width / 2f, cy = b.Top + b.Height / 2f, r = b.Width / 2f;
            float line = Math.Max(1.4f, r * 0.2f);
            if (state == 4)
            {
                // done and protected: a green padlock whose body carries the check
                float k = r * 0.92f;
                float bw = k * 1.7f, bh = k * 1.25f, bt = cy - k * 0.2f;
                float sw = k * 0.95f, w = Math.Max(1.6f, k * 0.34f);
                using (var fill = new SolidBrush(BadgeProgressGreen))
                {
                    float ymid = cy - k * 1.0f + w / 2 + sw / 2;
                    // (one filled shape, the same on both sides of the middle)
                    using (GraphicsPath shackle = ShacklePath(cx, ymid, sw / 2, sw / 2, w, bt + k * 0.2f)) g.FillPath(fill, shackle);
                    using (GraphicsPath body = RoundedRectF(new RectangleF(cx - bw / 2, bt, bw, bh), Math.Max(2f, k * 0.28f)))
                    {
                        g.FillPath(fill, body);
                    }
                }
                float cy2 = bt + bh / 2, r2 = k * 0.7f;
                using (var pen = new Pen(Color.White, Math.Max(1.5f, r2 * 0.26f)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                {
                    g.DrawLines(pen, new[] { new PointF(cx - r2 * 0.45f, cy2 + r2 * 0.02f), new PointF(cx - r2 * 0.1f, cy2 + r2 * 0.36f), new PointF(cx + r2 * 0.5f, cy2 - r2 * 0.34f) });
                }
                g.SmoothingMode = mode;
                return;
            }
            if (state == 3)
            {
                using (var fill = new SolidBrush(BadgeProgressGreen)) g.FillEllipse(fill, cx - r, cy - r, r * 2, r * 2);
                using (var pen = new Pen(Color.White, Math.Max(1.6f, r * 0.22f)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                {
                    g.DrawLines(pen, new[] { new PointF(cx - r * 0.45f, cy + r * 0.02f), new PointF(cx - r * 0.1f, cy + r * 0.36f), new PointF(cx + r * 0.5f, cy - r * 0.34f) });
                }
            }
            else
            {
                Color c = state == 2 ? BadgeProgressBlue : BadgeProgressGray;
                float rr = r - line / 2;
                using (var white = new SolidBrush(Color.White)) g.FillEllipse(white, cx - rr, cy - rr, rr * 2, rr * 2);
                using (var pen = new Pen(c, line)) g.DrawEllipse(pen, cx - rr, cy - rr, rr * 2, rr * 2);
                if (state == 2)
                {
                    float ri = r * 0.6f;
                    using (var fill = new SolidBrush(c)) g.FillPie(fill, cx - ri, cy - ri, ri * 2, ri * 2, -90, 180);
                }
            }
            g.SmoothingMode = mode;
        }

        // the pill of a folder: the sheets that are done out of those that have a progress; green with a check when all of them are done
        private void DrawProgressPill(Graphics g, Rectangle b, int done, int total)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (GraphicsPath path = RoundedRect(b, b.Height / 2))
            {
                if (done == total)
                {
                    using (var fill = new SolidBrush(BadgeProgressGreen)) g.FillPath(fill, path);
                    float cx = b.Left + b.Width / 2f, cy = b.Top + b.Height / 2f;
                    using (var pen = new Pen(Color.White, Math.Max(1.6f, b.Height * 0.14f)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
                    {
                        g.DrawLines(pen, new[] { new PointF(cx - b.Width * 0.14f, cy), new PointF(cx - b.Width * 0.04f, cy + b.Height * 0.2f), new PointF(cx + b.Width * 0.14f, cy - b.Height * 0.2f) });
                    }
                }
                else
                {
                    using (var fill = new SolidBrush(Color.White)) g.FillPath(fill, path);
                    using (var pen = new Pen(Color.FromArgb(185, 185, 190), 1f)) g.DrawPath(pen, path);
                    using (var font = new Font(TabFont.FontFamily, Math.Max(7f, b.Height * 0.62f), FontStyle.Bold, GraphicsUnit.Pixel))
                    using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                    using (var brush = new SolidBrush(Color.FromArgb(70, 70, 75)))
                    {
                        g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                        g.DrawString(done + "/" + total, font, brush, new RectangleF(b.Left, b.Top + 1, b.Width, b.Height), format);
                    }
                }
            }
            g.SmoothingMode = mode;
        }

        // a red triangle with an exclamation mark, in a white rim
        private static void DrawErrorBadge(Graphics g, Rectangle b, Color color, bool outline = true)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = b.Left + b.Width / 2f, cy = b.Top + b.Height / 2f, r = b.Width / 2f * 1.05f;
            if (outline)
            {
                using (var white = new SolidBrush(Color.White)) g.FillPolygon(white, TriangleAt(cx, cy, r, 1f));
            }
            using (var fill = new SolidBrush(color)) g.FillPolygon(fill, TriangleAt(cx, cy, r, outline ? 0.82f : 0.92f));
            using (var white = new SolidBrush(Color.White))
            {
                float w = Math.Max(1.5f, r * 0.2f);
                g.FillRectangle(white, cx - w / 2, cy - r * 0.35f, w, r * 0.55f);
                g.FillEllipse(white, cx - w / 2, cy + r * 0.38f, w, w);
            }
            g.SmoothingMode = mode;
        }

        private static PointF[] TriangleAt(float cx, float cy, float r, float k)
        {
            return new[] { new PointF(cx, cy - r * k), new PointF(cx + r * 1.1f * k, cy + r * 0.85f * k), new PointF(cx - r * 1.1f * k, cy + r * 0.85f * k) };
        }

        // an amber note with three lines, in a white rim
        private static void DrawMemoBadge(Graphics g, Rectangle b, Color color, bool outline = true)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = b.Left + b.Width / 2f, cy = b.Top + b.Height / 2f, r = b.Width / 2f * 1.05f;
            var outer = Rectangle.Round(new RectangleF(cx - r * 0.85f, cy - r * 0.95f, r * 1.7f, r * 1.9f));
            var inner = outline ? new RectangleF(cx - r * 0.75f, cy - r * 0.84f, r * 1.5f, r * 1.68f) : new RectangleF(cx - r * 0.8f, cy - r * 0.9f, r * 1.6f, r * 1.8f);
            if (outline)
            {
                using (GraphicsPath p = RoundedRect(outer, Math.Max(2, (int)(r * 0.25f))))
                using (var white = new SolidBrush(Color.White)) g.FillPath(white, p);
            }
            using (GraphicsPath p = RoundedRectF(inner, Math.Max(2f, r * 0.2f)))
            using (var fill = new SolidBrush(color)) g.FillPath(fill, p);
            using (var pen = new Pen(Color.White, Math.Max(1.5f, r * 0.2f)))
            {
                for (int i = 0; i < 3; i++)
                {
                    float y = cy - r * 0.5f + i * r * 0.4f;
                    g.DrawLine(pen, cx - r * 0.5f, y, i < 2 ? cx + r * 0.25f : cx - r * 0.1f, y);
                }
            }
            g.SmoothingMode = mode;
        }

        // a padlock (a body and a shackle), in a white rim
        private static void DrawLockBadge(Graphics g, Rectangle b, Color color, bool outline = true)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            float cx = b.Left + b.Width / 2f, cy = b.Top + b.Height / 2f, r = b.Width / 2f * 1.05f;
            float bodyTop = cy - r * 0.12f, bodyH = r * 1.0f, bodyW = r * 1.5f;
            float shackleW = r * 0.95f, shackleTop = cy - r * 0.95f;
            // the white rim: a thicker shackle and a larger body
            float rim = Math.Max(1.2f, r * 0.2f);
            using (var rimPen = new Pen(Color.White, r * 0.34f + rim))
            using (var pen = new Pen(color, r * 0.34f))
            {
                rimPen.StartCap = rimPen.EndCap = LineCap.Round;
                if (outline)
                {
                    g.DrawArc(rimPen, cx - shackleW / 2, shackleTop, shackleW, r * 1.4f, 180, 180);
                    g.DrawLine(rimPen, cx - shackleW / 2, shackleTop + r * 0.7f, cx - shackleW / 2, bodyTop + r * 0.1f);
                    g.DrawLine(rimPen, cx + shackleW / 2, shackleTop + r * 0.7f, cx + shackleW / 2, bodyTop + r * 0.1f);
                    var outer = Rectangle.Round(new RectangleF(cx - bodyW / 2 - rim * 0.6f, bodyTop - rim * 0.6f, bodyW + rim * 1.2f, bodyH + rim * 1.2f));
                    using (GraphicsPath p = RoundedRect(outer, Math.Max(2, (int)(r * 0.3f)))) using (var white = new SolidBrush(Color.White)) g.FillPath(white, p);
                }
            }
            // the shackle is one filled shape that is built the same on both sides of the middle (pens with caps were not the same on both)
            using (GraphicsPath shackle = ShacklePath(cx, shackleTop + r * 0.7f, shackleW / 2, r * 0.7f, r * 0.34f, bodyTop + r * 0.15f))
            using (var shackleFill = new SolidBrush(color))
                g.FillPath(shackleFill, shackle);
            var inner = new RectangleF(cx - bodyW / 2, bodyTop, bodyW, bodyH);
            using (GraphicsPath p = RoundedRectF(inner, Math.Max(2f, r * 0.22f))) using (var fill = new SolidBrush(color)) g.FillPath(fill, p);
            using (var white = new SolidBrush(Color.White))
            {
                float k = Math.Max(1.6f, r * 0.26f);
                g.FillEllipse(white, cx - k / 2, bodyTop + bodyH * 0.28f, k, k);
                g.FillRectangle(white, cx - k * 0.2f, bodyTop + bodyH * 0.28f + k * 0.5f, k * 0.4f, bodyH * 0.3f);
            }
            g.SmoothingMode = mode;
        }

        // The icon of a kind of badge (for the settings dialog).
        public static void DrawBadgeIcon(Graphics g, BadgeKinds kind, Rectangle bounds)
        {
            if (kind == BadgeKinds.Error) DrawErrorBadge(g, bounds, BadgeErrorColor);
            else if (kind == BadgeKinds.Memo) DrawMemoBadge(g, bounds, BadgeMemoColor);
            else if (kind == BadgeKinds.Protect) DrawLockBadge(g, bounds, BadgeProtectColor);
            else if (kind == BadgeKinds.Progress) DrawProgressBadge(g, bounds, 2);
        }

        // a round badge with the number of the sheets of a closed folder that have something to tell
        private void DrawCountBadge(Graphics g, Rectangle b, Color color, int count, bool outline = true)
        {
            var mode = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            if (outline)
            {
                using (var white = new SolidBrush(Color.White)) g.FillEllipse(white, b);
            }
            using (var fill = new SolidBrush(color)) g.FillEllipse(fill, outline ? Rectangle.Inflate(b, -Math.Max(1, Px(1)), -Math.Max(1, Px(1))) : b);
            using (var font = new Font(TabFont.FontFamily, Math.Max(7f, b.Height * 0.52f), FontStyle.Bold, GraphicsUnit.Pixel))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
            using (var brush = new SolidBrush(Color.White))
            {
                g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                g.DrawString(count > 99 ? "99+" : count.ToString(), font, brush, new RectangleF(b.Left - 2, b.Top + 1, b.Width + 4, b.Height), format);
            }
            g.SmoothingMode = mode;
        }

        // The text of the tip of the badge under the pointer (null = none).
        private string BadgeTipAt(int index, Point p)
        {
            if (PaintCanHit(index, p)) return "フォルダと中のシートを同一色に設定 (クリック: フォルダの色、ダブルクリック: その他の色)";
            if (index < 0 || index >= _rows.Count || _shownBadges == BadgeKinds.None) return null;
            foreach (RowBadgeSpec s in RowBadges(_rows[index], RowRect(index)))
            {
                if (Rectangle.Inflate(s.Bounds, Px(2), Px(4)).Contains(p)) return s.Tip;
            }
            return null;
        }

        // The mark of the active sheet, with every look.
        private ActiveMarker _activeMarker = ActiveMarker.Underline;

        public ActiveMarker ActiveMarker
        {
            get { return _activeMarker; }
            set { if (_activeMarker != value) { _activeMarker = value; Invalidate(); } }
        }

        // The wobble of the sheet that has just been selected: a timer repaints the strip while it runs, and the row of that
        // sheet is drawn last (so it lies over its neighbors) through a move, a turn and a scale.
        private WobbleType _wobbleType = WobbleType.None;
        private int _wobbleAmp = WobbleMotion.AmpDefault;
        private int _wobbleDur = WobbleMotion.DurDefault;
        private WobbleType _animType;
        private string _animSheet;
        private HashSet<string> _animFolders;
        private HashSet<string> _animSheets;
        private Dictionary<string, bool> _folderOpen = new Dictionary<string, bool>();
        private HashSet<string> _folderSel = new HashSet<string>();
        private bool _folderSelSeen;
        private System.Diagnostics.Stopwatch _animClock;
        private System.Windows.Forms.Timer _animTimer;
        private string _lastActive;
        private bool _activeSeen;
        private bool _transforming;

        private bool _soundOn;
        private int _soundVolume = WobbleSounds.VolumeDefault;
        private bool _muteNext;

        // The sound effect of the wobble.
        public void SetWobbleSound(bool on, int volume)
        {
            _soundOn = on;
            _soundVolume = Math.Max(0, Math.Min(100, volume));
        }

        public void SetWobble(WobbleType type, int amp, int dur)
        {
            _wobbleType = type;
            _wobbleAmp = Math.Max(WobbleMotion.AmpMin, Math.Min(WobbleMotion.AmpMax, amp));
            _wobbleDur = Math.Max(WobbleMotion.DurMin, Math.Min(WobbleMotion.DurMax, dur));
            if (type == WobbleType.None) StopWobble();
        }

        // Plays the wobble of the active sheet (a new random motion each time for WobbleType.Random).
        public void PlayWobble()
        {
            PlayWobble(true);
        }

        public void PlayWobble(bool withSound)
        {
            _muteNext = !withSound;
            try { if (_active != null) StartWobble(_active, null, null); }
            finally { _muteNext = false; }
        }

        // Starts the wobble of one sheet and/or some folders (all with the same motion).
        private void StartWobble(string sheet, HashSet<string> folders, HashSet<string> sheets)
        {
            if (_wobbleType == WobbleType.None || !IsHandleCreated) return;
            _animType = _wobbleType == WobbleType.Random ? WobbleMotion.PickRandom() : _wobbleType;
            if (_soundOn && !_muteNext) WobbleSounds.Play(_animType, _soundVolume);
            _animSheet = sheet;
            _animFolders = folders;
            _animSheets = sheets;
            _animClock = System.Diagnostics.Stopwatch.StartNew();
            if (_animTimer == null)
            {
                _animTimer = new System.Windows.Forms.Timer { Interval = 15 };
                _animTimer.Tick += (s, e) =>
                {
                    if (_animClock == null || _animClock.ElapsedMilliseconds >= _wobbleDur) StopWobble();
                    Invalidate();
                };
            }
            _animTimer.Start();
            Invalidate();
        }

        private void StopWobble()
        {
            if (_animTimer != null) _animTimer.Stop();
            _animClock = null;
            _animSheet = null;
            _animFolders = null;
            _animSheets = null;
        }

        private bool Glass { get { return _selectStyle == SelectStyle.FrostedGlass || _selectStyle == SelectStyle.GlossyGlass || _selectStyle == SelectStyle.FrostedGlass2; } }

        // A sample (in the settings dialog): the same drawing, three fixed rows, no "+" button, no operation.
        public bool IsSample { get; private set; }

        // The size of a sample of the given width (in DIP): three rows.
        public Size SampleSize(int widthDip)
        {
            return new Size(Px(widthDip), 3 * RowPx + Px(2));
        }

        // The size of a one-row sample of the given width (in DIP).
        public Size SingleSampleSize(int widthDip)
        {
            return new Size(Px(widthDip), RowPx + Px(2));
        }

        // A click on a sample.
        public event Action SampleClicked;

        // One row, the active sheet in orange with the given name: the sample of the wobble.
        public void SetSingleSample(SelectStyle style, ActiveMarker marker, string name)
        {
            _activeMarker = marker;
            IsSample = true;
            TabStop = false;
            _selectStyle = style;
            var rows = new List<StripRow> { new StripRow { Sheet = name } };
            var colors = new Dictionary<string, int> { { name, 0x3A80E9 } };
            SetSheets(new List<string> { name }, name, new List<string> { name }, colors, rows);
            _selFolders.Clear();
            Invalidate();
        }

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
            _rowBadges = new Dictionary<string, BadgeInfo> { { "Sheet1", new BadgeInfo { Errors = 2, Protected = true, Progress = 2 } }, { "Sheet2", new BadgeInfo { Memos = 1, Progress = 3 } } };
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

        // Marks must stand out from the row they are on: when the color of a mark is too near to the color
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
        // The place a dragged row comes from: the row, mostly covered with the background.
        private void DrawFadedPlace(Graphics g, Rectangle row)
        {
            var area = new Rectangle(row.Left + Px(InsetLeft) - Px(3), row.Top + Px(PlainTop) - 1, row.Width - Px(InsetLeft) * 2 + Px(6), row.Height - Px(PlainTop) - Px(BottomInset) + 2);
            using (GraphicsPath path = RoundedRect(area, Px(8)))
            using (var brush = new SolidBrush(Color.FromArgb(185, Theme.BarBack))) g.FillPath(brush, path);
        }

        private void DrawDropMarker(Graphics g, int y, int x)
        {
            int w = Px(MarkerWidth), h = Px(MarkerHeight);
            var points = new[] { new Point(x, y - h / 2), new Point(x + w, y), new Point(x, y + h / 2) };
            using (var brush = new SolidBrush(Theme.DropMarker)) g.FillPolygon(brush, points);
        }

        private void DrawAllButtons(Graphics g)
        {
            using (var font = new Font(TabFont.FontFamily, TabFont.Size * 0.8f, FontStyle.Regular, TabFont.Unit))
            using (var format = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoWrap })
            {
                g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                for (int k = 1; k <= 2; k++)
                {
                    Rectangle box = AllButtonRect(k == 1);
                    bool hot = _headerHot == k, down = _headerDown == k;
                    using (GraphicsPath path = RoundedRect(box, Px(HeaderRadius)))
                    {
                        using (var brush = new SolidBrush(down ? Theme.TabBorder : Theme.PlusHoverBack)) g.FillPath(brush, path);
                        using (var pen = new Pen(hot || down ? Theme.PlusColor : Theme.Separator)) g.DrawPath(pen, path);
                    }
                    using (var brush = new SolidBrush(Theme.PlusColor))
                    {
                        g.DrawString(k == 1 ? "全開" : "全閉", font, brush, new RectangleF(box.Left, box.Top + Px(HeaderTextShift), box.Width, box.Height), format);
                    }
                }
            }
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
            else if (selected && Crystal)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawCrystal(g, body, colored, color, 1.0);
            }
            else if (selected && Glass)
            {
                if (colored) FillColoredBox(g, left, right, row, radius, color);
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                DrawRowBadges(g, folder, row); // (under the glass)
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
            var textRect = new Rectangle(textLeft, row.Top - (selected && DoubleFrame ? Px(ActiveTextLift) : 0) + TextShift(folder.FolderName), Math.Max(1, right - textLeft - Px(TextPadRight) - roomRight - CheckReserve(folder) - RowBadgeReserve(folder)), row.Height);
            g.TextRenderingHint = _transforming ? TextRenderingHint.AntiAliasGridFit : TextRenderingHint.ClearTypeGridFit;
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
                    DrawMatch(g, folder.FolderName, bold, textRect, format, true);
                    g.DrawString(folder.FolderName, bold, brush, textRect, format);
                }
            }
            if (folder.SyncColor) DrawCheck(g, row, right);
            if (!(selected && Glass)) DrawRowBadges(g, folder, row);
            DrawSeparator(g, row);
        }

        // The mark at the right end of a folder row that keeps its sheets in the folder's color: a paint can
        // tipped to the left (silver metal, a white label band) and a blue drop of paint falling from its lip, over a floor
        // line. It has a soft light halo (a few pixels wide, white pens of growing width with a little opacity each, which
        // looks like a blur), so that it can be seen on the dark bar of the dark Office themes and on any folder color alike.
        private void DrawCheck(Graphics g, Rectangle row, int right)
        {
            float r = Px(CanRadius);
            float cx = right - Px(CheckPadRight) - Px(CheckWidth) / 2f;
            int top = row.Top + Px(PlainTop), bottom = row.Bottom - Px(BottomInset);
            float cy = (top + bottom) / 2f - r * 0.02f;
            // the can, tipped to the left (the points are mirrored at the vertical line through cx)
            PointF[] can =
            {
                new PointF(cx + r * 0.8f, cy - r * 0.1f), new PointF(cx, cy - r * 0.9f),
                new PointF(cx - r * 0.6f, cy - r * 0.3f), new PointF(cx + r * 0.2f, cy + r * 0.5f)
            };
            PointF lip = new PointF(cx - r * 0.78f, cy + r * 0.05f);
            float k = r * 0.7f;
            RectangleF floor = new RectangleF(cx - r * 0.95f, cy + r * 0.78f, r * 1.9f, r * 0.15f);
            using (var shape = new GraphicsPath())
            {
                shape.AddPolygon(can);
                shape.AddPolygon(new[] { new PointF(lip.X, lip.Y - k * 0.2f), new PointF(lip.X + k * 0.3f, lip.Y + k * 0.35f), new PointF(lip.X - k * 0.3f, lip.Y + k * 0.35f) });
                shape.AddEllipse(lip.X - k * 0.3f, lip.Y + k * 0.08f, k * 0.6f, k * 0.6f);
                shape.AddRectangle(floor);
                float halo = Math.Max(2f, Px(CheckHalo));
                const int steps = 6;
                for (int i = steps; i >= 1; i--)
                {
                    using (var pen = new Pen(Color.FromArgb(46, 255, 255, 255), 2 * halo * i / steps))
                    {
                        pen.LineJoin = LineJoin.Round;
                        g.DrawPath(pen, shape);
                    }
                }
            }

            // the body: silver metal shaded across the can, then a white label band
            var metal = Color.FromArgb(150, 152, 158);
            float[] at = { 0f, 0.14f, 0.2f, 0.5f, 0.78f, 1f };
            using (var body = new GraphicsPath())
            {
                body.AddPolygon(can);
                var colors = new Color[at.Length];
                for (int i = 0; i < at.Length; i++)
                {
                    double t = at[i];
                    double kk = 0.62 + 0.62 * Math.Sin(t * Math.PI * 0.95 + 0.1) + (t > 0.12 && t < 0.26 ? 0.2 : 0);
                    colors[i] = Color.FromArgb(Math.Min(255, (int)(metal.R * kk)), Math.Min(255, (int)(metal.G * kk)), Math.Min(255, (int)(metal.B * kk)));
                }
                using (var brush = new LinearGradientBrush(can[0], can[3], colors[0], colors[colors.Length - 1]))
                {
                    brush.InterpolationColors = new ColorBlend { Colors = colors, Positions = at };
                    g.FillPath(brush, body);
                }
                var state = g.Save();
                g.SetClip(body, CombineMode.Intersect);
                PointF L(PointF p, PointF q, float t) { return new PointF(p.X + (q.X - p.X) * t, p.Y + (q.Y - p.Y) * t); }
                var band = new[] { L(can[1], can[0], 0.38f), L(can[2], can[3], 0.38f), L(can[2], can[3], 0.62f), L(can[1], can[0], 0.62f) };
                using (var bandBrush = new SolidBrush(Color.FromArgb(230, 238, 238, 240))) g.FillPolygon(bandBrush, band);
                g.Restore(state);
                using (var pen = new Pen(Color.FromArgb(70, 72, 78), Math.Max(1f, r * 0.05f)) { LineJoin = LineJoin.Round })
                    g.DrawPolygon(pen, can);
                using (var pen = new Pen(Color.FromArgb(225, 227, 232), Math.Max(1.5f, r * 0.1f)))
                    g.DrawLine(pen, can[1], can[2]);
            }

            // the drop of paint (blue) and the floor line
            using (var blue = new SolidBrush(Color.FromArgb(40, 110, 230)))
            {
                g.FillPolygon(blue, new[] { new PointF(lip.X, lip.Y - k * 0.2f), new PointF(lip.X + k * 0.3f, lip.Y + k * 0.35f), new PointF(lip.X - k * 0.3f, lip.Y + k * 0.35f) });
                g.FillEllipse(blue, lip.X - k * 0.3f, lip.Y + k * 0.08f, k * 0.6f, k * 0.6f);
            }
            using (var shine = new SolidBrush(Color.FromArgb(120, 255, 255, 255)))
                g.FillEllipse(shine, lip.X - k * 0.17f, lip.Y + k * 0.22f, k * 0.12f * 1.5f, k * 0.16f * 1.5f);
            using (var line = new SolidBrush(Color.FromArgb(95, 97, 103)))
                g.FillRectangle(line, floor);
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
            else if (selected && Crystal)
            {
                var body = new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset));
                textColor = DrawCrystal(g, body, colored, tabColor, active ? 1.0 : 0.6);
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
                DrawRowBadges(g, tabRow, row); // (the glass covers the badges of the row as well)
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

            // the badges at the right end, over the look of the row (on glass they were drawn above, under the glass)
            if (!(selected && Glass)) DrawRowBadges(g, tabRow, row);

            // the check badge marker, over whichever look the row has
            if (badge && !(selected && Glass)) DrawCheckBadge(g, new Rectangle(left, row.Top + Px(PlainTop), right - left, row.Height - Px(PlainTop) - Px(BottomInset)), colored, tabColor);

            // the separator between tabs: Excel's thin line, turned to lie across the stack
            DrawSeparator(g, row);

            // text: left aligned; the text of a selected tab sits a little higher (the lines above and below)
            // an underline needs room below the text, so the text sits higher
            int lift = underline ? Px(StandardTextLift) : DoubleFrame && selected ? Px(ActiveTextLift) : 0;
            int room = triangles ? Px(MarkerRoom) : bars ? Px(BarRoom) : 0; // room for the triangles or the bars at both ends of the active sheet
            int roomLeft = badge ? Px(BadgeRoom) : room; // (the check badge takes the left end; it is the only mark then)
            if (IsProgressDone(name)) textColor = Mix(textColor, Theme.BarBack, 0.5); // (a finished sheet is a little paler)
            var textRect = new Rectangle(left + roomLeft + Px(TextPadLeft), row.Top - lift + TextShift(text), Math.Max(1, right - left - roomLeft - room - Px(TextPadLeft) - Px(TextPadRight) - RowBadgeReserve(tabRow)), row.Height);
            if (_editor == null || name != _editingSheet)
            {
                g.TextRenderingHint = _transforming ? TextRenderingHint.AntiAliasGridFit : TextRenderingHint.ClearTypeGridFit;
                using (var font = selected && (DoubleFrame || Glass || Inset || Stripes || Crystal) ? new Font(TabFont, FontStyle.Bold) : null)
                using (var brush = new SolidBrush(textColor))
                using (var format = new StringFormat(StringFormat.GenericDefault))
                {
                    format.Alignment = StringAlignment.Near;
                    format.LineAlignment = StringAlignment.Center;
                    format.Trimming = StringTrimming.EllipsisCharacter;
                    format.FormatFlags = StringFormatFlags.NoWrap | StringFormatFlags.NoClip;
                    DrawMatch(g, text, font ?? TabFont, textRect, format, false);
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

        // The inset type: the row looks pressed in. The body is a little darker than the bar
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

        // The stripes look: a white body (a colored row: its color) with fine diagonal stripes
        // (green on white, white on a dark color, dark on a light one) and an outline in the green (a colored row: a dark shade of its color).
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
                // white stripes on a dark color, dark ones (the color, 55% darker) on a light color
                Color stripe = !colored ? Color.FromArgb(55, Underline.R, Underline.G, Underline.B)
                    : TextOn(color) == TextColor ? Color.FromArgb(95, Mix(color, Color.Black, 0.55)) : Color.FromArgb(70, 255, 255, 255);
                using (var pen = new Pen(stripe, Math.Max(2f, Px(2.6))))
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

        // The crystal look: the row is cut by a diagonal line into a light face at the left (the color of the row made
        // lighter, or a pale blue for a row without color) and a dark one at the right; the light color fades out into the dark one, going
        // right from the cut (over about 40% of the width), and a white line marks the cut (with a faint light line along its left side).
        // The body is a little raised: half of the shadow of the neumorphic sketch (a light shadow at the lower right, a white glow at the
        // upper left; GDI+ has no blur, so a few translucent shapes of growing size stand for each), and it has a white edge. Returns the
        // color of the text on it.
        private Color DrawCrystal(Graphics g, Rectangle body, bool colored, Color color, double strength)
        {
            bool dark = Theme.BarBack.GetBrightness() < 0.4f;
            int radius = Px(6);
            Color baseColor = colored ? color : Color.FromArgb(150, 185, 225);
            Color light = Mix(baseColor, Color.White, 0.55), deep = Mix(baseColor, Color.Black, 0.12);
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
            // the cut: from (cutX + 8, top) to (cutX - 6, bottom), DIP
            float cutX = body.Left + body.Width * 0.38f;
            var top = new PointF(cutX + Px(8), body.Top);
            var bottom = new PointF(cutX - Px(6), body.Bottom);
            using (GraphicsPath path = RoundedRect(body, radius))
            {
                using (var brush = new SolidBrush(light)) g.FillPath(brush, path);
                Region clip = g.Clip;
                g.SetClip(path, CombineMode.Intersect);
                // the dark face, with a gradient across the cut (its direction is at a right angle to the cut) from the light color to the dark one
                float dx = bottom.X - top.X, dy = bottom.Y - top.Y, len = (float)Math.Sqrt(dx * dx + dy * dy);
                float nx = dy / len, ny = -dx / len; // at a right angle to the cut, pointing to the right
                if (nx < 0) { nx = -nx; ny = -ny; }
                var p1 = new PointF((top.X + bottom.X) / 2, (top.Y + bottom.Y) / 2);
                float span = body.Right - p1.X + Px(2);
                var p2 = new PointF(p1.X + nx * span, p1.Y + ny * span);
                float fade = Math.Min(1f, body.Width * 0.4f / span);
                using (var brush = new LinearGradientBrush(p1, p2, light, deep))
                {
                    brush.InterpolationColors = new ColorBlend(3)
                    {
                        Colors = new[] { light, deep, deep },
                        Positions = new[] { 0f, fade, 1f }
                    };
                    var face = new[] { top, new PointF(body.Right + Px(1), body.Top), new PointF(body.Right + Px(1), body.Bottom), bottom };
                    g.FillPolygon(brush, face);
                }
                using (var pen = new Pen(Color.FromArgb(90, 120, 150, 200), Math.Max(1f, Px(1.0))))
                {
                    g.DrawLine(pen, top.X - Px(1.5), top.Y, bottom.X - Px(1.5), bottom.Y);
                }
                using (var pen = new Pen(Color.White, Math.Max(1.5f, Px(1.6)))) g.DrawLine(pen, top, bottom);
                g.Clip = clip;
                using (var pen = new Pen(Color.White, Math.Max(1.5f, Px(1.2)))) g.DrawPath(pen, path);
            }
            return TextColor;
        }

        // The check badge marker: a round badge at the left end of the active sheet, with a check mark: a green
        // badge with a white check; on a colored row a white badge with a check in the color of the row.
        private void DrawCheckBadge(Graphics g, Rectangle body, bool colored, Color color)
        {
            int full = Math.Max(6, body.Height - Px(8)); // the first size; the badge keeps its left edge and is 0.75 times as large
            int d = Math.Max(5, (int)Math.Round(full * BadgeScale));
            var badge = new Rectangle(body.Left + (body.Height - full) / 2 + Px(1), body.Top + (body.Height - d) / 2, d, d);
            // a light colored row: a dark badge (the row's color, 60% darker) with a white check; a dark colored row: a white badge with
            // a check in the row's color (a lighter or darker one when it is too near to white)
            bool lightRow = colored && TextOn(color) == TextColor;
            Color fill = !colored ? Underline : lightRow ? Mix(color, Color.Black, 0.6) : Color.White;
            Color mark = colored && !lightRow ? Visible(color, Color.White) : Color.White;
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

        // The glass types: a pane of glass over the row, through which the row (its color,
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

        // The bars at both ends of the active sheet: thick rounded bars in the color of the other
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

        // The triangles at both ends of the active sheet, raise pixels above
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

        // A shackle (an upside-down U): a half ring on an ellipse of radii a and b around (cx, cyArc), pw thick, with two straight legs down to
        // `bottom`. It is one closed shape, built the same on the left and on the right.
        private static GraphicsPath ShacklePath(float cx, float cyArc, float a, float b, float pw, float bottom)
        {
            float ao = a + pw / 2, bo = b + pw / 2, ai = Math.Max(0.1f, a - pw / 2), bi = Math.Max(0.1f, b - pw / 2);
            var p = new GraphicsPath();
            p.AddArc(cx - ao, cyArc - bo, ao * 2, bo * 2, 180, 180);
            p.AddLine(cx + ao, cyArc, cx + ao, bottom);
            p.AddLine(cx + ao, bottom, cx + ai, bottom);
            p.AddLine(cx + ai, bottom, cx + ai, cyArc);
            p.AddArc(cx - ai, cyArc - bi, ai * 2, bi * 2, 0, -180);
            p.AddLine(cx - ai, cyArc, cx - ai, bottom);
            p.AddLine(cx - ai, bottom, cx - ao, bottom);
            p.CloseFigure();
            return p;
        }

        // A rounded rectangle on exact (float) edges. RoundedRect above ends one pixel short at the right and the bottom, which is right for the
        // outlines of rows, but it moves the middle of a small badge half a pixel to the left of what is drawn around its own middle.
        private static GraphicsPath RoundedRectF(RectangleF r, float radius)
        {
            var path = new GraphicsPath();
            float d = Math.Max(1f, radius * 2);
            var arc = new RectangleF(r.X, r.Y, d, d);
            path.AddArc(arc, 180, 90);
            arc.X = r.Right - d;
            path.AddArc(arc, 270, 90);
            arc.Y = r.Bottom - d;
            path.AddArc(arc, 0, 90);
            arc.X = r.X;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();
            return path;
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
            if (e.Button == MouseButtons.Right && TryStartChord(e.Location)) return; // (the left button is held: the copy icon, no menu, no focus)
            if (!_noFocusClick) Focus();
            _toggleOffFolder = null; // a drag swallows the button release, so nothing may be left over
            int tool = ToolAt(e.Location);
            if (tool != 0)
            {
                if (e.Button == MouseButtons.Left)
                {
                    if (tool == 5) { FocusSearch(); return; }
                    _toolDown = tool;
                    if (tool == 1 || tool == 2) StartLongPress();
                    Invalidate();
                }
                else if (e.Button == MouseButtons.Right)
                {
                    ContextMenuStrip = null; // (the menu of the sheets must not open on top of the history)
                    if (tool == 1 || tool == 2) ShowHistoryMenu();
                }
                return;
            }
            int all = e.Button == MouseButtons.Left ? AllButtonAt(e.Location) : 0;
            if (all != 0)
            {
                _headerDown = all;
                Invalidate();
                return;
            }
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

            // the progress mark of a sheet: a click changes it (and does not select the sheet)
            if (e.Button == MouseButtons.Left && ProgressBadgeHit(index, e.Location))
            {
                CycleProgress(row.Sheet);
                return;
            }

            // the paint can of a folder: the folder's color (and the folder is not selected)
            if (e.Button == MouseButtons.Left && PaintCanHit(index, e.Location))
            {
                PaintCanDown(index, e.Clicks);
                return;
            }

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
                    ArmChord(row);
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
                ArmChord(row);

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
            if (_chord) { ChordMouseUp(); return; }
            if (e.Button == MouseButtons.Left) _chordArmed = false;
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
            if (_toolDown != 0)
            {
                int down = _toolDown;
                _toolDown = 0;
                if (_longPress != null) _longPress.Stop();
                Invalidate();
                if (e.Button == MouseButtons.Left && !_longPressShown && ToolAt(e.Location) == down)
                {
                    if (down == 1 && _canBack && BackRequested != null) BackRequested();
                    else if (down == 2 && _canForward && ForwardRequested != null) ForwardRequested();
                    else if (down == 3) { ClearSearch(); FocusSearch(); }
                    else if (down == 4)
                    {
                        _folderSearch = !_folderSearch;
                        _searchSel = -1;
                        ApplyFilter();
                        LayoutSearch();
                        if (_query.Length > 0) FocusSearch();
                        Invalidate();
                    }
                }
                _longPressShown = false;
                return;
            }
            if (_headerDown != 0)
            {
                int down = _headerDown;
                _headerDown = 0;
                Invalidate();
                if (e.Button == MouseButtons.Left && AllButtonAt(e.Location) == down && AllFoldersRequested != null) AllFoldersRequested(down == 1);
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
            // Ctrl+F: the search box, only while the pane has the focus (Excel's own find otherwise)
            if (keyData == (Keys.Control | Keys.F) && !IsSample && _editor == null)
            {
                // (moved after the key has been dealt with: Excel gives the focus back to where it was when the key came)
                BeginInvoke(new Action(() => FocusSearch()));
                return true;
            }
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

        // Typing in the pane (not in a name being edited) goes into the search box.
        protected override void OnKeyPress(KeyPressEventArgs e)
        {
            base.OnKeyPress(e);
            if (IsSample || _editor != null || e.Handled) return;
            if (char.IsControl(e.KeyChar) || (ModifierKeys & Keys.Control) != 0 || (e.KeyChar == ' ' && _search.Text.Length == 0)) return;
            e.Handled = true;
            FocusSearch(e.KeyChar.ToString());
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);
            if (IsSample) return;
            if (e.KeyCode == Keys.Escape && Filtering)
            {
                e.Handled = true;
                ClearSearch();
                return;
            }
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
            if (_chord && _chordOnRow) { _chordOnRow = false; Invalidate(); }
            if (_headerHot != 0) { _headerHot = 0; Invalidate(); }
            if (_toolHot != 0) { _toolHot = 0; Invalidate(); }
            _hoverIndex = -1;
            _tip.Hide(this);
        }

        private void UpdateTip(Point p)
        {
            int tool = ToolAt(p);
            if (tool >= 1 && tool <= 4)
            {
                if (_hoverIndex == -2 - tool) return;
                _hoverIndex = -2 - tool;
                _tip.Hide(this);
                string tip = tool == 1 ? "戻る (長押し・右クリックで、履歴)" : tool == 2 ? "進む" : tool == 3 ? "検索を終わる (Esc)" : (_folderSearch ? "フォルダ名で検索中 (押すと、シート名の検索に切り替え)" : "シート名で検索中 (押すと、フォルダ名の検索に切り替え)");
                if (Control.MouseButtons == MouseButtons.None) _tip.Show(tip, this, p.X, ToolPx + Px(2), 3000);
                return;
            }
            int index = IndexFromPoint(p);
            string badgeTip = Control.MouseButtons == MouseButtons.None ? BadgeTipAt(index, p) : null;
            if (badgeTip != _badgeTip)
            {
                _badgeTip = badgeTip;
                _tip.Hide(this);
                if (badgeTip != null)
                {
                    _hoverIndex = index;
                    _tip.Show(badgeTip, this, p.X, RowRect(index).Bottom, 4000);
                    return;
                }
            }
            else if (badgeTip != null) return;
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
                - (row.Indented ? Px(FolderIndent) : 0) - (row.IsFolder ? Px(FolderMarkWidth + 4) : 0) - CheckReserve(row) - RowBadgeReserve(row);
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
            int hotAll = AllButtonAt(e.Location);
            if (hotAll != _headerHot) { _headerHot = hotAll; Invalidate(); }
            int hotTool = ToolAt(e.Location);
            if (hotTool == 5) hotTool = 0;
            if (hotTool != _toolHot) { _toolHot = hotTool; Invalidate(); }
            if (Filtering) { if (e.Button == MouseButtons.Left) _dragIndex = -1; } // the rows can be dragged only in the whole list
            UpdateChordPointer(e.Location);
            UpdateTip(e.Location);
            if (e.Button != MouseButtons.Left || _dragIndex < 0 || _editor != null || StructureProtected || Filtering) return;
            Size drag = SystemInformation.DragSize;
            if (Math.Abs(e.X - _dragOrigin.X) < drag.Width && Math.Abs(e.Y - _dragOrigin.Y) < drag.Height) return;

            int index = _dragIndex;
            _dragIndex = -1;
            _selectionPending = false; // a drag does not change the selection in Excel by itself
            _chordArmed = false;
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
            _dragCopy = (Control.ModifierKeys & Keys.Control) != 0;
            BeginGhost(index);
            try
            {
                DoDragDrop(row.IsFolder ? row.FolderName : row.Sheet, DragDropEffects.Move | DragDropEffects.Copy);
            }
            finally
            {
                _dragging = false;
                SetDropIndex(-1);
                if (_ghostOn && !_ghostSettling) EndGhost();
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
            _dragCopy = IsCopy(e);

            Point p = PointToClient(new Point(e.X, e.Y));
            MoveGhost(p);
            ScrollNearEdge(p);
            if (HasFolders) ResolveFolderDrop(p);
            else { _dropInside = false; _dropInto = _dropBeforeSheet = _dropBeforeFolder = null; SetDropIndex(InsertIndexAt(p)); }
        }

        protected override void OnDragLeave(EventArgs e)
        {
            base.OnDragLeave(e);
            SetDropIndex(-1);
            if (_ghostOn && !_ghostSettling) EndGhost();
        }

        protected override void OnDragDrop(DragEventArgs e)
        {
            base.OnDragDrop(e);
            int target = _dropIndex;
            string into = _dropInto, beforeSheet = _dropBeforeSheet, beforeFolder = _dropBeforeFolder;
            // where the ghost settles: the drop mark; a row dragged down leaves its own place above the mark
            if (_ghostOn && target >= 0 && _rows.Count > 0)
            {
                int markY = target < _rows.Count ? RowRect(target).Top : RowRect(_rows.Count - 1).Bottom;
                bool down = markY > _ghostRowTop && !_dragCopy;
                SettleGhost(markY, down, _dragSheets.Count > 0 ? _dragSheets.Count : 1);
                _dropSignature = _rowSignature;
                _dropNames = new HashSet<string>(_names);
                _wobbleSheets = _dragCopy ? new HashSet<string>() : new HashSet<string>(_dragSheets);
                _wobbleFolders = _dragCopy ? new HashSet<string>() : new HashSet<string>(_dragFolders);
                if (!_dragCopy)
                    foreach (StripRow r in _rows)
                        if (r.IsFolder && _dragFolders.Contains(r.FolderId)) foreach (string m in r.Members) _wobbleSheets.Add(m);
            }
            else if (_ghostOn) EndGhost();
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
            if (p.Y < HeaderPx) return 0;
            if (index < 0) return _rows.Count;
            Rectangle r = RowRect(index);
            return p.Y < r.Top + r.Height / 2 ? index : index + 1;
        }

        // DoDragDrop does not scroll the control, so do it near the top and bottom edges.
        private void ScrollNearEdge(Point p)
        {
            int edge = Math.Max(8, RowPx / 2);
            if (p.Y < HeaderPx + edge) SetScroll(_scrollY - RowPx / 2);
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
                if (PaintCanHit(index, e.Location)) return; // (the double click of the paint can: the color dialog)
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
                if (_animTimer != null) _animTimer.Dispose();
                EndRowAnim();
                if (_raTimer != null) _raTimer.Dispose();
                EndDoneAnim();
                if (_dnTimer != null) _dnTimer.Dispose();
                if (_longPress != null) _longPress.Dispose();
                _tip.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
