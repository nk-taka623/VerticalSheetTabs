using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VerticalSheetTabs
{
    // Keeps one custom task pane per Excel window (keyed by window handle)
    // and keeps each pane's sheet list in sync with its workbook.
    public class PaneManager
    {
        private class Entry
        {
            public CustomTaskPane Pane;
            public TabPaneControl Control;
            public Excel.Window Window;
            public Excel.Workbook Workbook;
            public WorkbookSettings Settings;
            // Pane width minus content width (title bar border, scrollbar gap); measured once.
            public int Chrome;
            public bool Initialized;
            public bool InitPending;
            // Visibility changes we cause ourselves must not be mistaken for the user closing the pane.
            public DateTime IgnoreVisibleUntil;
            // Consecutive failed syncs: the window is probably gone.
            public int Failures;
        }

        private readonly CustomTaskPaneCollection _collection;
        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();
        // Excel raises no events for rename, reorder or hide/unhide, so the lists
        // are re-read on a short timer (cheap: only names are compared).
        private readonly Timer _timer = new Timer { Interval = 300 };

        // The title bar of the task pane (Excel always shows one).
        public const string PaneTitle = "シートタブ";

        private readonly Excel.Application _app;

        public PaneManager(CustomTaskPaneCollection collection, Excel.Application app)
        {
            _collection = collection;
            _app = app;
            _timer.Tick += (s, e) => SyncAll();
            _timer.Start();
        }

        public void EnsurePane(Excel.Window window, Excel.Workbook workbook)
        {
            if (window == null) return;
            int hwnd = window.Hwnd;
            if (_entries.ContainsKey(hwnd)) return;

            // First open (no add-in property in the workbook): defaults, i.e. left side.
            var settings = WorkbookSettings.Load(workbook);

            var control = new TabPaneControl();
            var pane = _collection.Add(control, PaneTitle, window);
            pane.DockPosition = ToDock(settings.Side);
            pane.Width = 150; // temporary; the real width is applied in InitializeWidth

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook, Settings = settings };
            _entries[hwnd] = entry;

            pane.VisibleChanged += (s, e) => OnPaneVisibleChanged(entry);
            control.SelectionRequested += (names, active) => SelectSheets(entry, names, active);
            control.ViewCodeRequested += () => RunExcelCommand(entry, "ViewCode");
            control.ProtectSheetRequested += () => RunExcelCommand(entry, "SheetProtect");
            control.TabColorRequested += (names, choice) => ApplyTabColor(entry, names, choice);
            control.ThemeColorProvider = () => ThemeColors(entry);
            control.SelectAllRequested += () => SelectAllSheets(entry);
            control.UngroupRequested += () => UngroupSheets(entry);
            control.SettingsRequested += () => ShowSettings(entry);
            control.AddRequested += () => AddSheet(entry);
            control.InsertRequested += () => RunExcelCommand(entry, "SheetInsert");
            control.DeleteRequested += () => RunExcelCommand(entry, "SheetDelete");
            control.HideRequested += () => RunExcelCommand(entry, "SheetHide");
            control.UnhideRequested += () => RunExcelCommand(entry, "SheetUnhide");
            control.RenameRequested += (oldName, newName) => RenameSheet(entry, oldName, newName);
            control.MoveCopyRequested += () => RunExcelCommand(entry, "SheetMoveOrCopy");
            control.MoveRequested += (names, before, copy) => MoveOrCopySheets(entry, names, before, copy);

            ApplyMode(entry);
            SyncSheets(entry);
        }

        // Shows/hides the horizontal tabs (Excel's own) and the vertical pane per the display mode.
        private void ApplyMode(Entry entry)
        {
            bool vertical = entry.Settings.Mode != TabMode.HorizontalOnly;
            bool horizontal = entry.Settings.Mode != TabMode.VerticalOnly;
            try
            {
                entry.IgnoreVisibleUntil = DateTime.Now.AddMilliseconds(1000);
                SetHorizontalTabs(entry, horizontal);
                entry.Pane.Visible = vertical;
            }
            catch (Exception)
            {
                // The window may be closing.
            }
            // The pane is laid out asynchronously, so its content width is only
            // reliable a moment after it is shown.
            if (vertical && !entry.Initialized && !entry.InitPending)
            {
                entry.InitPending = true;
                var init = new Timer { Interval = 250 };
                init.Tick += (s, e) =>
                {
                    init.Stop();
                    init.Dispose();
                    InitializeWidth(entry);
                };
                init.Start();
            }
        }

        // Changing the tab bar marks the workbook as modified, which would make Excel ask "save
        // changes?" for a workbook the user never touched. So it is only set when it really changes,
        // and a workbook that was unmodified stays unmodified.
        private static void SetHorizontalTabs(Entry entry, bool show)
        {
            if (entry.Window.DisplayWorkbookTabs == show) return;
            bool wasSaved = entry.Workbook.Saved;
            entry.Window.DisplayWorkbookTabs = show;
            if (wasSaved) entry.Workbook.Saved = true;
        }

        // The user closed the pane with its close button: go back to the horizontal tabs only,
        // otherwise a "vertical only" workbook would be left without any tab bar.
        private void OnPaneVisibleChanged(Entry entry)
        {
            try
            {
                if (!_entries.ContainsValue(entry)) return;
                if (DateTime.Now < entry.IgnoreVisibleUntil) return;
                if (entry.Pane.Visible) return;
                if (_app.ActiveWindow == null || _app.ActiveWindow.Hwnd != entry.Window.Hwnd) return;
                entry.Settings.Mode = TabMode.HorizontalOnly;
                ApplyMode(entry);
            }
            catch (Exception)
            {
                // Window/pane closing.
            }
        }

        // Opens the settings dialog for the active window's workbook.
        public void ShowSettings()
        {
            Entry entry;
            Excel.Window active = _app.ActiveWindow;
            if (active == null || !_entries.TryGetValue(active.Hwnd, out entry)) return;
            ShowSettings(entry);
        }

        // Opens the settings dialog for one window's workbook (ribbon button and tab context menu).
        private void ShowSettings(Entry entry)
        {
            using (var form = new SettingsForm(entry.Settings.Mode, entry.Settings.Side))
            {
                if (form.ShowDialog(new ExcelWindowOwner(_app.Hwnd)) != System.Windows.Forms.DialogResult.OK) return;
                entry.Settings.Mode = form.Mode;
                entry.Settings.Side = form.Side;
            }
            try
            {
                entry.Pane.DockPosition = ToDock(entry.Settings.Side);
            }
            catch (Exception)
            {
                // Window/pane closing.
            }
            ApplyMode(entry);
        }

        private static Office.MsoCTPDockPosition ToDock(TabSide side)
        {
            return side == TabSide.Right
                ? Office.MsoCTPDockPosition.msoCTPDockPositionRight
                : Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        // Applies the stored width, or 8 half-width characters on first open.
        private void InitializeWidth(Entry entry)
        {
            try
            {
                if (!_entries.ContainsValue(entry)) return;
                int dpi = DpiHelper.GetDpi(entry.Control.Handle); // the real DPI (Control.DeviceDpi says 96 in a task pane)
                entry.Chrome = entry.Pane.Width - entry.Control.Width;
                int contentPx = entry.Settings.WidthDip > 0
                    ? (int)Math.Round(entry.Settings.WidthDip * dpi / 96.0)
                    : entry.Control.DefaultContentWidth();
                entry.Pane.Width = contentPx + entry.Chrome;
                entry.Initialized = true;
            }
            catch (Exception)
            {
                // The pane may have been closed meanwhile.
            }
            finally
            {
                entry.InitPending = false;
            }
        }

        // Called before the workbook is saved: stores the current state in its properties.
        public void SaveSettings(Excel.Workbook workbook)
        {
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                if (!entry.Workbook.Equals(workbook)) continue;
                try
                {
                    entry.Settings.Side = entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight
                        ? TabSide.Right : TabSide.Left;
                    // While the pane has never been shown (horizontal-only mode) the stored width stays as is.
                    if (entry.Initialized)
                    {
                        entry.Settings.WidthDip = (int)Math.Round(entry.Control.Width * 96.0 / DpiHelper.GetDpi(entry.Control.Handle));
                    }
                    entry.Settings.Save(workbook);
                }
                catch (Exception)
                {
                    // Window/pane closing: nothing to store.
                }
                return; // one window's state represents the workbook
            }
        }

        // Creates the panes of all open windows that have none (after a cancelled close, or for a
        // window that was opened without being activated).
        public void EnsureAllPanes()
        {
            foreach (Excel.Workbook wb in _app.Workbooks)
            {
                try
                {
                    foreach (Excel.Window w in wb.Windows)
                    {
                        try
                        {
                            EnsurePane(w, wb);
                        }
                        catch (Exception)
                        {
                            // A hidden window, or one that is closing.
                        }
                    }
                }
                catch (Exception)
                {
                    // Workbook closing.
                }
            }
        }

        // The add-in is going away: give Excel's own tab bar back and remove the panes.
        public void Shutdown()
        {
            _timer.Stop();
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                try
                {
                    if (entry.Settings.Mode == TabMode.VerticalOnly) SetHorizontalTabs(entry, true);
                }
                catch (Exception)
                {
                    // Window closing.
                }
                try
                {
                    _collection.Remove(entry.Pane);
                }
                catch (Exception)
                {
                    // Already removed.
                }
            }
            _entries.Clear();
        }

        public void RemovePane(Excel.Window window)
        {
            if (window == null) return;
            int hwnd = window.Hwnd;
            Entry entry;
            if (!_entries.TryGetValue(hwnd, out entry)) return;
            _entries.Remove(hwnd);
            _collection.Remove(entry.Pane);
        }

        // Reading every sheet's name, visibility and tab color costs several COM calls per sheet, so
        // with hundreds of sheets a fixed 300 ms period would keep Excel busy. The period follows the
        // measured cost of one pass (at most about 1/8 of the time is spent here), between 0.3 and 3 s.
        public void SyncAll()
        {
            var watch = System.Diagnostics.Stopwatch.StartNew();
            SyncEntries();
            watch.Stop();
            int next = (int)Math.Max(300, Math.Min(3000, watch.ElapsedMilliseconds * 8));
            if (_timer.Interval != next) _timer.Interval = next;
        }

        private DateTime _lastEnsure = DateTime.MinValue;

        private void SyncEntries()
        {
            // A pane that is missing for an open window (the close was cancelled, or the pane could not
            // be created while Excel was showing a dialog) is created again, at most every 3 seconds.
            if (DateTime.Now - _lastEnsure > TimeSpan.FromSeconds(3))
            {
                _lastEnsure = DateTime.Now;
                try
                {
                    if (_app.Windows.Count != _entries.Count) EnsureAllPanes();
                }
                catch (Exception)
                {
                    // Excel is busy or closing.
                }
            }

            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                if (SyncSheets(entry)) entry.Failures = 0;
                else if (++entry.Failures > 20) PurgeEntry(entry); // about 6 seconds without an answer
            }
        }

        // A window that stays unreachable (closed without our noticing) loses its pane entry.
        private void PurgeEntry(Entry entry)
        {
            foreach (var pair in new List<KeyValuePair<int, Entry>>(_entries))
            {
                if (pair.Value == entry) _entries.Remove(pair.Key);
            }
            try
            {
                _collection.Remove(entry.Pane);
            }
            catch (Exception)
            {
                // Already gone.
            }
        }

        // Returns false when Excel could not be asked (busy, or the window/workbook is gone).
        private static bool SyncSheets(Entry entry)
        {
            try
            {
                var names = new List<string>();
                var colors = new Dictionary<string, int>();
                bool anyHidden = false;
                foreach (object sheet in entry.Workbook.Sheets)
                {
                    string name;
                    if (IsVisibleSheet(sheet, out name))
                    {
                        names.Add(name);
                        int bgr;
                        if (TryGetTabColor(sheet, out bgr)) colors[name] = bgr;
                    }
                    else if (SheetVisibility(sheet) == Excel.XlSheetVisibility.xlSheetHidden) anyHidden = true;
                }
                entry.Control.HasHiddenSheets = anyHidden;
                entry.Control.StructureProtected = entry.Workbook.ProtectStructure;

                string active = null;
                object activeSheet = entry.Window.ActiveSheet;
                string activeName;
                if (activeSheet != null && IsVisibleSheet(activeSheet, out activeName)) active = activeName;

                // The selected sheets form Excel's group (a single entry when nothing is grouped).
                var selected = new List<string>();
                foreach (object sheet in (System.Collections.IEnumerable)entry.Window.SelectedSheets)
                {
                    string selName;
                    if (IsVisibleSheet(sheet, out selName)) selected.Add(selName);
                }
                if (selected.Count == 0 && active != null) selected.Add(active);

                entry.Control.ActiveSheetProtected = IsSheetProtected(activeSheet);
                entry.Control.SetSheets(names, active, selected, colors);
                return true;
            }
            catch (Exception)
            {
                // Excel is busy (e.g. a cell is being edited) or the window is closing.
                // The next timer tick will retry.
                return false;
            }
        }

        // The tab color as Excel stores it (0xBBGGRR); false when the tab has no color.
        private static bool TryGetTabColor(object sheet, out int bgr)
        {
            bgr = 0;
            Excel.Tab tab = null;
            var ws = sheet as Excel.Worksheet;
            if (ws != null) tab = ws.Tab;
            var ch = sheet as Excel.Chart;
            if (ch != null) tab = ch.Tab;
            if (tab == null) return false;

            object index = tab.ColorIndex;
            if (index is int && (int)index == (int)Excel.XlColorIndex.xlColorIndexNone) return false;
            object color = tab.Color;
            if (!(color is int)) return false;
            bgr = (int)color;
            return true;
        }

        private static bool IsSheetProtected(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) return ws.ProtectContents;
            var ch = sheet as Excel.Chart;
            return ch != null && ch.ProtectContents;
        }

        // The workbook's theme colors in the order of Excel's palette (white, black, two
        // backgrounds, six accents); null falls back to the Office theme in the palette.
        private static System.Drawing.Color[] ThemeColors(Entry entry)
        {
            try
            {
                Office.ThemeColorScheme scheme = entry.Workbook.Theme.ThemeColorScheme;
                // XlThemeColor / MsoThemeColorSchemeIndex: 2 light1, 1 dark1, 4 light2, 3 dark2, 5..10 accent1..6
                int[] order = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };
                var colors = new System.Drawing.Color[10];
                for (int i = 0; i < 10; i++)
                {
                    int bgr = scheme.Colors((Office.MsoThemeColorSchemeIndex)order[i]).RGB;
                    colors[i] = System.Drawing.Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
                }
                return colors;
            }
            catch (Exception)
            {
                return null;
            }
        }

        // Applies a tab color to every selected sheet, as Excel does for a group.
        private void ApplyTabColor(Entry entry, IList<string> names, TabColorChoice choice)
        {
            try
            {
                foreach (string name in names)
                {
                    object sheet = entry.Workbook.Sheets[name];
                    Excel.Tab tab = null;
                    var ws = sheet as Excel.Worksheet;
                    if (ws != null) tab = ws.Tab;
                    var ch = sheet as Excel.Chart;
                    if (ch != null) tab = ch.Tab;
                    if (tab == null) continue;

                    switch (choice.Kind)
                    {
                        case TabColorKind.None:
                            tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                            break;
                        case TabColorKind.Rgb:
                            tab.Color = choice.Bgr;
                            break;
                        default:
                            tab.ThemeColor = (Excel.XlThemeColor)choice.ThemeIndex;
                            tab.TintAndShade = choice.Tint;
                            break;
                    }
                }
            }
            catch (Exception)
            {
                ShowMessage("シート見出しの色を変更できませんでした。シートが保護されている可能性があります。");
            }
        }

        // Only "hidden" counts for the unhide command; very hidden sheets cannot be unhidden from the UI.
        private static Excel.XlSheetVisibility SheetVisibility(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) return ws.Visible;
            var ch = sheet as Excel.Chart;
            return ch != null ? ch.Visible : Excel.XlSheetVisibility.xlSheetVisible;
        }

        // Worksheets and chart sheets are shown; hidden and very hidden sheets are not.
        private static bool IsVisibleSheet(object sheet, out string name)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null)
            {
                name = ws.Name;
                return ws.Visible == Excel.XlSheetVisibility.xlSheetVisible;
            }
            var ch = sheet as Excel.Chart;
            if (ch != null)
            {
                name = ch.Name;
                return ch.Visible == Excel.XlSheetVisibility.xlSheetVisible;
            }
            name = null;
            return false;
        }

        // Runs one of Excel's own sheet commands on the active sheet (the right-clicked one), so
        // dialogs, confirmations and error messages are exactly Excel's.
        private void RunExcelCommand(Entry entry, string idMso)
        {
            try
            {
                entry.Window.Activate();
                _app.CommandBars.ExecuteMso(idMso);
            }
            catch (Exception)
            {
                // The command is not available now (e.g. structure protection); Excel greys it out
                // in its own menu, so there is nothing more to report.
            }
        }

        // The "+" button: a new worksheet after the active sheet, as Excel's own button does.
        private void AddSheet(Entry entry)
        {
            try
            {
                entry.Window.Activate();
                entry.Workbook.Worksheets.Add(Type.Missing, entry.Window.ActiveSheet, Type.Missing, Type.Missing);
            }
            catch (Exception)
            {
                ShowMessage("ブックの構造が保護されているため、シートを追加できません。");
            }
        }

        // Drag and drop: move (or Ctrl+drag copy) sheets in front of another visible sheet, or to
        // the end of the visible ones. Several sheets (a group) keep their order. Hidden sheets stay
        // where they are relative to their neighbours.
        private void MoveOrCopySheets(Entry entry, IList<string> names, string beforeName, bool copy)
        {
            try
            {
                entry.Window.Activate();
                object before = Type.Missing;
                object after = Type.Missing;
                if (beforeName != null) before = entry.Workbook.Sheets[beforeName];
                else after = LastVisibleSheet(entry.Workbook);

                if (names.Count > 1)
                {
                    var group = entry.Workbook.Sheets[names.Cast<object>().ToArray()] as Excel.Sheets;
                    if (group == null) return;
                    if (copy) group.Copy(before, after); else group.Move(before, after);
                    return;
                }

                object sheet = entry.Workbook.Sheets[names[0]];
                var ws = sheet as Excel.Worksheet;
                if (ws != null)
                {
                    if (copy) ws.Copy(before, after); else ws.Move(before, after);
                    return;
                }
                var ch = sheet as Excel.Chart;
                if (ch != null)
                {
                    if (copy) ch.Copy(before, after); else ch.Move(before, after);
                }
            }
            catch (Exception)
            {
                ShowMessage("ブックの構造が保護されているため、シートを移動またはコピーできません。");
            }
        }

        private static object LastVisibleSheet(Excel.Workbook workbook)
        {
            object last = Type.Missing;
            foreach (object sheet in workbook.Sheets)
            {
                string name;
                if (IsVisibleSheet(sheet, out name)) last = sheet;
            }
            return last;
        }

        private void RenameSheet(Entry entry, string oldName, string newName)
        {
            string problem = CheckSheetName(entry, oldName, newName);
            if (problem != null)
            {
                ShowMessage(problem);
                return;
            }
            try
            {
                object sheet = entry.Workbook.Sheets[oldName];
                var ws = sheet as Excel.Worksheet;
                if (ws != null) { ws.Name = newName; return; }
                var ch = sheet as Excel.Chart;
                if (ch != null) ch.Name = newName;
            }
            catch (Exception)
            {
                ShowMessage("シートの名前を変更できませんでした。ブックの構造が保護されている可能性があります。");
            }
        }

        private static readonly char[] InvalidNameChars = { ':', '\\', '/', '?', '*', '[', ']' };

        // The same rules as Excel's own sheet-name check.
        private static string CheckSheetName(Entry entry, string oldName, string newName)
        {
            if (newName.Trim().Length == 0 || newName.Length > 31 || newName.IndexOfAny(InvalidNameChars) >= 0
                || newName.StartsWith("'") || newName.EndsWith("'"))
            {
                return "入力した名前は無効です。次の点を確認してください。\r\n"
                    + "・名前が 31 文字を超えていない\r\n"
                    + "・名前に次の文字が含まれていない: : \\ / ? * [ ]\r\n"
                    + "・名前の先頭と末尾が ' (アポストロフィ) でない\r\n"
                    + "・名前が空白になっていない";
            }
            foreach (object sheet in entry.Workbook.Sheets)
            {
                // Hidden sheets count for the duplicate check as well.
                string name = SheetName(sheet);
                if (name != null && !string.Equals(name, oldName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(name, newName, StringComparison.OrdinalIgnoreCase))
                {
                    return "その名前は既に使われています。別の名前を入力してください。";
                }
            }
            return null;
        }

        private static string SheetName(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) return ws.Name;
            var ch = sheet as Excel.Chart;
            return ch != null ? ch.Name : null;
        }

        private void ShowMessage(string text)
        {
            System.Windows.Forms.MessageBox.Show(new ExcelWindowOwner(_app.Hwnd), text, "Microsoft Excel",
                System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

        // Makes these sheets the window's selection (Excel's sheet group) with one of them active.
        // The active sheet is selected first (replacing the old selection): adding sheets to a
        // selection with Select(false) does not change which sheet is active.
        private static void SelectSheets(Entry entry, IList<string> names, string active)
        {
            try
            {
                // Already the selection in Excel: do not touch it, a click on the active tab must not
                // move the focus into the grid (that would cancel a double click that renames).
                if (SameSelection(entry, names, active)) return;

                var ordered = new List<string> { active };
                foreach (string n in names) if (n != active) ordered.Add(n);

                // Selecting one sheet at a time costs about 12 ms each (3.5 s for 300 sheets), so a
                // big group is selected with one call.
                if (ordered.Count > 12)
                {
                    var group = entry.Workbook.Sheets[ordered.Cast<object>().ToArray()] as Excel.Sheets;
                    if (group != null)
                    {
                        group.Select(true);
                        return;
                    }
                }

                bool first = true;
                foreach (string n in ordered)
                {
                    SelectSheet(entry.Workbook.Sheets[n], first);
                    first = false;
                }
            }
            catch (Exception)
            {
                // The sheet may have just been renamed or deleted; the next sync fixes the list.
            }
        }

        private static bool SameSelection(Entry entry, IList<string> names, string active)
        {
            try
            {
                string current;
                object activeSheet = entry.Window.ActiveSheet;
                if (activeSheet == null || !IsVisibleSheet(activeSheet, out current) || current != active) return false;

                var selected = new HashSet<string>();
                foreach (object sheet in (System.Collections.IEnumerable)entry.Window.SelectedSheets)
                {
                    string name;
                    if (IsVisibleSheet(sheet, out name)) selected.Add(name);
                }
                return selected.SetEquals(names);
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void SelectSheet(object sheet, bool replace)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) { ws.Select(replace); return; }
            var ch = sheet as Excel.Chart;
            if (ch != null) ch.Select(replace);
        }

        // "Select all sheets": Excel's own command (fast, and the active sheet stays active).
        private void SelectAllSheets(Entry entry)
        {
            RunExcelCommand(entry, "SelectAllSheets");
        }

        // "Ungroup sheets": only the active sheet stays selected.
        private static void UngroupSheets(Entry entry)
        {
            try
            {
                string name;
                object activeSheet = entry.Window.ActiveSheet;
                if (activeSheet != null && IsVisibleSheet(activeSheet, out name)) SelectSheet(activeSheet, true);
            }
            catch (Exception)
            {
                // Busy or closing.
            }
        }
    }
}
