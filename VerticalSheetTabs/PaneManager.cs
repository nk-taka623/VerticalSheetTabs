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
            // The folder structure (Agents.md chapter 12); written to the workbook when it is saved.
            public FolderModel Folders = new FolderModel();
            // The sheet names (hidden ones too) the keys below were read for.
            public string KeySignature;
            public List<string> Keys = new List<string>();
            // The names (all sheets, hidden ones too) in Excel's order that Keys belongs to.
            public List<string> AllNames = new List<string>();
            // True while the add-in changes the folders or the sheet order: no sync in between.
            public bool Busy;
        }

        private readonly CustomTaskPaneCollection _collection;
        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();
        // Excel raises no events for rename, reorder or hide/unhide, so the lists
        // are re-read on a short timer (cheap: only names are compared).
        private readonly Timer _timer = new Timer { Interval = 300 };

        // The title bar of the task pane (Excel always shows one).
        public const string PaneTitle = "縦タブ";

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
            Entry existing;
            if (_entries.TryGetValue(hwnd, out existing))
            {
                // Excel can give a new workbook the window handle of one that was closed; the old entry
                // (and its pane) is then stale and must not stop the new workbook from getting a pane.
                bool sameWorkbook;
                try { sameWorkbook = existing.Workbook.Equals(workbook); }
                catch (Exception) { sameWorkbook = false; }
                if (sameWorkbook) return;
                PurgeEntry(existing);
            }

            // First open (no add-in property in the workbook): defaults, i.e. left side.
            var settings = WorkbookSettings.Load(workbook);

            var control = new TabPaneControl();
            var pane = _collection.Add(control, PaneTitle, window);
            pane.DockPosition = ToDock(settings.Side);
            pane.Width = 150; // temporary; the real width is applied in InitializeWidth

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook, Settings = settings };
            entry.Folders = FolderStore.Load(workbook);
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
            control.ExcelCommandRequested += command => RunExcelCommand(entry, command);
            control.FolderToggleRequested += id => ToggleFolder(entry, id);
            control.FolderMoveRequested += (sheets, folders, into, bs, bf, copy) =>
            {
                if (copy) CopyItems(entry, sheets, folders, into, bs, bf);
                else MoveItems(entry, sheets, folders, into, bs, bf);
            };
            control.FolderRenameRequested += (id, name) => RenameFolder(entry, id, name);
            control.NewFolderRequested += names => CreateFolder(entry, names);
            control.FolderCommandRequested += (command, folders) => FolderCommand(entry, command, folders);
            control.FolderColorRequested += (folders, choice) =>
            {
                foreach (string id in folders)
                {
                    FolderNode f = entry.Folders.FindFolder(id);
                    if (f != null) f.Color = choice;
                }
                MarkChanged(entry);
                SyncSheets(entry);
            };
            control.FolderTabColorRequested += (folders, choice) =>
            {
                RefreshKeys(entry);
                ApplyTabColor(entry, MemberNames(entry, folders), choice);
            };
            control.DeleteKeyRequested += (sheets, folders) => DeleteItems(entry, sheets, folders);
            control.DeleteFoldersRequested += (folders, withSheets) =>
            {
                if (withSheets) DeleteItems(entry, new string[0], folders);
                else DeleteFoldersOnly(entry, folders);
            };

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
                // A frame width that is not plausible means the pane was not laid out yet.
                if (entry.Chrome < 0 || entry.Chrome > (int)Math.Round(60 * dpi / 96.0)) entry.Chrome = (int)Math.Round(16 * dpi / 96.0);
                int contentPx = entry.Settings.WidthDip > 0
                    ? (int)Math.Round(entry.Settings.WidthDip * dpi / 96.0)
                    : entry.Control.DefaultContentWidth();
                entry.Pane.Width = contentPx + entry.Chrome;
                entry.Initialized = true;
                VerifyWidth(entry, contentPx, 0);
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

        // The pane is laid out by Excel after the width was set, so the frame width measured at that moment
        // can be wrong (a busy Excel, another DPI). The content width is read back a little later and
        // corrected by the difference, a few times; this way the stored width comes back exactly.
        private void VerifyWidth(Entry entry, int contentPx, int attempt)
        {
            if (attempt >= 4) return;
            var timer = new Timer { Interval = 300 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    if (!_entries.ContainsValue(entry) || !entry.Pane.Visible) return;
                    // the user is dragging the splitter: leave it alone
                    if (System.Windows.Forms.Control.MouseButtons != System.Windows.Forms.MouseButtons.None) return;
                    int actual = entry.Control.Width;
                    if (actual > 0 && Math.Abs(actual - contentPx) > 1)
                    {
                        entry.Pane.Width += contentPx - actual;
                        VerifyWidth(entry, contentPx, attempt + 1);
                    }
                }
                catch (Exception)
                {
                    // The pane was closed meanwhile.
                }
            };
            timer.Start();
        }

        // Called before the workbook is saved: stores the current state in its properties.
        public void SaveSettings(Excel.Workbook workbook)
        {
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                if (!entry.Workbook.Equals(workbook)) continue;
                try
                {
                    FolderStore.Save(entry.Workbook, entry.Folders);
                }
                catch (Exception)
                {
                }
                try
                {
                    entry.Settings.Side = entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight
                        ? TabSide.Right : TabSide.Left;
                    // While the pane has never been shown (horizontal-only mode) the stored width stays as is.
                    if (entry.Initialized && entry.Control.Width >= 40)
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
            if (entry.Busy) return true;
            try
            {
                var names = new List<string>();
                var colors = new Dictionary<string, int>();
                var allNames = new List<string>();
                var visibleAt = new List<bool>();
                bool anyHidden = false;
                foreach (object sheet in entry.Workbook.Sheets)
                {
                    string name;
                    bool visible = IsVisibleSheet(sheet, out name);
                    allNames.Add(name ?? "?");
                    visibleAt.Add(visible);
                    if (visible)
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
                List<StripRow> rows = BuildRows(entry, allNames, visibleAt);
                entry.Control.SetSheets(names, active, selected, colors, rows);
                return true;
            }
            catch (Exception)
            {
                // Excel is busy (e.g. a cell is being edited) or the window is closing.
                // The next timer tick will retry.
                return false;
            }
        }

        // The rows of the vertical list when the workbook has folders; null = one row per visible sheet.
        // Sheets of a closed folder get no row.
        private static List<StripRow> BuildRows(Entry entry, List<string> allNames, List<bool> visibleAt)
        {
            if (!entry.Folders.HasFolders) return null;

            string signature = string.Join("|", allNames);
            if (signature != entry.KeySignature || entry.Keys.Count != allNames.Count)
            {
                entry.Keys = FolderStore.SheetKeys(entry.Workbook, true);
                entry.AllNames = new List<string>(allNames);
                entry.KeySignature = signature;
                entry.Folders.Reconcile(entry.Keys);
            }
            var nameOf = new Dictionary<string, string>();
            for (int i = 0; i < entry.Keys.Count && i < allNames.Count; i++)
            {
                if (visibleAt[i]) nameOf[entry.Keys[i]] = allNames[i];
            }

            var rows = new List<StripRow>();
            foreach (FolderNode node in entry.Folders.Nodes)
            {
                string name;
                if (!node.IsFolder)
                {
                    if (nameOf.TryGetValue(node.SheetCode, out name)) rows.Add(new StripRow { Sheet = name });
                    continue;
                }
                var folderRow = new StripRow
                {
                    FolderId = node.FolderId,
                    FolderName = node.Name,
                    Expanded = node.Expanded,
                    ColorBgr = FolderColorBgr(entry, node.Color)
                };
                rows.Add(folderRow);
                foreach (string code in node.Sheets)
                {
                    if (!nameOf.TryGetValue(code, out name)) continue;
                    folderRow.Members.Add(name);
                    if (node.Expanded) rows.Add(new StripRow { Sheet = name, Indented = true });
                }
            }
            return rows;
        }

        // A folder's color as 0xBBGGRR, -1 for none.
        private static int FolderColorBgr(Entry entry, TabColorChoice choice)
        {
            if (choice == null) return -1;
            if (choice.Kind == TabColorKind.Rgb) return choice.Bgr;
            if (choice.Kind != TabColorKind.Theme) return -1;
            System.Drawing.Color[] theme = ThemeColors(entry);
            int[] order = { 2, 1, 4, 3, 5, 6, 7, 8, 9, 10 };
            int pos = Array.IndexOf(order, choice.ThemeIndex);
            if (theme == null || pos < 0) return -1;
            System.Drawing.Color c = TabColorPalette.ApplyTint(theme[pos], choice.Tint);
            return c.R | (c.G << 8) | (c.B << 16);
        }

        // ---- folder commands (Agents.md chapter 12) ----

        // Reads the sheets of the workbook now, gives them ids and brings the model in line with them.
        private static void RefreshKeys(Entry entry)
        {
            var names = new List<string>();
            foreach (object sheet in entry.Workbook.Sheets) names.Add(SheetName(sheet) ?? "?");
            entry.Keys = FolderStore.SheetKeys(entry.Workbook, true);
            entry.AllNames = names;
            entry.KeySignature = string.Join("|", names);
            entry.Folders.Reconcile(entry.Keys);
        }

        private static string KeyOf(Entry entry, string sheetName)
        {
            int i = entry.AllNames.IndexOf(sheetName);
            return i >= 0 && i < entry.Keys.Count ? entry.Keys[i] : null;
        }

        private static List<string> KeysOf(Entry entry, IEnumerable<string> sheetNames)
        {
            var keys = new List<string>();
            foreach (string n in sheetNames)
            {
                string k = KeyOf(entry, n);
                if (k != null && !keys.Contains(k)) keys.Add(k);
            }
            return keys;
        }

        // Puts the sheets of the workbook in the order of the model (folders are runs of sheets), with
        // as few moves as possible. A grouped selection would move along, so it is undone first and
        // restored afterwards.
        private void ApplyExcelOrder(Entry entry, IList<string> selectedNames, string activeName)
        {
            List<string> desired = entry.Folders.Flatten();
            var current = new List<string>(entry.Keys);
            if (desired.Count != current.Count || desired.SequenceEqual(current)) return;

            bool screen = true;
            try
            {
                screen = _app.ScreenUpdating;
                _app.ScreenUpdating = false;
                entry.Window.Activate();
                UngroupSheets(entry);
                for (int i = 0; i < desired.Count; i++)
                {
                    if (current[i] == desired[i]) continue;
                    int j = current.IndexOf(desired[i]);
                    if (j < 0) break;
                    object sheet = entry.Workbook.Sheets[j + 1];
                    object before = entry.Workbook.Sheets[i + 1];
                    var ws = sheet as Excel.Worksheet;
                    if (ws != null) ws.Move(before);
                    else
                    {
                        var ch = sheet as Excel.Chart;
                        if (ch != null) ch.Move(before);
                    }
                    current.RemoveAt(j);
                    current.Insert(i, desired[i]);
                }
            }
            catch (Exception)
            {
                ShowMessage("ブックの構造が保護されているため、シートを移動できません。");
            }
            finally
            {
                try { _app.ScreenUpdating = screen; } catch (Exception) { }
            }
            entry.KeySignature = null;
            if (selectedNames != null && selectedNames.Count > 0 && activeName != null) SelectSheets(entry, selectedNames, activeName);
        }

        // Runs a change of the folders: the sheets are read, the change is made in the model, the
        // sheet order of Excel follows, and the list is redrawn.
        private void ChangeFolders(Entry entry, Func<bool> change, IList<string> selectedNames, string activeName)
        {
            if (entry.Busy) return;
            entry.Busy = true;
            try
            {
                RefreshKeys(entry);
                if (!change()) return;
                ApplyExcelOrder(entry, selectedNames, activeName);
            }
            catch (Exception)
            {
                ShowMessage("フォルダの操作を完了できませんでした。");
            }
            finally
            {
                entry.Busy = false;
            }
            entry.KeySignature = null;
            SyncSheets(entry);
        }

        private string CurrentActive(Entry entry)
        {
            try
            {
                string name;
                object active = entry.Window.ActiveSheet;
                return active != null && IsVisibleSheet(active, out name) ? name : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private List<string> CurrentSelection(Entry entry)
        {
            var list = new List<string>();
            try
            {
                foreach (object sheet in (System.Collections.IEnumerable)entry.Window.SelectedSheets)
                {
                    string n;
                    if (IsVisibleSheet(sheet, out n)) list.Add(n);
                }
            }
            catch (Exception)
            {
            }
            return list;
        }

        private void MoveItems(Entry entry, IList<string> sheets, IList<string> folders, string into, string beforeSheet, string beforeFolder)
        {
            if (entry.Workbook.ProtectStructure)
            {
                ShowMessage("ブックの構造が保護されているため、シートを移動できません。");
                return;
            }
            IList<string> selected = CurrentSelection(entry);
            string active = CurrentActive(entry);
            ChangeFolders(entry, () =>
            {
                List<string> keys = KeysOf(entry, sheets);
                string beforeKey = beforeSheet == null ? null : KeyOf(entry, beforeSheet);
                return entry.Folders.Move(keys, folders, into, beforeKey, beforeFolder);
            }, selected, active);
        }

        // Ctrl+drag: copies the sheets (and the sheets of dragged folders) to the drop position. A
        // dragged folder is copied as a new folder with copies of its sheets.
        private void CopyItems(Entry entry, IList<string> sheets, IList<string> folders, string into, string beforeSheet, string beforeFolder)
        {
            if (entry.Workbook.ProtectStructure)
            {
                ShowMessage("ブックの構造が保護されているため、シートをコピーできません。");
                return;
            }
            if (entry.Busy) return;
            entry.Busy = true;
            try
            {
                entry.Window.Activate();
                RefreshKeys(entry);
                var oldKeys = new List<string>(entry.Keys);

                // the sources in Excel's order
                var wanted = new HashSet<string>(KeysOf(entry, sheets));
                foreach (string k in entry.Folders.SheetsOfFolders(folders)) wanted.Add(k);
                var sources = oldKeys.Where(wanted.Contains).ToList();
                if (sources.Count == 0) return;
                var sourceNames = sources.Select(k => entry.AllNames[oldKeys.IndexOf(k)]).ToList();

                // where Excel puts the copies: in front of this sheet (null = after the last one)
                string anchorKey = null;
                if (into != null)
                {
                    FolderNode target = entry.Folders.FindFolder(into);
                    if (beforeSheet != null) anchorKey = KeyOf(entry, beforeSheet);
                    else if (target != null && target.Sheets.Count > 0) anchorKey = KeyAfter(entry, target.Sheets[target.Sheets.Count - 1]);
                    else if (target != null) anchorKey = KeyAfterNode(entry, target);
                }
                else if (beforeSheet != null) anchorKey = KeyOf(entry, beforeSheet);
                else if (beforeFolder != null)
                {
                    FolderNode f = entry.Folders.FindFolder(beforeFolder);
                    anchorKey = f == null ? null : f.Sheets.Count > 0 ? f.Sheets[0] : KeyAfterNode(entry, f);
                }
                CopySheetsTo(entry, sourceNames, anchorKey == null ? null : entry.AllNames[oldKeys.IndexOf(anchorKey)]);

                RefreshKeys(entry);
                var copies = entry.Keys.Where(k => !oldKeys.Contains(k)).ToList();
                if (copies.Count != sources.Count) { entry.KeySignature = null; return; }

                // the model: the copies sit where Excel put them; give them their folders
                entry.Folders.RemoveSheets(copies);
                string afterKey = null;
                int last = entry.Keys.IndexOf(copies[copies.Count - 1]);
                if (last + 1 < entry.Keys.Count) afterKey = entry.Keys[last + 1];

                if (into != null && entry.Folders.FindFolder(into) != null)
                {
                    FolderNode target = entry.Folders.FindFolder(into);
                    int at = afterKey == null ? -1 : target.Sheets.IndexOf(afterKey);
                    if (at < 0) target.Sheets.AddRange(copies); else target.Sheets.InsertRange(at, copies);
                }
                else
                {
                    var units = new List<FolderNode>();
                    for (int i = 0; i < sources.Count; i++)
                    {
                        FolderNode src = entry.Folders.FolderOf(sources[i]);
                        if (src != null && folders.Contains(src.FolderId))
                        {
                            FolderNode unit = units.LastOrDefault();
                            if (unit == null || !unit.IsFolder || unit.FolderId != "copy:" + src.FolderId)
                            {
                                unit = FolderModel.NewFolder(UniqueFolderName(entry, src.Name));
                                unit.Color = src.Color;
                                unit.Expanded = src.Expanded;
                                units.Add(unit);
                                unit.FolderId = "copy:" + src.FolderId; // replaced below
                            }
                            unit.Sheets.Add(copies[i]);
                        }
                        else units.Add(new FolderNode { SheetCode = copies[i] });
                    }
                    foreach (FolderNode u in units) if (u.IsFolder) u.FolderId = Guid.NewGuid().ToString("N");

                    int index = entry.Folders.Nodes.Count;
                    if (afterKey != null)
                    {
                        int i = entry.Folders.Nodes.FindIndex(n => n.IsFolder ? n.Sheets.Contains(afterKey) : n.SheetCode == afterKey);
                        if (i >= 0) index = i;
                    }
                    entry.Folders.Nodes.InsertRange(index, units);
                }
                entry.KeySignature = null;
            }
            catch (Exception)
            {
                ShowMessage("シートをコピーできませんでした。");
            }
            finally
            {
                entry.Busy = false;
            }
            entry.KeySignature = null;
            SyncSheets(entry);
        }

        private static string UniqueFolderName(Entry entry, string baseName)
        {
            for (int n = 2; ; n++)
            {
                string suffix = " (" + n + ")";
                string name = (baseName.Length + suffix.Length > FolderModel.MaxNameLength
                    ? baseName.Substring(0, FolderModel.MaxNameLength - suffix.Length) : baseName) + suffix;
                if (!entry.Folders.Nodes.Any(x => x.IsFolder && string.Equals(x.Name, name, StringComparison.OrdinalIgnoreCase))) return name;
            }
        }

        // The key of the sheet that follows the given one in Excel's order, null at the end.
        private static string KeyAfter(Entry entry, string key)
        {
            int i = entry.Keys.IndexOf(key);
            return i >= 0 && i + 1 < entry.Keys.Count ? entry.Keys[i + 1] : null;
        }

        // For an empty folder: the first sheet of the next entries.
        private static string KeyAfterNode(Entry entry, FolderNode node)
        {
            List<FolderNode> nodes = entry.Folders.Nodes;
            for (int i = nodes.IndexOf(node) + 1; i < nodes.Count; i++)
            {
                if (!nodes[i].IsFolder) return nodes[i].SheetCode;
                if (nodes[i].Sheets.Count > 0) return nodes[i].Sheets[0];
            }
            return null;
        }

        // Copies sheets (by name, in order) in front of the sheet beforeName, or after the last sheet.
        private static void CopySheetsTo(Entry entry, IList<string> names, string beforeName)
        {
            object before = Type.Missing;
            object after = Type.Missing;
            if (beforeName != null) before = entry.Workbook.Sheets[beforeName];
            else after = entry.Workbook.Sheets[entry.Workbook.Sheets.Count];

            if (names.Count > 1)
            {
                var group = entry.Workbook.Sheets[names.Cast<object>().ToArray()] as Excel.Sheets;
                if (group != null) group.Copy(before, after);
                return;
            }
            object sheet = entry.Workbook.Sheets[names[0]];
            var ws = sheet as Excel.Worksheet;
            if (ws != null) { ws.Copy(before, after); return; }
            var ch = sheet as Excel.Chart;
            if (ch != null) ch.Copy(before, after);
        }

        // The names of all sheets of the folders, hidden ones too.
        private static List<string> MemberNames(Entry entry, IList<string> folderIds)
        {
            var names = new List<string>();
            foreach (string key in entry.Folders.SheetsOfFolders(folderIds))
            {
                int i = entry.Keys.IndexOf(key);
                if (i >= 0 && i < entry.AllNames.Count) names.Add(entry.AllNames[i]);
            }
            return names;
        }

        // Bulk commands on the sheets of folders (Agents.md 12.4). Excel's own commands work on the
        // selected sheets, so the visible sheets of the folders are selected first.
        private void FolderCommand(Entry entry, string command, IList<string> folderIds)
        {
            try
            {
                RefreshKeys(entry);
                List<string> all = MemberNames(entry, folderIds);
                var visible = new List<string>();
                var hidden = new List<object>();
                foreach (string name in all)
                {
                    object sheet = entry.Workbook.Sheets[name];
                    string n;
                    if (IsVisibleSheet(sheet, out n)) visible.Add(name);
                    else if (SheetVisibility(sheet) == Excel.XlSheetVisibility.xlSheetHidden) hidden.Add(sheet);
                }

                if (command == "unhide")
                {
                    foreach (object sheet in hidden)
                    {
                        var ws = sheet as Excel.Worksheet;
                        if (ws != null) { ws.Visible = Excel.XlSheetVisibility.xlSheetVisible; continue; }
                        var ch = sheet as Excel.Chart;
                        if (ch != null) ch.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    }
                    SyncSheets(entry);
                    return;
                }
                if (visible.Count == 0) return;
                entry.Window.Activate();
                SelectSheets(entry, visible, visible[0]);
                switch (command)
                {
                    case "hide": RunExcelCommand(entry, "SheetHide"); break;
                    case "protect": RunExcelCommand(entry, "SheetProtect"); break;
                    case "movecopy": RunExcelCommand(entry, "SheetMoveOrCopy"); break;
                }
            }
            catch (Exception)
            {
                ShowMessage("フォルダ内のシートの操作を完了できませんでした。");
            }
            SyncSheets(entry);
        }

        private void CreateFolder(Entry entry, IList<string> sheetNames)
        {
            if (entry.Workbook.ProtectStructure)
            {
                ShowMessage("ブックの構造が保護されているため、シートを移動できません。");
                return;
            }
            IList<string> selected = CurrentSelection(entry);
            string active = CurrentActive(entry);
            string createdId = null;
            ChangeFolders(entry, () =>
            {
                List<string> keys = KeysOf(entry, sheetNames.Count > 0 ? sheetNames : (IList<string>)selected);
                if (keys.Count == 0) return false;
                int n = 1;
                while (entry.Folders.Nodes.Any(x => x.IsFolder && x.Name == "フォルダ" + n)) n++;
                createdId = entry.Folders.CreateFolder("フォルダ" + n, keys).FolderId;
                return true;
            }, selected, active);
            if (createdId != null) entry.Control.BeginRenameFolder(createdId);
        }

        private static readonly char[] InvalidFolderChars = { ':', '\\', '/', '?', '*', '[', ']' };

        private void RenameFolder(Entry entry, string folderId, string newName)
        {
            FolderNode folder = entry.Folders.FindFolder(folderId);
            if (folder == null) return;
            if (newName.Trim().Length == 0 || newName.Length > FolderModel.MaxNameLength
                || newName.IndexOfAny(InvalidFolderChars) >= 0 || newName.StartsWith("'") || newName.EndsWith("'"))
            {
                ShowMessage("入力した名前は無効です。次の点を確認してください。\r\n"
                    + "・名前が 31 文字を超えていない\r\n"
                    + "・名前に次の文字が含まれていない: : \\ / ? * [ ]\r\n"
                    + "・名前の先頭と末尾が ' (アポストロフィ) でない\r\n"
                    + "・名前が空白になっていない");
                return;
            }
            if (entry.Folders.Nodes.Any(n => n.IsFolder && n.FolderId != folderId
                && string.Equals(n.Name, newName, StringComparison.OrdinalIgnoreCase)))
            {
                ShowMessage("その名前は既に使われています。別の名前を入力してください。");
                return;
            }
            folder.Name = newName;
            MarkChanged(entry);
            SyncSheets(entry);
        }

        // The folders go, their sheets stay where they are (outside any folder).
        private void DeleteFoldersOnly(Entry entry, IList<string> folderIds)
        {
            foreach (string id in folderIds) entry.Folders.RemoveFolder(id);
            MarkChanged(entry);
            SyncSheets(entry);
        }

        // Delete key / "delete folder and its sheets". Only sheets: Excel's own delete command (with
        // its own confirmation). With folders: our confirmation, then the sheets of the folders and
        // the selected sheets are deleted without Excel's prompts.
        private void DeleteItems(Entry entry, IList<string> sheetNames, IList<string> folderIds)
        {
            if (folderIds.Count == 0)
            {
                RunExcelCommand(entry, "SheetDelete");
                return;
            }
            if (entry.Workbook.ProtectStructure)
            {
                ShowMessage("ブックの構造が保護されているため、シートを削除できません。");
                return;
            }
            var result = System.Windows.Forms.MessageBox.Show(new ExcelWindowOwner(_app.Hwnd),
                "フォルダと中のシートを削除しますが、よろしいですか？", "Microsoft Excel",
                System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning);
            if (result != System.Windows.Forms.DialogResult.Yes) return;

            if (entry.Busy) return;
            entry.Busy = true;
            bool alerts = true;
            try
            {
                RefreshKeys(entry);
                var doomed = new HashSet<string>(KeysOf(entry, sheetNames));
                foreach (string k in entry.Folders.SheetsOfFolders(folderIds)) doomed.Add(k);

                // Excel keeps at least one visible sheet
                bool survivor = false;
                for (int i = 0; i < entry.Keys.Count && !survivor; i++)
                {
                    if (doomed.Contains(entry.Keys[i])) continue;
                    string n;
                    if (IsVisibleSheet(entry.Workbook.Sheets[i + 1], out n)) survivor = true;
                }
                if (!survivor)
                {
                    ShowMessage("ブックには、表示されているシートが少なくとも1つ必要です。");
                    return;
                }

                alerts = _app.DisplayAlerts;
                _app.DisplayAlerts = false;
                for (int i = entry.Keys.Count - 1; i >= 0; i--)
                {
                    if (!doomed.Contains(entry.Keys[i])) continue;
                    object sheet = entry.Workbook.Sheets[i + 1];
                    var ws = sheet as Excel.Worksheet;
                    if (ws != null) { ws.Delete(); continue; }
                    var ch = sheet as Excel.Chart;
                    if (ch != null) ch.Delete();
                }
                entry.Folders.RemoveSheets(doomed);
                entry.Folders.Nodes.RemoveAll(n => n.IsFolder && folderIds.Contains(n.FolderId));
            }
            catch (Exception)
            {
                ShowMessage("シートを削除できませんでした。");
            }
            finally
            {
                try { _app.DisplayAlerts = alerts; } catch (Exception) { }
                entry.Busy = false;
            }
            entry.KeySignature = null;
            SyncSheets(entry);
        }

        // The folder structure is saved with the workbook, so a change of it (open/close, color,
        // name) makes the workbook "changed" like any other edit.
        private static void MarkChanged(Entry entry)
        {
            try { entry.Workbook.Saved = false; } catch (Exception) { }
        }

        private void ToggleFolder(Entry entry, string folderId)
        {
            FolderNode node = entry.Folders.Nodes.Find(n => n.IsFolder && n.FolderId == folderId);
            if (node == null) return;
            node.Expanded = !node.Expanded;
            MarkChanged(entry);
            SyncSheets(entry);
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
                string activeKey = null;
                bool selectionHidden = false;
                if (entry.Folders.HasFolders)
                {
                    // the folder of the active sheet, looked up before the new sheet exists
                    RefreshKeys(entry);
                    activeKey = KeyOf(entry, CurrentActive(entry));
                    // every selected sheet is in a closed folder: the list shows no selection at all
                    List<string> selectedKeys = KeysOf(entry, CurrentSelection(entry));
                    selectionHidden = selectedKeys.Count > 0 && selectedKeys.All(k =>
                    {
                        FolderNode f = entry.Folders.FolderOf(k);
                        return f != null && !f.Expanded;
                    });
                }
                var before = new HashSet<string>(entry.Keys);
                if (selectionHidden)
                {
                    // the new sheet goes to the bottom of the list, outside any folder
                    entry.Workbook.Worksheets.Add(Type.Missing, entry.Workbook.Sheets[entry.Workbook.Sheets.Count], Type.Missing, Type.Missing);
                    entry.KeySignature = null;
                    return;
                }
                entry.Workbook.Worksheets.Add(Type.Missing, entry.Window.ActiveSheet, Type.Missing, Type.Missing);
                FolderNode folder = activeKey == null ? null : entry.Folders.FolderOf(activeKey);
                if (folder != null)
                {
                    // Excel puts the sheet after the active one; behind the last sheet of a folder the
                    // sheet would be outside it, so it is taken into the folder here
                    RefreshKeys(entry);
                    string added = entry.Keys.FirstOrDefault(k => !before.Contains(k));
                    if (added != null && entry.Folders.FolderOf(added) == null)
                    {
                        string anchor = null;
                        int at = folder.Sheets.IndexOf(activeKey);
                        if (at >= 0 && at + 1 < folder.Sheets.Count) anchor = folder.Sheets[at + 1];
                        entry.Folders.Move(new[] { added }, new string[0], folder.FolderId, anchor, null);
                    }
                    entry.KeySignature = null;
                }
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
            bool busyBefore = entry.Busy;
            try
            {
                // Already the selection in Excel: do not touch it, a click on the active tab must not
                // move the focus into the grid (that would cancel a double click that renames).
                if (SameSelection(entry, names, active)) return;

                // Selecting sheet by sheet raises Excel's events; a sync in between would see a half-made selection.
                entry.Busy = true;
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
            finally
            {
                entry.Busy = busyBefore;
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
