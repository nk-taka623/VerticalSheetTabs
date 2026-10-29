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
            // Excel's "Protect Sheet" dialog is open because of a click on a done mark: a second request must not start another one.
            public bool ProtectAsking;
            public System.Drawing.Color[] ThemeCache;   // the workbook's theme colors, read at most every few seconds
            public DateTime ThemeCacheAt = DateTime.MinValue;
            public int ThemeCacheSeconds = 5;
            // The theme of the saved file, kept while the file is the same (path, time, size), and the names of the built-in styles that
            // gave each theme color the last time (they are read directly while they still give it).
            public string FileThemeStamp;
            public System.Drawing.Color[] FileTheme;
            public Dictionary<int, string> StyleSources = new Dictionary<int, string>();
            public bool InitPending;
            // Visibility changes we cause ourselves must not be mistaken for the user closing the pane.
            public DateTime IgnoreVisibleUntil;
            // Whether the add-in wants the pane shown (what it last set): a hidden pane while this is true was closed by the user.
            public bool WantVisible;
            // Consecutive failed syncs: the window is probably gone.
            public int Failures;
            // The folder structure; written to the workbook when it is saved.
            public FolderModel Folders = new FolderModel();
            // The folder part of the workbook could not be read: it is not deleted when the workbook is saved.
            public bool FoldersUnreadable;
            // The handle of the window (the key of the entry).
            public int Hwnd;
            // The sheet names (hidden ones too) the keys below were read for.
            public string KeySignature;
            public List<string> Keys = new List<string>();
            // The names (all sheets, hidden ones too) in Excel's order that Keys belongs to.
            public List<string> AllNames = new List<string>();
            // True while the add-in changes the folders or the sheet order: no sync in between.
            public bool Busy;
            // True while a sync of this window is running (see SyncSheets).
            public bool Syncing;
            // The sheet-name rule: the names the grouping was last checked for, and a
            // flag for a sheet move the sync made itself (the workbook is then really changed).
            // Folders that keep their sheets in the folder's color: the sheets each had at the last look
            // (a sheet that is new in the list is colored).
            public Dictionary<string, HashSet<string>> SyncSeen = new Dictionary<string, HashSet<string>>();
            public string RuleSignature;
            public bool RuleChanged;
            // Just opened with the sheet-name rule: the sheets are put in order without making the workbook
            // "changed" (opening it and closing it again must not ask "save changes?"), until the order holds.
            public bool QuietRegroup;
            // Sheets whose key could not be read: consecutive tries, and keys that were taken although partly guessed.
            public int KeyMisses;
            public bool KeysGuessed;
            // The last SelectSheets really changed the selection of Excel (and the focus may go to the sheet).
            public bool SelectionChanged;
            // The sheets seen last; written to the workbook when it is saved.
            public SheetHistory History = new SheetHistory();
            // The active sheet and the sheet names the history was last updated for.
            public string HistActive;
            public string HistNames;
            // The name each sheet of the history had when it was last seen (by key): a change of the sheet names is checked against these
            // few sheets, and all keys are read only when one of them is not where it was.
            public Dictionary<string, string> HistHints = new Dictionary<string, string>();
            // The keys of all sheets read in the running sync (for the folders), for the sheet names TickSignature: the history uses them
            // instead of reading them again. Null when the sync read none.
            public List<string> TickKeys;
            public bool TickKeysComplete;
            public string TickSignature;
            // What the sheets hold (errors, notes) for the badges, by sheet name, and the sheets to look at again
            // with the time of the change that makes them so (looked at again when it has been quiet for a while).
            public Dictionary<string, BadgeInfo> Badges = new Dictionary<string, BadgeInfo>();
            public Dictionary<string, DateTime> BadgeDirty = new Dictionary<string, DateTime>();
            // When each sheet was last looked at, and how long that took (ms): the sheets are looked at again now and then, and a sheet that is
            // slow to look at is not looked at again at once.
            public Dictionary<string, DateTime> BadgeScanned = new Dictionary<string, DateTime>();
            public Dictionary<string, int> BadgeCost = new Dictionary<string, int>();
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

            // The looks come from the PC's settings file, the width and the sheet-name rule from the workbook.
            var settings = WorkbookSettings.Load(workbook);

            var control = new TabPaneControl();
            var pane = _collection.Add(control, PaneTitle, window);
            pane.DockPosition = ToDock(settings.Side);
            pane.Width = 150; // temporary; the real width is applied in InitializeWidth

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook, Settings = settings };
            entry.QuietRegroup = settings.NameRule;
            entry.Hwnd = hwnd;
            entry.Folders = FolderStore.Load(workbook);
            entry.FoldersUnreadable = entry.Folders.Unreadable;
            entry.History = HistoryStore.Load(workbook);
            _entries[hwnd] = entry;

            control.OutlineVisible = settings.OutlineVisible;
            control.OutlineLevel = settings.OutlineLevel;
            control.SelectionStyle = settings.SelectStyle;
            control.ActiveMarker = settings.ActiveMarker;
            control.ShownBadges = settings.Badges;
            control.RowAnimation = settings.RowAnim;
            control.DoneAnimation = settings.DoneAnim;
            control.ProgressRequested += (sheets, state) => SetProgress(entry, sheets, state);
            control.ProgressProtectRequested += sheet => ToggleSheetProtect(entry, sheet);
            control.SetWobble(settings.Wobble, settings.WobbleAmp, settings.WobbleDur);
            control.SetWobbleSound(settings.WobbleSound, settings.WobbleVolume);

            pane.VisibleChanged += (s, e) => OnPaneVisibleChanged(entry);
            control.IsExcelEditing = IsExcelEditing;
            control.MoreColorsProvider = () => PickMoreColor(entry);
            control.SelectionRequested += (names, active) =>
            {
                CommitCellEdit(entry);
                SelectSheets(entry, names, active);
            };
            control.GridFocusRequested += hadFocus => FocusGrid(entry, hadFocus);
            control.GridFocusForced += () => { entry.SelectionChanged = true; FocusGrid(entry, true); };
            control.BackRequested += () => NavigateHistory(entry, 1);
            control.ForwardRequested += () => NavigateHistory(entry, -1);
            control.HistoryProvider = () => HistoryItems(entry);
            control.HistoryJumpRequested += key => JumpToHistory(entry, key);
            control.HistoryClearRequested += () => ClearHistory(entry);
            control.ViewCodeRequested += () => RunExcelCommand(entry, "ViewCode");
            control.ProtectSheetRequested += () => RunExcelCommand(entry, "SheetProtect");
            control.TabColorRequested += (names, choice) => ApplyTabColor(entry, names, choice);
            control.ThemeColorProvider = () => ThemeColors(entry);
            control.SelectAllRequested += () => SelectAllSheets(entry);
            control.UngroupRequested += () => UngroupSheets(entry);
            control.SettingsRequested += () => ShowSettings(entry);
            control.AddRequested += () => AddSheet(entry);
            control.InsertRequested += () => InsertSheet(entry);
            control.DeleteRequested += () => RunExcelCommand(entry, "SheetDelete");
            control.HideRequested += () => RunExcelCommand(entry, "SheetHide");
            control.UnhideRequested += () => RunExcelCommand(entry, "SheetUnhide");
            control.RenameRequested += (oldName, newName) => RenameSheet(entry, oldName, newName);
            control.MoveCopyRequested += () => RunExcelCommand(entry, "SheetMoveOrCopy");
            control.MoveRequested += (names, before, copy) => MoveOrCopySheets(entry, names, before, copy);
            control.ExcelCommandRequested += command => RunExcelCommand(entry, command);
            control.FolderToggleRequested += id => ToggleFolder(entry, id);
            control.AllFoldersRequested += open => FolderCommand(entry, open ? "openall" : "closeall", new List<string>());
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
                RefreshKeys(entry);
                foreach (string id in folders)
                {
                    FolderNode f = entry.Folders.FindFolder(id);
                    if (f == null) continue;
                    f.Color = choice;
                    // a folder that keeps its sheets in its color: they follow
                    if (f.SyncColor) ApplyTabColor(entry, MemberNames(entry, new[] { id }), choice);
                }
                MarkChanged(entry);
                SyncSheets(entry);
            };
            control.DeleteKeyRequested += (sheets, folders) => DeleteItems(entry, sheets, folders);
            control.DeleteFoldersRequested += (folders, withSheets) =>
            {
                if (withSheets) DeleteItems(entry, new string[0], folders);
                else DeleteFoldersOnly(entry, folders);
            };

            // A workbook that an older version saved in the "vertical only" mode has its horizontal tabs hidden: they come back.
            if (settings.LegacyVerticalOnly)
            {
                try { SetHorizontalTabs(entry, true); }
                catch (Exception) { }
            }
            ApplyMode(entry);
            SyncSheets(entry);
            NotifyBadgeEvents();
        }

        // The events of the workbook (a cell changed, a sheet calculated or left) are only listened to while a badge is shown in some
        // window: every event costs Excel a call into the add-in, and a macro that writes cells makes thousands of them.
        public Action<bool> BadgeEventsWanted;
        private bool _badgeEventsOn;

        private void NotifyBadgeEvents()
        {
            bool wanted = _entries.Values.Any(e => e.Settings.Badges != BadgeKinds.None);
            if (wanted == _badgeEventsOn) return;
            _badgeEventsOn = wanted;
            try { if (BadgeEventsWanted != null) BadgeEventsWanted(wanted); }
            catch (Exception) { }
        }

        // The vertical pane is shown while this window's display mode is "both" and the horizontal tabs are shown (a workbook whose
        // horizontal tabs are hidden is meant to have no tabs: the add-in never changes that).
        private static bool WantsVertical(Entry entry)
        {
            return entry.Settings.Mode == TabMode.Both && entry.Window.DisplayWorkbookTabs;
        }

        // Shows or hides the vertical pane.
        private void ApplyMode(Entry entry)
        {
            bool vertical = false;
            try
            {
                vertical = WantsVertical(entry);
                if (entry.Pane.Visible != vertical || entry.WantVisible != vertical)
                {
                    entry.IgnoreVisibleUntil = DateTime.Now.AddMilliseconds(1000);
                    entry.WantVisible = vertical;
                    entry.Pane.Visible = vertical;
                }
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

        // Shows or hides the horizontal tabs (only used to give them back to a workbook an older version had hidden them in).
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

        // The user closed the pane with its close button: this window goes to "horizontal only". It is not written to the
        // PC's settings; a window opened later shows the vertical tabs again.
        private void OnPaneVisibleChanged(Entry entry)
        {
            try
            {
                if (!_entries.ContainsValue(entry)) return;
                if (DateTime.Now < entry.IgnoreVisibleUntil) return;
                // (a pane hidden by the add-in itself is not wanted visible)
                if (entry.Pane.Visible || !entry.WantVisible) return;
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
            bool nameRuleTurnedOn = false;
            BadgeKinds oldBadges = entry.Settings.Badges;
            bool oldNameRule = entry.Settings.NameRule;
            // The file may have been changed by another window or Excel since this window read it: the dialog starts with the
            // file's values (except the mode, which is this window's own state: the user may have closed the pane with x).
            TabMode windowMode = entry.Settings.Mode;
            PcSettings.ApplyTo(entry.Settings);
            entry.Settings.Mode = windowMode;
            // what the dialog starts with: only the values it changes are written to the PC's settings
            Dictionary<string, string> before = PcSettings.ToValues(entry.Settings);
            bool tabsHidden = false;
            try { tabsHidden = !entry.Window.DisplayWorkbookTabs; }
            catch (Exception) { }
            using (var form = new SettingsForm(tabsHidden, entry.Settings.Mode, entry.Settings.Side, entry.Settings.OutlineVisible, entry.Settings.OutlineLevel, entry.Settings.NameRule, entry.Settings.SelectStyle, entry.Settings.ActiveMarker, entry.Settings.Wobble, entry.Settings.WobbleAmp, entry.Settings.WobbleDur, entry.Settings.WobbleSound, entry.Settings.WobbleVolume, entry.Settings.WobbleSampleStyle, entry.Settings.Badges, entry.Settings.RowAnim, entry.Settings.DoneAnim))
            {
                if (form.ShowDialog(new ExcelWindowOwner(_app.Hwnd)) != System.Windows.Forms.DialogResult.OK) return;
                entry.Settings.Mode = form.Mode;
                entry.Settings.Side = form.Side;
                entry.Settings.OutlineVisible = form.OutlineVisible;
                entry.Settings.OutlineLevel = form.OutlineLevel;
                nameRuleTurnedOn = !entry.Settings.NameRule && form.NameRule;
                entry.Settings.NameRule = form.NameRule;
                entry.Settings.SelectStyle = form.SelectStyle;
                entry.Settings.ActiveMarker = form.ActiveMarker;
                entry.Settings.Badges = form.Badges;
                entry.Settings.RowAnim = form.RowAnim;
                entry.Settings.DoneAnim = form.DoneAnim;
                if (oldBadges != form.Badges) { entry.Badges.Clear(); entry.BadgeDirty.Clear(); entry.BadgeScanned.Clear(); entry.BadgeCost.Clear(); } // (the sheets are looked at again for the new choice)
                entry.Settings.Wobble = form.Wobble;
                entry.Settings.WobbleAmp = form.WobbleAmp;
                entry.Settings.WobbleDur = form.WobbleDur;
                entry.Settings.WobbleSound = form.WobbleSound;
                entry.Settings.WobbleVolume = form.WobbleVolume;
                entry.Settings.WobbleSampleStyle = form.WobbleSampleStyle;
            }
            PcSettings.Update(before, PcSettings.ToValues(entry.Settings));
            if (entry.Settings.NameRule != oldNameRule) MarkChanged(entry);
            try
            {
                entry.Control.OutlineVisible = entry.Settings.OutlineVisible;
                entry.Control.OutlineLevel = entry.Settings.OutlineLevel;
                entry.Control.SelectionStyle = entry.Settings.SelectStyle;
                entry.Control.ActiveMarker = entry.Settings.ActiveMarker;
                entry.Control.ShownBadges = entry.Settings.Badges;
                entry.Control.RowAnimation = entry.Settings.RowAnim;
                entry.Control.DoneAnimation = entry.Settings.DoneAnim;
                entry.Control.SetWobble(entry.Settings.Wobble, entry.Settings.WobbleAmp, entry.Settings.WobbleDur);
                entry.Control.SetWobbleSound(entry.Settings.WobbleSound, entry.Settings.WobbleVolume);
            }
            catch (Exception)
            {
                // Control closing.
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
            NotifyBadgeEvents();
            if (nameRuleTurnedOn) EnableNameRule(entry);
        }

        private static Office.MsoCTPDockPosition ToDock(TabSide side)
        {
            return side == TabSide.Right
                ? Office.MsoCTPDockPosition.msoCTPDockPositionRight
                : Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
        }

        // Applies the stored width, or 12 half-width characters on first open.
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
                    // sheets added without the sync having seen them get their id now, so that the file
                    // has the ids and opening it does not have to give them (and change the workbook)
                    if (entry.Folders.HasFolders) RefreshKeys(entry);
                    FolderStore.Save(entry.Workbook, entry.Folders, entry.FoldersUnreadable && !entry.Folders.HasFolders);
                }
                catch (Exception)
                {
                }
                HistoryStore.Save(entry.Workbook, entry.History);
                try
                {
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
                            if (!w.Visible) continue; // (a hidden window gets its pane when it is shown)
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

        // The add-in is going away: remove the panes (Excel's own tab bar is never touched).
        public void Shutdown()
        {
            _timer.Stop();
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
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
            // a dialog of Excel is open: neither the panes are looked after nor the sheets read
            if (ExcelMainWindowDisabled(null)) return;

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
                AllowPaneInput(entry);
                CheckHorizontalTabs(entry);
                if (SyncSheets(entry, true)) entry.Failures = 0; // (the dialog check above holds for this tick)
                else if (++entry.Failures > 20) DropIfWindowGone(entry); // about 6 seconds without an answer
            }
        }

        // Excel raised SheetActivate: the windows of that sheet's workbook are synced at once. The other workbooks did not change; the
        // timer syncs them as always.
        public void SyncSheetOf(object sheet)
        {
            string book = null;
            try
            {
                var ws = sheet as Excel.Worksheet;
                var ch = sheet as Excel.Chart;
                var wb = (ws != null ? ws.Parent : ch != null ? ch.Parent : null) as Excel.Workbook;
                if (wb != null) book = wb.Name;
            }
            catch (Exception)
            {
                // Excel is busy
            }
            if (book == null)
            {
                SyncAll();
                return;
            }
            if (ExcelMainWindowDisabled(null)) return;
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                string name;
                try { name = entry.Workbook.Name; }
                catch (Exception) { continue; } // (the timer looks after a window that is gone)
                if (name != book) continue;
                AllowPaneInput(entry);
                CheckHorizontalTabs(entry);
                if (SyncSheets(entry, true)) entry.Failures = 0;
                else if (++entry.Failures > 20) DropIfWindowGone(entry);
            }
        }

        // Excel raises no event when the user shows or hides the horizontal tabs (the option "show sheet tabs"), so the vertical pane
        // follows it from here: shown while they are shown, hidden while they are not.
        private void CheckHorizontalTabs(Entry entry)
        {
            if (entry.Busy || entry.ProtectAsking || entry.Syncing) return;
            try
            {
                if (WantsVertical(entry) != entry.WantVisible) ApplyMode(entry);
            }
            catch (Exception)
            {
                // The window is closing; the sync notices.
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindow(IntPtr hWnd);

        // Excel did not answer for a while. That is also so while it calculates or runs a macro for long, and the entry must survive that:
        // it holds what is not in the workbook yet (folders that were made, the window's display mode). The entry goes only when the
        // window itself is gone (closed without our noticing).
        private void DropIfWindowGone(Entry entry)
        {
            bool alive;
            try { alive = entry.Hwnd != 0 && IsWindow(new IntPtr(entry.Hwnd)); }
            catch (Exception) { alive = false; }
            if (alive) entry.Failures = 0; // (asked again after another round of failures)
            else PurgeEntry(entry);
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
            NotifyBadgeEvents();
        }

        // Returns false when Excel could not be asked (busy, or the window/workbook is gone).
        // dialogChecked: the caller has just found no dialog of Excel open (the timer's sync of all windows).
        private static bool SyncSheets(Entry entry, bool dialogChecked = false)
        {
            if (entry.Busy) return true;
            // While Excel's "Protect Sheet" dialog is open the sync stays out of Excel.
            if (entry.ProtectAsking) return true;
            // While one of Excel's dialogs is open (Excel disables its main window), the sync stays out of Excel: some questions it
            // asks (the workbook's theme colors, for one) are answered with the message "this command cannot be used on a protected
            // sheet" as soon as the workbook has a protected sheet.
            if (!dialogChecked && ExcelMainWindowDisabled(entry)) return true;
            // A sync that is already running is not started again: the moves the sync makes itself (the
            // sheet-name rule) make Excel raise SheetActivate, whose handler would start a sync inside
            // the running one, which sees the sheets half moved and moves more; and Excel delivers timer
            // ticks while a call into it is running. The outer sync reads again at the next tick.
            if (entry.Syncing) return true;
            entry.Syncing = true;
            // The sync only reads. Should anything in it still change the workbook (an id given to a new
            // sheet, Excel's own reaction to a read), a workbook that was unchanged stays unchanged: opening
            // a workbook must never make Excel ask "save changes?". The sync runs to its end on the UI
            // thread, so a change of the flag within it is never the user's.
            bool wasSaved = false;
            try { wasSaved = entry.Workbook.Saved; } catch (Exception) { }
            try
            {
                return SyncSheetsCore(entry);
            }
            finally
            {
                entry.Syncing = false;
                // (a move of sheets by the name rule is a real change, see EnforceNameGrouping)
                try { if (wasSaved && !entry.RuleChanged && !entry.Workbook.Saved) entry.Workbook.Saved = true; } catch (Exception) { }
                entry.RuleChanged = false;
            }
        }

        private static bool SyncSheetsCore(Entry entry)
        {
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
                if (RuleOn(entry) && EnforceNameGrouping(entry, allNames, true)) return true; // read again at the next tick
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

                bool activeProtected = IsSheetProtected(activeSheet);
                entry.Control.ActiveSheetProtected = activeProtected;
                // The rows first: the badges may take a while to look at a sheet, and the history uses the keys the folders read.
                entry.TickKeys = null;
                List<StripRow> rows = BuildRows(entry, allNames, visibleAt);
                entry.Control.SetSheets(names, active, selected, colors, rows);
                // a change of the rows is drawn now: the badges below may take a while (the sync would draw only at its end)
                if (entry.Settings.Badges != BadgeKinds.None) entry.Control.PaintNow();
                TrackHistory(entry, activeSheet, active, allNames);
                UpdateBadges(entry, names, active, activeProtected);
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
            bool rule = RuleOn(entry);
            if (rule)
            {
                // the folders are the ones the names describe
                if (!allNames.Any(NameRule.IsFolderSheet))
                {
                    if (entry.Folders.HasFolders) entry.Folders = new FolderModel();
                    return null;
                }
            }
            else if (!entry.Folders.HasFolders) return null;

            string signature = string.Join("|", allNames);
            if (signature != entry.KeySignature || entry.Keys.Count != allNames.Count)
            {
                bool complete;
                List<string> keys = FolderStore.SheetKeys(entry.Workbook, true, true, out complete);
                entry.TickKeys = keys;
                entry.TickKeysComplete = complete;
                entry.TickSignature = signature;
                if (!complete && ++entry.KeyMisses < 3)
                {
                    // A sheet could not be read (Excel is busy): a guessed key would take the sheet out of its
                    // folder and, back with its real key, make it look new (and colored by a folder that keeps
                    // its color). The model stays as it is and the keys are read again at the next tick.
                    entry.KeySignature = null;
                    if (entry.Keys.Count != allNames.Count) return null;
                }
                else
                {
                    entry.KeyMisses = 0;
                    bool guessed = entry.KeysGuessed;
                    entry.KeysGuessed = !complete;
                    entry.Keys = keys;
                    entry.AllNames = new List<string>(allNames);
                    entry.KeySignature = signature;
                    if (rule) entry.Folders = NameRule.BuildModel(allNames, entry.Keys, entry.Folders);
                    else entry.Folders.Reconcile(entry.Keys);
                    // keys that were (or are) partly guessed cannot tell a new sheet from a known one: the
                    // folders that keep their color only note their sheets this time
                    if (guessed || !complete) entry.SyncSeen.Clear();
                    ApplySyncColors(entry);
                }
            }
            var nameOf = new Dictionary<string, string>();
            var nameOfAll = new Dictionary<string, string>();
            for (int i = 0; i < entry.Keys.Count && i < allNames.Count; i++)
            {
                nameOfAll[entry.Keys[i]] = allNames[i];
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
                    ColorBgr = FolderColorBgr(entry, node.Color),
                    SyncColor = node.SyncColor
                };
                if (rule)
                {
                    // the name of the folder and the names of its sheets share 28 characters
                    int longest = 0;
                    foreach (string code in node.Sheets)
                    {
                        string full;
                        if (nameOfAll.TryGetValue(code, out full)) longest = Math.Max(longest, NameRule.BaseOf(full).Length);
                    }
                    folderRow.EditLimit = Math.Max(1, NameRule.MaxCombined - longest);
                }
                rows.Add(folderRow);
                foreach (string code in node.Sheets)
                {
                    if (!nameOf.TryGetValue(code, out name)) continue;
                    folderRow.Members.Add(name);
                    var sheetRow = new StripRow { Sheet = name, Indented = true };
                    if (rule)
                    {
                        sheetRow.Display = NameRule.BaseOf(name);
                        sheetRow.EditLimit = Math.Max(1, NameRule.MaxCombined - node.Name.Length);
                    }
                    folderRow.Children.Add(sheetRow); // (the search looks at the sheets of closed folders as well)
                    if (node.Expanded) rows.Add(sheetRow);
                }
            }
            return rows;
        }

        // ---- the badges of the rows: errors and notes ----

        // How long a sheet has to be left alone (no edit, no recalculation) before it is looked at again, and how long one tick may look.
        private static readonly TimeSpan BadgeQuiet = TimeSpan.FromMilliseconds(500);
        private const int BadgeBudgetMs = 30;
        // A sheet is looked at again after this many seconds (the events cannot tell of every change: a note, the protection of a sheet).
        private const int BadgeRefreshSeconds = 10;

        // Called by the sync. With no badge chosen it does nothing at all. The sheets are looked at one after the other within a small
        // time budget per tick, so that a workbook with many sheets does not stop Excel, and a sheet that is being edited is left alone
        // until it has been quiet for a while.
        private static void UpdateBadges(Entry entry, List<string> visibleNames, string activeName, bool activeProtected)
        {
            try
            {
                if (entry.Settings.Badges == BadgeKinds.None)
                {
                    if (entry.Badges.Count > 0 || entry.BadgeDirty.Count > 0)
                    {
                        entry.Badges.Clear();
                        entry.BadgeDirty.Clear();
                        entry.Control.SetBadges(null);
                    }
                    return;
                }
                var wanted = new HashSet<string>(visibleNames);
                foreach (string gone in entry.Badges.Keys.Where(k => !wanted.Contains(k)).ToList()) entry.Badges.Remove(gone);
                foreach (string gone in entry.BadgeDirty.Keys.Where(k => !wanted.Contains(k)).ToList()) entry.BadgeDirty.Remove(gone);
                foreach (string n in visibleNames)
                {
                    if (!entry.Badges.ContainsKey(n) && !entry.BadgeDirty.ContainsKey(n)) entry.BadgeDirty[n] = DateTime.MinValue;
                }
                foreach (string gone in entry.BadgeScanned.Keys.Where(k => !wanted.Contains(k)).ToList()) { entry.BadgeScanned.Remove(gone); entry.BadgeCost.Remove(gone); }
                var clock = System.Diagnostics.Stopwatch.StartNew();
                DateTime now = DateTime.UtcNow;
                BadgeKinds kinds = entry.Settings.Badges;

                // 1. the active sheet: its notes and its protection are cheap to read, so they are read at every tick (no event tells of them)
                BadgeInfo activeInfo;
                if (activeName != null && entry.Badges.TryGetValue(activeName, out activeInfo))
                {
                    activeInfo.Protected = activeProtected;
                    int memos;
                    if ((kinds & BadgeKinds.Memo) != 0 && ReadMemoCount(entry.Workbook, activeName, out memos)) activeInfo.Memos = memos;
                }

                // 2. the sheets that were changed, calculated or left (or never looked at): when they have been quiet for a short while,
                // oldest first; a sheet that took long to look at is not looked at again within twenty times that long
                foreach (string n in entry.BadgeDirty.Where(kv => now - kv.Value >= BadgeQuiet).OrderBy(kv => kv.Value).Select(kv => kv.Key).ToList())
                {
                    if (clock.ElapsedMilliseconds > BadgeBudgetMs) break;
                    int cost;
                    DateTime last;
                    if (entry.BadgeDirty[n] != DateTime.MinValue && entry.BadgeCost.TryGetValue(n, out cost) && entry.BadgeScanned.TryGetValue(n, out last)
                        && (now - last).TotalMilliseconds < cost * 20) continue;
                    ScanOne(entry, n, now);
                }

                // 3. the other sheets, now and then: whatever a missed event left behind is put right within a short time
                if (clock.ElapsedMilliseconds < BadgeBudgetMs / 2)
                {
                    foreach (string n in visibleNames
                        .Where(v => !entry.BadgeDirty.ContainsKey(v) && (!entry.BadgeScanned.ContainsKey(v) || (now - entry.BadgeScanned[v]).TotalSeconds >= BadgeRefreshSeconds))
                        .OrderBy(v => entry.BadgeScanned.ContainsKey(v) ? entry.BadgeScanned[v] : DateTime.MinValue).ToList())
                    {
                        if (clock.ElapsedMilliseconds > BadgeBudgetMs) break;
                        ScanOne(entry, n, now);
                    }
                }
                entry.Control.SetBadges(entry.Badges);
            }
            catch (Exception)
            {
                // Excel is busy; the next tick tries again.
            }
        }

        private static void ScanOne(Entry entry, string name, DateTime now)
        {
            var clock = System.Diagnostics.Stopwatch.StartNew();
            BadgeInfo info;
            if (!ScanBadges(entry.Workbook, name, entry.Settings.Badges, out info)) return; // Excel is busy: the next tick
            entry.Badges[name] = info;
            entry.BadgeDirty.Remove(name);
            entry.BadgeScanned[name] = now;
            entry.BadgeCost[name] = (int)clock.ElapsedMilliseconds;
        }

        // ---- the progress of a sheet: a custom property of the sheet, "todo" / "doing" / "done" ----

        private const string ProgressProperty = "VTab.Progress";

        private static int ReadProgress(Excel.Worksheet ws)
        {
            try
            {
                foreach (Excel.CustomProperty cp in ws.CustomProperties)
                {
                    if (cp.Name != ProgressProperty) continue;
                    string v = Convert.ToString(cp.Value);
                    return v == "todo" ? 1 : v == "doing" ? 2 : v == "done" ? 3 : 0;
                }
            }
            catch (Exception)
            {
            }
            return 0;
        }

        private static void WriteProgress(Excel.Worksheet ws, int state)
        {
            for (int i = ws.CustomProperties.Count; i >= 1; i--)
            {
                Excel.CustomProperty cp = ws.CustomProperties[i];
                if (cp.Name == ProgressProperty) cp.Delete();
            }
            if (state >= 1 && state <= 3) ws.CustomProperties.Add(ProgressProperty, state == 1 ? "todo" : state == 2 ? "doing" : "done");
        }

        // A click on the "done" mark: the sheet is made the active one, and Excel's own command to protect it (or to take its protection off)
        // runs, with its dialog and password. Whether the sheet is protected afterwards is read at the next tick (the active sheet's protection
        // is read at every tick), and the mark follows.
        private void ToggleSheetProtect(Entry entry, string sheet)
        {
            if (entry.ProtectAsking) return; // (the dialog is open already; a second click, or the second click of a double click, does nothing)
            entry.ProtectAsking = true;
            try
            {
                var target = entry.Workbook.Sheets[sheet] as Excel.Worksheet;
                if (target != null && target.ProtectContents) return; // (already protected: the mark never takes the protection off)
                CommitCellEdit(entry);
                // Only this sheet is selected: SelectSheets keeps a group when the sheet is in it, and Excel's "Protect Sheet" on a group that has
                // a protected sheet is refused ("this command cannot be used on a protected sheet").
                if (target != null) target.Select(true);
                else SelectSheets(entry, new List<string> { sheet }, sheet);
                RunExcelCommand(entry, "SheetProtect");
            }
            catch (Exception)
            {
                // Excel is busy.
            }
            finally
            {
                entry.ProtectAsking = false;
            }
        }

        // Sets the progress of sheets (0 = none): the workbook is changed (it is asked to be saved), and the mark shows at once.
        private void SetProgress(Entry entry, List<string> sheets, int state)
        {
            try
            {
                bool changed = false;
                foreach (string name in sheets)
                {
                    var ws = entry.Workbook.Sheets[name] as Excel.Worksheet;
                    if (ws == null || ReadProgress(ws) == state) continue;
                    WriteProgress(ws, state);
                    changed = true;
                    BadgeInfo info;
                    if (!entry.Badges.TryGetValue(name, out info)) { info = new BadgeInfo(); entry.Badges[name] = info; }
                    info.Progress = state;
                }
                if (!changed) return;
                MarkChanged(entry);
                entry.Control.SetBadges(entry.Badges);
            }
            catch (Exception)
            {
                // Excel is busy, or the sheet cannot take a property.
            }
        }

        // The number of notes and comments of a sheet; false when Excel could not answer now.
        private static bool ReadMemoCount(Excel.Workbook workbook, string sheetName, out int count)
        {
            count = 0;
            try
            {
                var ws = workbook.Sheets[sheetName] as Excel.Worksheet;
                if (ws == null) return false;
                count = ws.Comments.Count;
                try { count += ws.CommentsThreaded.Count; } catch (Exception) { }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // A sheet whose used range has more cells than this has its errors not counted (no badge for them).
        private const double BadgeErrorMaxCells = 2000000;

        // Counts the cells with an error (of formulas) and the notes and comments of one sheet. False when Excel could not answer now.
        private static bool ScanBadges(Excel.Workbook workbook, string sheetName, BadgeKinds kinds, out BadgeInfo info)
        {
            info = new BadgeInfo();
            var ws = workbook.Sheets[sheetName] as Excel.Worksheet;
            if (ws == null) return true; // a chart sheet: nothing to tell
            try
            {
                try { info.Protected = ws.ProtectContents; } catch (Exception) { }
                if ((kinds & BadgeKinds.Progress) != 0) info.Progress = ReadProgress(ws);
                if ((kinds & BadgeKinds.Error) != 0)
                {
                    // The formulas that show an error, counted by one formula over the used range: the time follows the number of cells
                    // (about 60 ms for a million). SpecialCells would be quicker for a few errors, but it takes seconds, up to minutes, when
                    // the errors are scattered (50,000 of them in 100,000 cells took 6 s), and cannot be stopped.
                    Excel.Range used = null;
                    try
                    {
                        used = ws.UsedRange;
                        // (a used range that spans the whole sheet is not counted: its cells would have to be held in memory)
                        if (Convert.ToDouble(used.CountLarge) <= BadgeErrorMaxCells)
                        {
                            string address = used.get_Address(false, false, Excel.XlReferenceStyle.xlA1, Type.Missing, Type.Missing);
                            object count = ws.Evaluate("SUMPRODUCT(--ISERROR(" + address + "),--ISFORMULA(" + address + "))");
                            if (count is double && (double)count >= 0) info.Errors = (int)Math.Min(int.MaxValue, (double)count);
                        }
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        return false; // Excel is busy or editing: tried again
                    }
                    finally
                    {
                        if (used != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(used);
                    }
                }
                if ((kinds & BadgeKinds.Memo) != 0)
                {
                    info.Memos = ws.Comments.Count;
                    try { info.Memos += ws.CommentsThreaded.Count; } catch (Exception) { } // (older Excel has no threaded comments)
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // Called by the Excel events (a cell changed, a sheet calculated or left): the sheet is looked at again when it has been quiet.
        public void MarkBadgesDirty(object sheet, bool atOnce = false)
        {
            if (!_badgeEventsOn) return; // (before any call into Excel)
            try
            {
                var ws = sheet as Excel.Worksheet;
                if (ws == null) return;
                string name = ws.Name;
                var book = ws.Parent as Excel.Workbook;
                string bookName = book != null ? book.Name : null;
                foreach (Entry entry in _entries.Values)
                {
                    if (entry.Settings.Badges == BadgeKinds.None) continue;
                    if (bookName != null && entry.Workbook.Name != bookName) continue;
                    entry.BadgeDirty[name] = atOnce ? DateTime.MinValue : DateTime.UtcNow; // (a sheet that is left is looked at at once)
                }
            }
            catch (Exception)
            {
            }
        }

        // ---- the sheets seen last ----

        // Called by the sync: a sheet that became active (by a click, a key, a macro, the tabs of Excel) is put into the history;
        // when sheets were added, renamed or deleted, the sheets that no longer exist are taken out.
        private static void TrackHistory(Entry entry, object activeSheet, string activeName, List<string> allNames)
        {
            try
            {
                if (activeName != null && activeSheet != null && activeName != entry.HistActive)
                {
                    entry.HistActive = activeName;
                    string key = FolderStore.KeyOf(activeSheet, true);
                    entry.History.Push(key);
                    if (key != null) entry.HistHints[key] = activeName;
                }
                string namesSignature = string.Join("|", allNames);
                if (namesSignature != entry.HistNames)
                {
                    entry.HistNames = namesSignature;
                    if (entry.History.Keys.Count > 0)
                    {
                        // The keys the folders read in this sync are used; else, when every sheet of the history is still where it was seen
                        // (its name has its key), no sheet of the history is gone and nothing is read; else all keys are read, as before.
                        bool complete = false;
                        List<string> keys;
                        if (entry.TickKeys != null && entry.TickSignature == namesSignature)
                        {
                            keys = entry.TickKeys;
                            complete = entry.TickKeysComplete;
                        }
                        else if (HistoryInPlace(entry, allNames)) keys = null;
                        else keys = FolderStore.SheetKeys(entry.Workbook, false, true, out complete);
                        if (keys != null && complete)
                        {
                            entry.History.Prune(new HashSet<string>(keys));
                            NoteHistoryHints(entry, keys, allNames);
                        }
                    }
                }
                entry.Control.SetNavigation(entry.History.CanBack, entry.History.CanForward);
            }
            catch (Exception)
            {
                // Excel is busy; the next tick tries again.
            }
        }

        // True when every sheet of the history is found under the name it was last seen with, with its key: then none of them is gone.
        // False when one of them could not be found that way (renamed, deleted, made again under the same name, or not seen yet since the
        // workbook was opened); the caller then reads all keys. With few sheets, reading all is as quick, and is done.
        private static bool HistoryInPlace(Entry entry, List<string> allNames)
        {
            if (entry.History.Keys.Count >= allNames.Count) return false;
            try
            {
                var names = new HashSet<string>(allNames);
                foreach (string key in entry.History.Keys)
                {
                    if (key.StartsWith("n:", StringComparison.Ordinal))
                    {
                        // a chart sheet's key is its name
                        string chartName = key.Substring(2);
                        if (!names.Contains(chartName) || !(entry.Workbook.Sheets[chartName] is Excel.Chart)) return false;
                        continue;
                    }
                    string hint;
                    if (!entry.HistHints.TryGetValue(key, out hint) || !names.Contains(hint)) return false;
                    var ws = entry.Workbook.Sheets[hint] as Excel.Worksheet;
                    if (ws == null || FolderStore.KeyOf(ws, false) != key) return false;
                }
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // The names the sheets of the history have now (keys and names: all sheets in Excel's order).
        private static void NoteHistoryHints(Entry entry, List<string> keys, List<string> allNames)
        {
            var hints = new Dictionary<string, string>();
            var inHistory = new HashSet<string>(entry.History.Keys);
            for (int i = 0; i < keys.Count && i < allNames.Count; i++)
            {
                if (inHistory.Contains(keys[i])) hints[keys[i]] = allNames[i];
            }
            entry.HistHints = hints;
        }

        // The visible sheets by key (what the history can go to).
        private static Dictionary<string, string> VisibleSheetsByKey(Entry entry)
        {
            var map = new Dictionary<string, string>();
            List<string> keys = FolderStore.SheetKeys(entry.Workbook, false);
            int i = 0;
            foreach (object sheet in entry.Workbook.Sheets)
            {
                string name;
                if (IsVisibleSheet(sheet, out name) && i < keys.Count) map[keys[i]] = name;
                i++;
            }
            return map;
        }

        // 1 = back (older), -1 = forward (newer)
        private void NavigateHistory(Entry entry, int direction)
        {
            try
            {
                Dictionary<string, string> map = VisibleSheetsByKey(entry);
                int at = entry.History.Find(direction, map.ContainsKey);
                if (at < 0) return;
                GoToHistory(entry, at, map);
            }
            catch (Exception)
            {
            }
        }

        // "履歴をクリア": every sheet is taken out; the sheet shown now is entered again by the next sync, so "back" works from here on.
        private void ClearHistory(Entry entry)
        {
            entry.History.Keys.Clear();
            entry.History.Cursor = 0;
            entry.HistActive = null;
            entry.Control.SetNavigation(false, false);
        }

        private void JumpToHistory(Entry entry, string key)
        {
            try
            {
                Dictionary<string, string> map = VisibleSheetsByKey(entry);
                int at = entry.History.Keys.IndexOf(key);
                if (at < 0 || !map.ContainsKey(key)) return;
                GoToHistory(entry, at, map);
            }
            catch (Exception)
            {
            }
        }

        private void GoToHistory(Entry entry, int at, Dictionary<string, string> map)
        {
            string key = entry.History.Keys[at];
            string name = map[key];
            CommitCellEdit(entry);
            entry.History.Cursor = at; // the sync then finds the active sheet where the history stands: nothing is added
            // a sheet in a closed folder: the folder is opened to show it
            FolderNode folder = entry.Folders.Nodes.Find(n => n.IsFolder && !n.Expanded && n.Sheets.Contains(key));
            if (folder != null) ToggleFolder(entry, folder.FolderId);
            SelectSheets(entry, new List<string> { name }, name);
            entry.Control.SetNavigation(entry.History.CanBack, entry.History.CanForward);
            FocusGrid(entry, true);
        }

        private static List<HistoryItem> HistoryItems(Entry entry)
        {
            var items = new List<HistoryItem>();
            try
            {
                Dictionary<string, string> map = VisibleSheetsByKey(entry);
                for (int i = 0; i < entry.History.Keys.Count && items.Count < 10; i++)
                {
                    string name;
                    if (map.TryGetValue(entry.History.Keys[i], out name))
                        items.Add(new HistoryItem { Key = entry.History.Keys[i], Name = name, Current = i == entry.History.Cursor });
                }
            }
            catch (Exception)
            {
            }
            return items;
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

        // ---- folder commands ----

        // Reads the sheets of the workbook now, gives them ids and brings the model in line with them.
        private static void RefreshKeys(Entry entry)
        {
            List<string> names = ReadNames(entry);
            if (RuleOn(entry) && EnforceNameGrouping(entry, names, false)) names = ReadNames(entry);
            entry.Keys = FolderStore.SheetKeys(entry.Workbook, true);
            entry.AllNames = names;
            entry.KeySignature = string.Join("|", names);
            if (RuleOn(entry)) entry.Folders = NameRule.BuildModel(names, entry.Keys, entry.Folders);
            else entry.Folders.Reconcile(entry.Keys);
        }

        // ---- the sheet-name rule ----

        private static bool RuleOn(Entry entry)
        {
            return entry.Settings != null && entry.Settings.NameRule;
        }

        private static List<string> ReadNames(Entry entry)
        {
            var names = new List<string>();
            foreach (object sheet in entry.Workbook.Sheets) names.Add(SheetName(sheet) ?? "?");
            return names;
        }

        // A sheet "A | D" that is not next to its folder goes to the right of the nearest sheet to its
        // left that starts with "A | ". names is the current order of all sheets; true when sheets moved.
        // In the sync (gated = true) the check is made only when the names changed.
        private static bool EnforceNameGrouping(Entry entry, List<string> names, bool gated)
        {
            string signature = string.Join("|", names);
            if (gated)
            {
                if (signature == entry.RuleSignature) return false;
                entry.RuleSignature = signature;
            }
            List<KeyValuePair<int, int>> moves = NameRule.GroupingMoves(new List<string>(names));
            if (moves.Count == 0)
            {
                if (gated) entry.QuietRegroup = false; // the order holds: from now on a move is a real change
                return false;
            }
            int moved = 0;
            bool screen = true;
            try
            {
                screen = entry.Workbook.Application.ScreenUpdating;
                entry.Workbook.Application.ScreenUpdating = false;
                foreach (KeyValuePair<int, int> m in moves)
                {
                    object sheet = entry.Workbook.Sheets[m.Key + 1];
                    object before = entry.Workbook.Sheets[m.Value + 1];
                    var ws = sheet as Excel.Worksheet;
                    if (ws != null) ws.Move(before);
                    else
                    {
                        var ch = sheet as Excel.Chart;
                        if (ch != null) ch.Move(before);
                    }
                    moved++;
                }
            }
            catch (Exception)
            {
                // Structure protected, or Excel is busy: the sheets stay where they are.
            }
            finally
            {
                try { entry.Workbook.Application.ScreenUpdating = screen; } catch (Exception) { }
            }
            // Nothing could be moved (a protected structure): the workbook is not changed, and the same order
            // is not tried again at every tick (RuleSignature keeps it; a change of the names tries again).
            if (moved == 0)
            {
                if (gated) entry.QuietRegroup = false;
                return false;
            }
            entry.RuleSignature = null;
            entry.KeySignature = null;
            // Putting a just opened workbook in order is not the user's change: the sync keeps it unchanged.
            if (gated && entry.QuietRegroup) return true;
            entry.RuleChanged = true;
            MarkChanged(entry);
            return true;
        }

        // The names the sheets get from the folders of the model: "Folder | Sheet" inside a folder, the
        // sheet's own name outside. Null when done, else why it cannot be done (nothing is renamed then).
        private static string ApplyNamesFromModel(Entry entry)
        {
            var target = new Dictionary<string, string>();
            foreach (FolderNode node in entry.Folders.Nodes)
            {
                if (node.IsFolder)
                {
                    foreach (string key in node.Sheets)
                    {
                        int i = entry.Keys.IndexOf(key);
                        if (i >= 0 && i < entry.AllNames.Count) target[key] = NameRule.Compose(node.Name, NameRule.BaseOf(entry.AllNames[i]));
                    }
                }
                else
                {
                    int i = entry.Keys.IndexOf(node.SheetCode);
                    if (i >= 0 && i < entry.AllNames.Count) target[node.SheetCode] = NameRule.BaseOf(entry.AllNames[i]);
                }
            }

            var final = new List<string>(entry.AllNames);
            for (int i = 0; i < entry.Keys.Count && i < final.Count; i++)
            {
                string t;
                if (target.TryGetValue(entry.Keys[i], out t)) final[i] = t;
            }
            for (int i = 0; i < final.Count; i++)
            {
                if (final[i].Length > NameRule.MaxSheetName)
                    return "フォルダ名とシート名を合わせた長さが " + NameRule.MaxCombined + " 文字を超えるため、操作できません。\r\n(" + final[i] + ")";
            }
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string n in final)
            {
                if (!seen.Add(n)) return "同じ名前のシートができるため、操作できません。\r\n(" + n + ")";
            }

            try
            {
                var current = new List<string>(entry.AllNames);
                var pending = new List<int>();
                for (int i = 0; i < final.Count; i++) if (!string.Equals(current[i], final[i], StringComparison.Ordinal)) pending.Add(i);
                int guard = 0;
                while (pending.Count > 0 && guard++ < pending.Count * 3 + 6)
                {
                    bool progress = false;
                    foreach (int i in pending.ToList())
                    {
                        bool taken = false;
                        for (int k = 0; k < current.Count; k++)
                        {
                            if (k != i && string.Equals(current[k], final[i], StringComparison.OrdinalIgnoreCase)) { taken = true; break; }
                        }
                        if (taken) continue;
                        SetSheetName(entry.Workbook.Sheets[i + 1], final[i]);
                        current[i] = final[i];
                        pending.Remove(i);
                        progress = true;
                    }
                    if (!progress)
                    {
                        // names that swap places: one goes through a temporary name
                        int i = pending[0];
                        string temp = "~" + Guid.NewGuid().ToString("N").Substring(0, 10);
                        SetSheetName(entry.Workbook.Sheets[i + 1], temp);
                        current[i] = temp;
                    }
                }
                entry.AllNames = final;
            }
            catch (Exception)
            {
                return "シートの名前を変更できませんでした。ブックの構造が保護されている可能性があります。";
            }
            return null;
        }

        private static void SetSheetName(object sheet, string name)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) { ws.Name = name; return; }
            var ch = sheet as Excel.Chart;
            if (ch != null) ch.Name = name;
        }

        // The rule was just turned on: the folders that exist get their sheet names ("Folder | Sheet").
        private void EnableNameRule(Entry entry)
        {
            if (entry.Busy) return;
            entry.QuietRegroup = false;
            entry.Busy = true;
            try
            {
                entry.Settings.NameRule = false; // the model as it is, not as the (still plain) names describe it
                RefreshKeys(entry);
                entry.Settings.NameRule = true;
                if (entry.Folders.HasFolders)
                {
                    string problem = ApplyNamesFromModel(entry);
                    if (problem != null)
                    {
                        entry.Settings.NameRule = false;
                        ShowMessage("シート名のルールを使えません。\r\n" + problem);
                        return;
                    }
                }
                MarkChanged(entry);
            }
            catch (Exception)
            {
                entry.Settings.NameRule = false;
                ShowMessage("シート名のルールを使えません。");
            }
            finally
            {
                entry.Busy = false;
                entry.KeySignature = null;
                entry.RuleSignature = null;
            }
            SyncSheets(entry);
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
        private void ChangeFolders(Entry entry, Func<bool> change, IList<string> selectedNames, string activeName, string activateKey = null)
        {
            if (entry.Busy) return;
            entry.Busy = true;
            try
            {
                RefreshKeys(entry);
                // the selection and the active sheet are followed by their keys: the sheets may get new names with their folders
                List<string> selectedKeys = selectedNames == null ? null : KeysOf(entry, selectedNames);
                string activeKey = activeName == null ? null : KeyOf(entry, activeName);
                string snapshot = entry.Folders.ToXml();
                if (!change()) return;
                if (RuleOn(entry))
                {
                    // the sheets are renamed after their folders; when that cannot be done, nothing changes
                    string problem = ApplyNamesFromModel(entry);
                    if (problem != null)
                    {
                        entry.Folders = FolderModel.Parse(snapshot);
                        entry.KeySignature = null;
                        ShowMessage(problem);
                        return;
                    }
                }
                if (activateKey != null)
                {
                    // the sheet that was dragged is selected and active at the end of the move, not the one that was active before
                    // (otherwise the old one would show for a moment); its name may have changed with its folder
                    // (read straight from Excel: RefreshKeys would take the model back to the old sheet order)
                    List<string> keysNow = FolderStore.SheetKeys(entry.Workbook, true);
                    List<string> namesNow = ReadNames(entry);
                    int at = keysNow.IndexOf(activateKey);
                    if (at >= 0 && at < namesNow.Count)
                    {
                        activeName = namesNow[at];
                        selectedNames = new List<string> { activeName };
                    }
                }
                else if (selectedKeys != null && selectedKeys.Count > 0 && activeKey != null)
                {
                    // the same sheets stay selected, and the same one active, under their names after the change
                    // (read straight from Excel: RefreshKeys would take the model back to the old sheet order)
                    List<string> keysNow = FolderStore.SheetKeys(entry.Workbook, true);
                    List<string> namesNow = ReadNames(entry);
                    var selectedNow = new List<string>();
                    foreach (string k in selectedKeys)
                    {
                        int at = keysNow.IndexOf(k);
                        if (at >= 0 && at < namesNow.Count) selectedNow.Add(namesNow[at]);
                    }
                    int activeAt = keysNow.IndexOf(activeKey);
                    if (selectedNow.Count > 0 && activeAt >= 0 && activeAt < namesNow.Count)
                    {
                        selectedNames = selectedNow;
                        activeName = namesNow[activeAt];
                    }
                }
                ApplyExcelOrder(entry, selectedNames, activeName);
                if (activateKey != null && activeName != null) SelectSheets(entry, selectedNames, activeName);
                MarkChanged(entry);
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
            // a single sheet that is dragged becomes the active sheet (as it does on Excel's own tabs); its name may change with its folder
            string draggedKey = sheets.Count == 1 && folders.Count == 0 ? KeyOf(entry, sheets[0]) : null;
            ChangeFolders(entry, () =>
            {
                List<string> keys = KeysOf(entry, sheets);
                string beforeKey = beforeSheet == null ? null : KeyOf(entry, beforeSheet);
                return entry.Folders.Move(keys, folders, into, beforeKey, beforeFolder);
            }, selected, active, draggedKey);
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
                var namesBefore = new HashSet<string>(entry.AllNames, StringComparer.OrdinalIgnoreCase);
                CopySheetsTo(entry, sourceNames, anchorKey == null ? null : entry.AllNames[oldKeys.IndexOf(anchorKey)]);
                if (RuleOn(entry) && folders.Count == 0) NameCopiesForDrop(entry, namesBefore, into);

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
                                unit.SyncColor = src.SyncColor;
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
                if (RuleOn(entry))
                {
                    // the copies are named after the folders they went to; when that cannot be done they
                    // keep the names Excel gave them
                    string problem = ApplyNamesFromModel(entry);
                    if (problem != null) ShowMessage(problem);
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

        // The sheet-name rule: the copies Excel just made are named for the place they were dropped at
        // ("Folder | Sheet" in a folder, the sheet's own name outside), before the grouping sees them.
        private static void NameCopiesForDrop(Entry entry, HashSet<string> namesBefore, string intoFolderId)
        {
            try
            {
                FolderNode target = intoFolderId == null ? null : entry.Folders.FindFolder(intoFolderId);
                List<string> now = ReadNames(entry);
                for (int i = 0; i < now.Count; i++)
                {
                    if (namesBefore.Contains(now[i])) continue;
                    string wanted = target == null ? NameRule.BaseOf(now[i]) : NameRule.Compose(target.Name, NameRule.BaseOf(now[i]));
                    if (wanted == now[i] || wanted.Length > NameRule.MaxSheetName) continue;
                    if (now.Any(n => string.Equals(n, wanted, StringComparison.OrdinalIgnoreCase))) continue;
                    SetSheetName(entry.Workbook.Sheets[i + 1], wanted);
                    now[i] = wanted;
                }
            }
            catch (Exception)
            {
                // the copies keep the names Excel gave them
            }
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

        // Bulk commands on the sheets of folders. Excel's own commands work on the
        // selected sheets, so the visible sheets of the folders are selected first.
        private void FolderCommand(Entry entry, string command, IList<string> folderIds)
        {
            try
            {
                RefreshKeys(entry);
                if (command == "deselect")
                {
                    DeselectFolderSheets(entry, folderIds);
                    return;
                }
                if (command == "closeall" || command == "openall")
                {
                    // all folders of the pane, whatever is selected (the buttons "全開" and "全閉"; the folder menu no longer has these)
                    bool open = command == "openall";
                    bool changed = false;
                    foreach (FolderNode node in entry.Folders.Nodes)
                    {
                        if (!node.IsFolder || node.Expanded == open) continue;
                        node.Expanded = open;
                        changed = true;
                    }
                    if (changed) MarkChanged(entry);
                    SyncSheets(entry);
                    return;
                }
                if (command == "synccolor")
                {
                    ToggleSyncColor(entry, folderIds);
                    SyncSheets(entry);
                    return;
                }
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
                }
            }
            catch (Exception)
            {
                ShowMessage("フォルダ内のシートの操作を完了できませんでした。");
            }
            SyncSheets(entry);
        }

        // "Deselect all sheets of the folder": when sheets outside the folders are selected, all sheets of
        // the folders are taken out of the selection. Otherwise (nothing outside is selected, Excel needs
        // one selected sheet) the first selected sheet of each folder, in the folder's order, stays and the
        // others are taken out.
        private void DeselectFolderSheets(Entry entry, IList<string> folderIds)
        {
            IList<string> selection = CurrentSelection(entry);
            string active = CurrentActive(entry);
            if (selection.Count == 0 || active == null) return;
            var selected = new HashSet<string>(selection);

            // Sheets outside the folders are selected: they are what stays selected, all sheets of the
            // folders are taken out.
            var members = new HashSet<string>(MemberNames(entry, folderIds));
            List<string> outside = selection.Where(n => !members.Contains(n)).ToList();
            if (outside.Count > 0)
            {
                if (outside.Count == selection.Count) return; // no sheet of the folders is selected
                entry.Window.Activate();
                SelectSheets(entry, outside, members.Contains(active) ? outside[0] : active);
                return;
            }

            var keptFor = new Dictionary<string, string>(); // removed sheet -> the sheet of its folder that stays
            foreach (string id in folderIds)
            {
                string kept = null;
                foreach (string name in MemberNames(entry, new[] { id }))
                {
                    if (!selected.Contains(name)) continue;
                    if (kept == null) kept = name;
                    else keptFor[name] = kept;
                }
            }
            if (keptFor.Count == 0) return;
            var remaining = selection.Where(n => !keptFor.ContainsKey(n)).ToList();
            string newActive;
            if (!keptFor.TryGetValue(active, out newActive)) newActive = active;
            entry.Window.Activate();
            SelectSheets(entry, remaining, newActive);
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
            if (RuleOn(entry))
            {
                if (NameRule.HasSeparator(newName))
                {
                    ShowMessage("「" + NameRule.Separator + "」を含む名前は、縦タブでは入力できません。");
                    return;
                }
                // the sheets of the folder are renamed with it
                ChangeFolders(entry, () =>
                {
                    FolderNode f = entry.Folders.FindFolder(folderId);
                    if (f == null) return false;
                    f.Name = newName;
                    return true;
                }, null, null);
                return;
            }
            folder.Name = newName;
            MarkChanged(entry);
            SyncSheets(entry);
        }

        // The folders go, their sheets stay where they are (outside any folder).
        private void DeleteFoldersOnly(Entry entry, IList<string> folderIds)
        {
            if (RuleOn(entry))
            {
                // the sheets lose the "Folder | " of their names
                ChangeFolders(entry, () =>
                {
                    foreach (string id in folderIds) entry.Folders.RemoveFolder(id);
                    return true;
                }, null, null);
                return;
            }
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
            // The sync asks at every tick for the folders that are colored by a theme color; the workbook's theme is read again after a few seconds only
            if (entry.ThemeCache != null && DateTime.Now - entry.ThemeCacheAt < TimeSpan.FromSeconds(entry.ThemeCacheSeconds)) return entry.ThemeCache;
            // Excel gives no theme (Workbook.Theme is null) while any sheet of the workbook is protected. The colors
            // are then taken from the built-in cell styles, which Excel answers all the same (9 of the 10 colors; the live theme), and the
            // rest from the saved file's theme, or, when the workbook was never saved, from the Office theme. Never null, so that a folder
            // colored with a theme color shows that color, and the palette shows the same colors as the folders get.
            var colors = ReadThemeColors(entry);
            int seconds = 5;
            if (colors == null)
            {
                colors = ReadThemeColorsFromFile(entry) ?? TabColorPalette.DefaultTheme();
                OverrideFromStyles(entry, colors);
                seconds = 30; // (reading the styles takes longer)
            }
            entry.ThemeCache = colors;
            entry.ThemeCacheAt = DateTime.Now;
            entry.ThemeCacheSeconds = seconds;
            return colors;
        }

        // The theme colors that the built-in cell styles show: the accent styles' fills are the accent colors; the font colors of the styles
        // "20% - Accent 1" (dark 1), "Accent 1" (light 1) and "Title" (dark 2) are theme colors 2, 1 and 4 without a tint. Background 2
        // (light 2) is not used by any built-in style. Found by the theme index of the style, not by its (translated) name.
        private static void OverrideFromStyles(Entry entry, System.Drawing.Color[] colors)
        {
            // The styles that gave the 9 colors the last time are read directly; when one of them no longer gives its color, all styles are
            // looked at again (any built-in style with that theme color and no tint shows the same color, so the result is the same).
            if (entry.StyleSources.Count >= 9 && FromKnownStyles(entry, colors)) return;
            try
            {
                // position in colors[]: 0 light1, 1 dark1, 2 light2, 3 dark2, 4..9 accent1..6; Excel's theme index of the style: font 1 light1, 2 dark1, 4 dark2; fill 5..10 accents
                var found = new HashSet<int>();
                var sources = new Dictionary<int, string>();
                foreach (Excel.Style style in entry.Workbook.Styles)
                {
                    try
                    {
                        if (!style.BuiltIn) continue;
                        int fillIndex = style.Interior.ThemeColor;
                        if (fillIndex >= 5 && fillIndex <= 10 && style.Interior.TintAndShade == 0 && found.Add(fillIndex))
                        {
                            colors[fillIndex - 1] = FromBgr(Convert.ToInt32(style.Interior.Color));
                            sources[fillIndex] = style.Name;
                        }
                        int fontIndex = style.Font.ThemeColor;
                        if ((fontIndex == 1 || fontIndex == 2 || fontIndex == 4) && style.Font.TintAndShade == 0 && found.Add(fontIndex))
                        {
                            colors[fontIndex == 1 ? 0 : fontIndex == 2 ? 1 : 3] = FromBgr(Convert.ToInt32(style.Font.Color));
                            sources[-fontIndex] = style.Name; // (the fonts as negative numbers, apart from the fills' 5..10)
                        }
                    }
                    catch (Exception)
                    {
                        // a style without that color
                    }
                    if (found.Count >= 9) break;
                }
                entry.StyleSources = sources;
            }
            catch (Exception)
            {
                // the colors of the file (or the Office theme) stay
                entry.StyleSources = new Dictionary<int, string>();
            }
        }

        // The colors from the styles noted the last time; false (colors untouched) when one of them does not give its color any more.
        private static bool FromKnownStyles(Entry entry, System.Drawing.Color[] colors)
        {
            var found = new System.Drawing.Color[10];
            var got = new bool[10];
            try
            {
                foreach (KeyValuePair<int, string> source in entry.StyleSources)
                {
                    Excel.Style style = entry.Workbook.Styles[source.Value];
                    if (!style.BuiltIn) return false;
                    int index = source.Key;
                    if (index > 0)
                    {
                        if (style.Interior.ThemeColor != index || style.Interior.TintAndShade != 0) return false;
                        found[index - 1] = FromBgr(Convert.ToInt32(style.Interior.Color));
                        got[index - 1] = true;
                    }
                    else
                    {
                        int fontIndex = -index;
                        if (style.Font.ThemeColor != fontIndex || style.Font.TintAndShade != 0) return false;
                        int at = fontIndex == 1 ? 0 : fontIndex == 2 ? 1 : 3;
                        found[at] = FromBgr(Convert.ToInt32(style.Font.Color));
                        got[at] = true;
                    }
                }
            }
            catch (Exception)
            {
                return false;
            }
            for (int i = 0; i < 10; i++) if (got[i]) colors[i] = found[i];
            return true;
        }

        private static System.Drawing.Color FromBgr(int bgr)
        {
            return System.Drawing.Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF);
        }

        private static System.Drawing.Color[] ReadThemeColors(Entry entry)
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

        // The theme colors of the saved file (xl/theme/theme1.xml of the .xlsx): light1, dark1, light2, dark2, accent1..6; null when
        // the workbook has no file of that kind or the file cannot be read.
        private static System.Drawing.Color[] ReadThemeColorsFromFile(Entry entry)
        {
            // the file is read again only when it changed (saved): the colors of the same file are the same
            string stamp = null;
            try
            {
                string path = entry.Workbook.FullName;
                if (!string.IsNullOrEmpty(path) && System.IO.File.Exists(path))
                {
                    var info = new System.IO.FileInfo(path);
                    stamp = path + "|" + info.LastWriteTimeUtc.Ticks + "|" + info.Length;
                }
            }
            catch (Exception)
            {
                stamp = null;
            }
            if (stamp == null) return null;
            if (stamp != entry.FileThemeStamp)
            {
                entry.FileTheme = ReadThemeFile(entry);
                entry.FileThemeStamp = stamp;
            }
            return entry.FileTheme == null ? null : (System.Drawing.Color[])entry.FileTheme.Clone(); // (a copy: the styles' colors are put into it)
        }

        private static System.Drawing.Color[] ReadThemeFile(Entry entry)
        {
            try
            {
                string path = entry.Workbook.FullName;
                if (string.IsNullOrEmpty(path) || !System.IO.File.Exists(path)) return null;
                using (var stream = new System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite | System.IO.FileShare.Delete))
                using (var zip = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read))
                {
                    var item = zip.GetEntry("xl/theme/theme1.xml");
                    if (item == null) return null;
                    if (item.Length > 2L * SafeXml.ThemeMaxChars) return null; // (a theme is a few KB)
                    using (var reader = item.Open())
                    {
                        var doc = SafeXml.Load(reader, SafeXml.ThemeMaxChars);
                        System.Xml.Linq.XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                        var scheme = doc.Descendants(a + "clrScheme").FirstOrDefault();
                        if (scheme == null) return null;
                        string[] order = { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6" };
                        var colors = new System.Drawing.Color[10];
                        for (int i = 0; i < order.Length; i++)
                        {
                            var node = scheme.Element(a + order[i]);
                            if (node == null) return null;
                            var rgb = node.Element(a + "srgbClr");
                            string hex = rgb != null ? (string)rgb.Attribute("val") : null;
                            if (hex == null)
                            {
                                var sys = node.Element(a + "sysClr");
                                hex = sys != null ? (string)sys.Attribute("lastClr") : null;
                            }
                            if (hex == null || hex.Length != 6) return null;
                            colors[i] = System.Drawing.Color.FromArgb(Convert.ToInt32(hex.Substring(0, 2), 16), Convert.ToInt32(hex.Substring(2, 2), 16), Convert.ToInt32(hex.Substring(4, 2), 16));
                        }
                        return colors;
                    }
                }
            }
            catch (Exception)
            {
                return null;
            }
        }

        // Excel's tab (and cell) colors number the theme colors differently from the XlThemeColor names:
        // 1 is white (Background 1), 2 is black (Text 1), 3 is Background 2 and 4 is Text 2 (measured in
        // Excel; the names say the opposite). The palette and the folder colors use the names' numbering
        // (2 = light 1 ...); this turns it into the number Excel's Tab.ThemeColor wants.
        private static int TabThemeColorIndex(int themeIndex)
        {
            switch (themeIndex)
            {
                case 1: return 2;
                case 2: return 1;
                case 3: return 4;
                case 4: return 3;
                default: return themeIndex;
            }
        }

        // Applies a tab color to every selected sheet, as Excel does for a group.
        private void ApplyTabColor(Entry entry, IList<string> names, TabColorChoice choice)
        {
            try
            {
                SetTabColor(entry, names, choice);
            }
            catch (Exception)
            {
                ShowMessage("シート見出しの色を変更できませんでした。シートが保護されている可能性があります。");
            }
        }

        private static void SetTabColor(Entry entry, IList<string> names, TabColorChoice choice)
        {
            if (choice == null) choice = new TabColorChoice { Kind = TabColorKind.None };
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
                        tab.ThemeColor = (Excel.XlThemeColor)TabThemeColorIndex(choice.ThemeIndex);
                        tab.TintAndShade = choice.Tint;
                        break;
                }
            }
        }

        // Folders that keep their sheets in the folder's color (the check mark): a sheet that is new in such a
        // folder (moved in, added, inserted, copied, or moved there in Excel's own tabs) gets the folder's
        // color. The first look at a folder (the workbook was opened) only notes its sheets. A change of
        // the workbook that the sync makes is real, so it is not taken back (RuleChanged).
        private static void ApplySyncColors(Entry entry)
        {
            bool changed = false;
            try
            {
                var present = new HashSet<string>();
                foreach (FolderNode node in entry.Folders.Nodes)
                {
                    if (!node.IsFolder) continue;
                    present.Add(node.FolderId);
                    if (!node.SyncColor)
                    {
                        entry.SyncSeen.Remove(node.FolderId);
                        continue;
                    }
                    var now = new HashSet<string>(node.Sheets);
                    HashSet<string> seen;
                    if (!entry.SyncSeen.TryGetValue(node.FolderId, out seen))
                    {
                        entry.SyncSeen[node.FolderId] = now;
                        continue;
                    }
                    List<string> arrived = node.Sheets.Where(k => !seen.Contains(k)).ToList();
                    entry.SyncSeen[node.FolderId] = now;
                    if (arrived.Count == 0) continue;

                    var names = new List<string>();
                    foreach (string key in arrived)
                    {
                        int i = entry.Keys.IndexOf(key);
                        if (i >= 0 && i < entry.AllNames.Count) names.Add(entry.AllNames[i]);
                    }
                    if (names.Count == 0) continue;
                    SetTabColor(entry, names, node.Color);
                    changed = true;
                }
                foreach (string id in entry.SyncSeen.Keys.ToList())
                {
                    if (!present.Contains(id)) entry.SyncSeen.Remove(id);
                }
            }
            catch (Exception)
            {
                // A protected sheet or a busy Excel: the sheets keep their colors.
            }
            if (changed)
            {
                entry.RuleChanged = true;
                MarkChanged(entry);
            }
        }

        // The menu item "change all to the folder's color": it switches the check mark of the folders. When
        // it is switched on, all sheets of the folder (hidden ones too) get the folder's color at once; when it
        // is switched off, the sheets keep the colors they have. With several folders chosen: if any of
        // them is off, all are switched on, otherwise all are switched off.
        private void ToggleSyncColor(Entry entry, IList<string> folderIds)
        {
            var folders = new List<FolderNode>();
            foreach (string id in folderIds)
            {
                FolderNode f = entry.Folders.FindFolder(id);
                if (f != null) folders.Add(f);
            }
            if (folders.Count == 0) return;
            bool turnOn = folders.Any(f => !f.SyncColor);
            foreach (FolderNode f in folders)
            {
                f.SyncColor = turnOn;
                if (!turnOn)
                {
                    entry.SyncSeen.Remove(f.FolderId);
                    continue;
                }
                try
                {
                    SetTabColor(entry, MemberNames(entry, new[] { f.FolderId }), f.Color);
                }
                catch (Exception)
                {
                    ShowMessage("シート見出しの色を変更できませんでした。シートが保護されている可能性があります。");
                }
                entry.SyncSeen[f.FolderId] = new HashSet<string>(f.Sheets);
            }
            MarkChanged(entry);
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

        // "More colors": Excel's own "Colors" dialog (hexagon palette and custom tab), as the horizontal tab
        // shows it. It is Excel's editor of a workbook palette color, so an entry of the palette is lent to it:
        // it starts with the active tab's color, the chosen color is read back, and the entry and the
        // workbook's "saved" state are put back as they were.
        // Excel's dialog cannot be shown in every state (a cell is being edited, another window is active, the
        // workbook does not accept a change of its palette). Then the system color dialog takes its place, so
        // that choosing "More colors" never silently does nothing. Null: cancelled (or nothing could be shown).
        // The dialog opens where the user clicked: the double click of a folder's paint can, or "その他の色" of a palette.
        private int? PickMoreColor(Entry entry)
        {
            using (new DialogPlacer(Cursor.Position)) return PickMoreColorCore(entry);
        }

        private int? PickMoreColorCore(Entry entry)
        {
            int start = -1;
            try
            {
                int c;
                if (TryGetTabColor(entry.Window.ActiveSheet, out c)) start = c;
            }
            catch (Exception)
            {
            }

            int? picked = null;
            bool shown = false;
            try
            {
                shown = TryExcelColorDialog(entry, start, out picked);
            }
            catch (Exception)
            {
                shown = false;
            }
            if (shown) return picked;

            try
            {
                using (var dialog = new System.Windows.Forms.ColorDialog { FullOpen = true, AnyColor = true })
                {
                    if (start >= 0) dialog.Color = System.Drawing.Color.FromArgb(start & 0xFF, (start >> 8) & 0xFF, (start >> 16) & 0xFF);
                    if (dialog.ShowDialog(new ExcelWindowOwner(_app.Hwnd)) != System.Windows.Forms.DialogResult.OK) return null;
                    System.Drawing.Color chosen = dialog.Color;
                    return chosen.R | (chosen.G << 8) | (chosen.B << 16);
                }
            }
            catch (Exception)
            {
                try { ShowMessage("色の設定画面を開けませんでした。"); } catch (Exception) { }
                return null;
            }
        }

        // True when Excel's dialog was shown (picked: the color, or null when it was cancelled); false when it
        // could not be shown.
        private bool TryExcelColorDialog(Entry entry, int start, out int? picked)
        {
            picked = null;
            const int Slot = 56; // the last of the 56 palette colors, the one least likely to be used
            Excel.Workbook wb = entry.Workbook;

            // a cell being edited is entered first (as for the other commands); Excel shows no dialog while editing
            CommitCellEdit(entry);
            if (IsExcelEditing()) return false;

            // the dialog edits the palette of the active workbook: it has to be this one
            try { entry.Window.Activate(); } catch (Exception) { }
            try
            {
                if (!wb.Equals(_app.ActiveWorkbook)) return false;
            }
            catch (Exception)
            {
                return false;
            }

            bool wasSaved = false;
            try { wasSaved = wb.Saved; } catch (Exception) { }
            object original = null;
            try
            {
                original = wb.Colors[Slot];
                if (start >= 0) wb.Colors[Slot] = start;
                try { SetForegroundWindow(new IntPtr(_app.Hwnd)); } catch (Exception) { }
                var watch = System.Diagnostics.Stopwatch.StartNew();
                bool ok = (bool)_app.Dialogs[Excel.XlBuiltInDialog.xlDialogEditColor].Show(Slot);
                if (!ok)
                {
                    // Excel answers "false" at once when it cannot show the dialog; cancelling takes a person longer
                    return watch.ElapsedMilliseconds >= 500;
                }
                picked = Convert.ToInt32(wb.Colors[Slot]);
                return true;
            }
            catch (Exception)
            {
                picked = null;
                return false;
            }
            finally
            {
                try
                {
                    if (original != null) wb.Colors[Slot] = original;
                    if (wasSaved) wb.Saved = true;
                }
                catch (Exception)
                {
                }
            }
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // The menu item "Insert": Excel puts a new worksheet in front of the active sheet. When the active sheet
        // is in a folder, the new sheet goes into that folder, also in front of the folder's first sheet (where
        // the folder would otherwise end behind the new sheet). Without the name rule the model takes it in;
        // with the name rule the sheet is named "Folder | Sheet" at once (a plain name in the middle of the
        // folder's sheets would make the sync move the folder's sheets around it). The sync is stopped meanwhile.
        private void InsertSheet(Entry entry)
        {
            if (entry.Busy) return;
            CommitCellEdit(entry);
            entry.Busy = true;
            try
            {
                FolderNode folder = null;
                string activeKey = null;
                List<string> keysBefore = null;
                if (entry.Folders.HasFolders)
                {
                    RefreshKeys(entry);
                    activeKey = KeyOf(entry, CurrentActive(entry));
                    folder = activeKey == null ? null : entry.Folders.FolderOf(activeKey);
                    keysBefore = new List<string>(entry.Keys);
                }
                var namesBefore = new HashSet<string>(ReadNames(entry), StringComparer.OrdinalIgnoreCase);
                RunExcelCommand(entry, "SheetInsert");
                if (folder != null)
                {
                    List<string> now = ReadNames(entry);
                    var added = new List<int>();
                    for (int i = 0; i < now.Count; i++) if (!namesBefore.Contains(now[i])) added.Add(i);
                    if (added.Count > 0)
                    {
                        if (RuleOn(entry))
                        {
                            foreach (int i in added)
                            {
                                string full = NameRule.Compose(folder.Name, NameRule.BaseOf(now[i]));
                                if (full.Length > NameRule.MaxSheetName) continue;
                                if (now.Any(n => string.Equals(n, full, StringComparison.OrdinalIgnoreCase))) continue;
                                SetSheetName(entry.Workbook.Sheets[i + 1], full);
                                now[i] = full;
                            }
                        }
                        else
                        {
                            RefreshKeys(entry);
                            List<string> newKeys = entry.Keys.Where(k => !keysBefore.Contains(k)).ToList();
                            if (newKeys.Count > 0) entry.Folders.Move(newKeys, new string[0], folder.FolderId, activeKey, null);
                        }
                    }
                }
                entry.KeySignature = null;
            }
            catch (Exception)
            {
                // The command is not available now (e.g. structure protection): nothing was inserted.
            }
            finally
            {
                entry.Busy = false;
            }
            SyncSheets(entry);
        }

        // The "+" button: a new worksheet after the active sheet, as Excel's own button does.
        private void AddSheet(Entry entry)
        {
            // Excel's events for the new sheet start a sync while the sheet still has its own name; with the
            // name rule on, that would move the sheet (or the sheets around it) before it is named
            if (entry.Busy) return;
            CommitCellEdit(entry);
            entry.Busy = true;
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
                if (folder != null && RuleOn(entry))
                {
                    // named after the folder at once, so that the sheet is part of the folder's run of sheets
                    // where Excel put it (and not moved away from the middle of the folder)
                    try
                    {
                        object newSheet = entry.Window.ActiveSheet; // the new sheet is the active one now
                        string own = SheetName(newSheet);
                        string full = NameRule.Compose(folder.Name, own);
                        if (own != null && full.Length <= NameRule.MaxSheetName) SetSheetName(newSheet, full);
                    }
                    catch (Exception)
                    {
                        // it keeps Excel's name and is moved into the folder by the code below
                    }
                }
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
                        if (RuleOn(entry))
                        {
                            string problem = ApplyNamesFromModel(entry);
                            if (problem != null) ShowMessage(problem);
                        }
                    }
                    entry.KeySignature = null;
                }
            }
            catch (Exception)
            {
                ShowMessage("ブックの構造が保護されているため、シートを追加できません。");
            }
            finally
            {
                entry.Busy = false;
                // the new row shows at once, not at the next tick of the timer (as after "Insert")
                SyncSheets(entry);
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
                    if (copy) { ws.Copy(before, after); return; }
                    // the moved sheet becomes the active one without the old one showing in between (no sync until it is selected)
                    bool wasBusy = entry.Busy;
                    entry.Busy = true;
                    try
                    {
                        ws.Move(before, after);
                        SelectSheets(entry, names, names[0]);
                    }
                    finally
                    {
                        entry.Busy = wasBusy;
                    }
                    entry.KeySignature = null;
                    SyncSheets(entry);
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
            if (RuleOn(entry))
            {
                // newName is what was typed: the part after "A | " for a sheet in a folder
                string folderName, own;
                bool inFolder = NameRule.TryParse(oldName, out folderName, out own);
                if (NameRule.HasSeparator(newName))
                {
                    ShowMessage("「" + NameRule.Separator + "」を含む名前は、縦タブでは入力できません。");
                    return;
                }
                if (inFolder)
                {
                    if (folderName.Length + newName.Length > NameRule.MaxCombined)
                    {
                        ShowMessage("フォルダ名とシート名を合わせた長さは、" + NameRule.MaxCombined + " 文字までです。");
                        return;
                    }
                    newName = NameRule.Compose(folderName, newName);
                }
            }
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
            entry.SelectionChanged = false;
            try
            {
                // Already the selection in Excel: do not touch it, a click on the active tab must not
                // move the focus into the grid (that would cancel a double click that renames).
                if (SameSelection(entry, names, active)) return;
                entry.SelectionChanged = true;

                // The same group with another active sheet: activating a sheet of the group keeps the group, so it
                // is not taken apart and put together again on screen.
                HashSet<string> current = CurrentGroup(entry);
                if (current != null && current.Count > 1 && current.SetEquals(names))
                {
                    ActivateSheet(entry.Workbook.Sheets[active]);
                    return;
                }

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

                // Excel would show each step of the selection (the tabs and the sheet drawn again for every sheet): the
                // screen is updated once, at the end.
                bool frozen = false, screen = true;
                try
                {
                    if (ordered.Count > 1)
                    {
                        screen = entry.Workbook.Application.ScreenUpdating;
                        entry.Workbook.Application.ScreenUpdating = false;
                        frozen = true;
                    }
                    bool first = true;
                    foreach (string n in ordered)
                    {
                        SelectSheet(entry.Workbook.Sheets[n], first);
                        first = false;
                    }
                }
                finally
                {
                    if (frozen) try { entry.Workbook.Application.ScreenUpdating = screen; } catch (Exception) { }
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

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetParent(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindowEnabled(IntPtr hWnd);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnableWindow(IntPtr hWnd, bool enable);

        // SendInput puts all the events of a key press into the input queue at once, so that the user's own keys cannot
        // come between them.
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct KeyInput
        {
            public uint Type;
            public KeyInputUnion U;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Explicit)]
        private struct KeyInputUnion
        {
            [System.Runtime.InteropServices.FieldOffset(0)] public MouseInputData Mouse; // (the largest member: it gives the union its size)
            [System.Runtime.InteropServices.FieldOffset(0)] public KeyboardInputData Keyboard;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct MouseInputData
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct KeyboardInputData
        {
            public ushort wVk;
            public ushort wScan;
            public uint dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern uint SendInput(uint count, KeyInput[] inputs, int size);

        // Presses and releases the keys: each pair is (key, true = released).
        private static void SendKeys(params KeyValuePair<byte, bool>[] events)
        {
            var inputs = new KeyInput[events.Length];
            for (int i = 0; i < events.Length; i++)
            {
                inputs[i].Type = 1; // INPUT_KEYBOARD
                inputs[i].U.Keyboard.wVk = events[i].Key;
                inputs[i].U.Keyboard.dwFlags = events[i].Value ? (uint)KeyUp : 0u;
            }
            SendInput((uint)inputs.Length, inputs, System.Runtime.InteropServices.Marshal.SizeOf(typeof(KeyInput)));
        }

        // Waits for a message to arrive (or the time to pass) instead of sleeping, then lets the messages be handled.
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint MsgWaitForMultipleObjects(uint count, IntPtr[] handles, bool waitAll, uint milliseconds, uint wakeMask);

        private const uint QsAllInput = 0x04FF;

        private static void PumpMessages(int milliseconds)
        {
            MsgWaitForMultipleObjects(0, null, false, (uint)milliseconds, QsAllInput);
            System.Windows.Forms.Application.DoEvents();
        }

        private const int KeyUp = 0x0002;
        private const byte VkControl = 0x11;
        private const byte VkReturn = 0x0D;
        private const byte VkF6 = 0x75;

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint threadId, EnumWindowsProc callback, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // A dialog of Excel (a message, "save changes?", a command's dialog) is open: Excel disables its main window (the main window of
        // every workbook window, when there are several). While that lasts, nothing here asks Excel anything.
        // Excel's main windows belong to its UI thread, the thread the add-in runs on: only that thread's windows are looked at, not
        // every window of the desktop.
        private static bool ExcelMainWindowDisabled(Entry entry)
        {
            try
            {
                bool disabled = false;
                var name = new System.Text.StringBuilder(32);
                EnumThreadWindows(GetCurrentThreadId(), (h, l) =>
                {
                    name.Clear();
                    GetClassName(h, name, name.Capacity);
                    if (name.ToString() == "XLMAIN" && !IsWindowEnabled(h)) { disabled = true; return false; }
                    return true;
                }, IntPtr.Zero);
                return disabled;
            }
            catch (Exception)
            {
                return false;
            }
        }

        // A cell (or the formula bar) is being edited: the keyboard focus is in Excel's edit window.
        private static bool IsExcelEditing()
        {
            try
            {
                IntPtr focus = GetFocus();
                if (focus == IntPtr.Zero) return false;
                var name = new System.Text.StringBuilder(64);
                GetClassName(focus, name, name.Capacity);
                string cls = name.ToString();
                return cls == "EXCEL6" || cls == "EXCEL<";
            }
            catch (Exception)
            {
                return false;
            }
        }

        // While a cell is edited, Excel disables its task panes, so that a click on the vertical tabs would
        // not arrive at all. Excel's own tabs do work then (they enter the text and switch the sheet), so
        // the pane is enabled again for as long as the editing lasts; Excel enables its panes itself
        // afterwards. Only the windows between the pane and Excel's main window are touched, and only
        // while editing (with a dialog open Excel disables the main window, and the pane stays disabled).
        private static void AllowPaneInput(Entry entry)
        {
            try
            {
                if (!IsExcelEditing()) return;
                IntPtr h = entry.Control.Handle;
                var name = new System.Text.StringBuilder(64);
                for (int i = 0; i < 10 && h != IntPtr.Zero; i++)
                {
                    name.Clear();
                    GetClassName(h, name, name.Capacity);
                    if (name.ToString() == "XLMAIN") break;
                    if (!IsWindowEnabled(h)) EnableWindow(h, true);
                    h = GetParent(h);
                }
            }
            catch (Exception)
            {
                // The window is closing.
            }
        }

        // The user chooses another sheet in the vertical tabs while a cell is being edited. Excel's own
        // tabs enter what was typed into the cell (which stays the active one) and then switch the sheet.
        // The sheet cannot be switched while editing, so Ctrl+Enter is pressed for the user (it enters the
        // text and keeps the cell), and the switch is made when the editing has ended. The click did not
        // take the focus away from the edit window (TabStrip.WndProc), so the keys reach it.
        private static void CommitCellEdit(Entry entry)
        {
            if (!IsExcelEditing()) return;
            SendKeys(
                new KeyValuePair<byte, bool>(VkControl, false), new KeyValuePair<byte, bool>(VkReturn, false),
                new KeyValuePair<byte, bool>(VkReturn, true), new KeyValuePair<byte, bool>(VkControl, true));
            // the keys are handled by Excel's message loop, which is this thread: let it run
            for (int i = 0; i < 60; i++)
            {
                System.Windows.Forms.Application.DoEvents();
                if (!IsExcelEditing()) return;
                PumpMessages(25);
            }
        }

        private delegate bool EnumChildProc(IntPtr hWnd, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumChildProc callback, IntPtr lParam);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder name, int max);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        // After a click in the vertical tabs has chosen another sheet: the keyboard focus goes to the sheet
        // (Excel's window "EXCEL7"), as it does after a click on Excel's own tabs, so that the next key
        // press is typed into the selected cell. A click that chose nothing new leaves the focus where it
        // is (it could be the first click of a double click that renames).
        // While the pane has the focus, Excel does not pass the keys on to the sheet, whatever window gets
        // the focus afterwards. Such a click (hadFocus) comes after a click on a folder, on the active tab or a right
        // click (which give the pane the focus); a click that chooses a sheet does not (TabStrip.WndProc). Excel is
        // then sent F6, which ends the pane's hold on the keyboard.
        private static void FocusGrid(Entry entry, bool hadFocus)
        {
            if (!entry.SelectionChanged) return;
            entry.SelectionChanged = false;
            try
            {
                if (hadFocus)
                {
                    // F6, the key with which Excel itself moves the focus on to the next part of its window, ends the
                    // pane's hold. Hiding and showing the pane again (as before) also ended it, but Excel then laid the
                    // window out twice and it flickered.
                    SendKeys(new KeyValuePair<byte, bool>(VkF6, false), new KeyValuePair<byte, bool>(VkF6, true));
                    for (int i = 0; i < 10; i++) PumpMessages(10);
                }
                IntPtr grid = IntPtr.Zero;
                var name = new System.Text.StringBuilder(64);
                EnumChildWindows(new IntPtr(entry.Window.Hwnd), (h, l) =>
                {
                    name.Clear();
                    GetClassName(h, name, name.Capacity);
                    if (name.ToString() == "EXCEL7" && IsWindowVisible(h)) { grid = h; return false; }
                    return true;
                }, IntPtr.Zero);
                if (grid != IntPtr.Zero) SetFocus(grid);
            }
            catch (Exception)
            {
                // The window is closing; the focus stays where it is.
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

        // The visible sheets Excel has selected now; null when it cannot be read.
        private static HashSet<string> CurrentGroup(Entry entry)
        {
            try
            {
                var selected = new HashSet<string>();
                foreach (object sheet in (System.Collections.IEnumerable)entry.Window.SelectedSheets)
                {
                    string name;
                    if (IsVisibleSheet(sheet, out name)) selected.Add(name);
                }
                return selected;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static void ActivateSheet(object sheet)
        {
            var ws = sheet as Excel.Worksheet;
            if (ws != null) { ws.Activate(); return; }
            var ch = sheet as Excel.Chart;
            if (ch != null) ch.Activate();
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
