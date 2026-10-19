using System;
using System.Collections.Generic;
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
        }

        private readonly CustomTaskPaneCollection _collection;
        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();
        // Excel raises no events for rename, reorder or hide/unhide, so the lists
        // are re-read on a short timer (cheap: only names are compared).
        private readonly Timer _timer = new Timer { Interval = 300 };

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
            var pane = _collection.Add(control, "Sheet Tabs", window);
            pane.DockPosition = ToDock(settings.Side);
            pane.Width = 150; // temporary; the real width is applied in InitializeWidth

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook, Settings = settings };
            _entries[hwnd] = entry;

            pane.VisibleChanged += (s, e) => OnPaneVisibleChanged(entry);
            control.SheetSelected += name => ActivateSheet(entry, name);
            control.SettingsRequested += () => ShowSettings(entry);
            control.AddRequested += () => AddSheet(entry);
            control.InsertRequested += () => RunExcelCommand(entry, "SheetInsert");
            control.DeleteRequested += () => RunExcelCommand(entry, "SheetDelete");
            control.HideRequested += () => RunExcelCommand(entry, "SheetHide");
            control.UnhideRequested += () => RunExcelCommand(entry, "SheetUnhide");
            control.RenameRequested += (oldName, newName) => RenameSheet(entry, oldName, newName);
            control.MoveCopyRequested += () => RunExcelCommand(entry, "SheetMoveOrCopy");
            control.MoveRequested += (name, before, copy) => MoveOrCopySheet(entry, name, before, copy);

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
                entry.Window.DisplayWorkbookTabs = horizontal;
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

        // Applies the stored width, or 16 half-width characters on first open.
        private void InitializeWidth(Entry entry)
        {
            try
            {
                if (!_entries.ContainsValue(entry)) return;
                int dpi = entry.Control.DeviceDpi;
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
                        entry.Settings.WidthDip = (int)Math.Round(entry.Control.Width * 96.0 / entry.Control.DeviceDpi);
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

        public void RemovePane(Excel.Window window)
        {
            if (window == null) return;
            int hwnd = window.Hwnd;
            Entry entry;
            if (!_entries.TryGetValue(hwnd, out entry)) return;
            _entries.Remove(hwnd);
            _collection.Remove(entry.Pane);
        }

        public void SyncAll()
        {
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                SyncSheets(entry);
            }
        }

        private static void SyncSheets(Entry entry)
        {
            try
            {
                var names = new List<string>();
                bool anyHidden = false;
                foreach (object sheet in entry.Workbook.Sheets)
                {
                    string name;
                    if (IsVisibleSheet(sheet, out name)) names.Add(name);
                    else if (SheetVisibility(sheet) == Excel.XlSheetVisibility.xlSheetHidden) anyHidden = true;
                }
                entry.Control.HasHiddenSheets = anyHidden;

                string active = null;
                object activeSheet = entry.Window.ActiveSheet;
                string activeName;
                if (activeSheet != null && IsVisibleSheet(activeSheet, out activeName)) active = activeName;

                entry.Control.SetSheets(names, active);
            }
            catch (Exception)
            {
                // Excel is busy (e.g. a cell is being edited) or the window is closing.
                // The next timer tick will retry.
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

        // Drag and drop: move (or Ctrl+drag copy) a sheet in front of another visible sheet, or to
        // the end of the visible ones. Hidden sheets stay where they are relative to their neighbours.
        private void MoveOrCopySheet(Entry entry, string name, string beforeName, bool copy)
        {
            try
            {
                entry.Window.Activate();
                object sheet = entry.Workbook.Sheets[name];
                object before = Type.Missing;
                object after = Type.Missing;
                if (beforeName != null) before = entry.Workbook.Sheets[beforeName];
                else after = LastVisibleSheet(entry.Workbook);

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

        private static void ActivateSheet(Entry entry, string name)
        {
            try
            {
                object sheet = entry.Workbook.Sheets[name];
                var ws = sheet as Excel.Worksheet;
                if (ws != null) { ws.Activate(); return; }
                var ch = sheet as Excel.Chart;
                if (ch != null) ch.Activate();
            }
            catch (Exception)
            {
                // The sheet may have just been renamed or deleted; the next sync fixes the list.
            }
        }
    }
}
