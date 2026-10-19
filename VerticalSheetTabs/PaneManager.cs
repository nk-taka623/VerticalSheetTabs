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
                foreach (object sheet in entry.Workbook.Sheets)
                {
                    string name;
                    if (IsVisibleSheet(sheet, out name)) names.Add(name);
                }

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
