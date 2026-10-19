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
        }

        private readonly CustomTaskPaneCollection _collection;
        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();
        // Excel raises no events for rename, reorder or hide/unhide, so the lists
        // are re-read on a short timer (cheap: only names are compared).
        private readonly Timer _timer = new Timer { Interval = 300 };

        public PaneManager(CustomTaskPaneCollection collection)
        {
            _collection = collection;
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
            pane.DockPosition = settings.Side == TabSide.Right
                ? Office.MsoCTPDockPosition.msoCTPDockPositionRight
                : Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Width = 150; // temporary; the real width is applied in InitializeWidth
            pane.Visible = true;

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook, Settings = settings };
            _entries[hwnd] = entry;

            // The pane is laid out asynchronously, so the content width is only
            // reliable a moment after it is shown.
            var init = new Timer { Interval = 250 };
            init.Tick += (s, e) =>
            {
                init.Stop();
                init.Dispose();
                InitializeWidth(entry);
            };
            init.Start();

            pane.DockPositionChanged += (s, e) => RefreshInfo(entry);
            control.Resize += (s, e) => RefreshInfo(entry);
            control.LeftClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            control.RightClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            control.SheetSelected += name => ActivateSheet(entry, name);
            RefreshInfo(entry);
            SyncSheets(entry);
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
                RefreshInfo(entry);
            }
            catch (Exception)
            {
                // The pane may have been closed meanwhile.
            }
        }

        // Called before the workbook is saved: stores the current state in its properties.
        public void SaveSettings(Excel.Workbook workbook)
        {
            foreach (Entry entry in new List<Entry>(_entries.Values))
            {
                if (!entry.Initialized) continue;
                if (!entry.Workbook.Equals(workbook)) continue;
                try
                {
                    entry.Settings.Side = entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight
                        ? TabSide.Right : TabSide.Left;
                    entry.Settings.WidthDip = (int)Math.Round(entry.Control.Width * 96.0 / entry.Control.DeviceDpi);
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

        private static void RefreshInfo(Entry entry)
        {
            try
            {
                string dock = entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft
                    ? "Left"
                    : entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight
                        ? "Right" : entry.Pane.DockPosition.ToString();
                entry.Control.SetInfo(dock, entry.Pane.Width, entry.Chrome);
            }
            catch (Exception)
            {
                // The window or pane may already be closing.
            }
        }
    }
}
