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

            var control = new TabPaneControl();
            var pane = _collection.Add(control, "Sheet Tabs", window);
            // First open: left side (see Agents.md). Later this comes from the workbook property.
            pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            pane.Width = 150;
            pane.Visible = true;

            var entry = new Entry { Pane = pane, Control = control, Window = window, Workbook = workbook };
            _entries[hwnd] = entry;

            pane.DockPositionChanged += (s, e) => RefreshInfo(entry);
            control.Resize += (s, e) => RefreshInfo(entry);
            control.LeftClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            control.RightClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            control.SheetSelected += name => ActivateSheet(entry, name);
            RefreshInfo(entry);
            SyncSheets(entry);
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
                entry.Control.SetInfo(dock);
            }
            catch (Exception)
            {
                // The window or pane may already be closing.
            }
        }
    }
}
