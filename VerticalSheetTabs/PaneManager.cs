using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VerticalSheetTabs
{
    // Keeps one custom task pane per Excel window (keyed by window handle).
    public class PaneManager
    {
        private class Entry
        {
            public CustomTaskPane Pane;
            public TabPaneControl Control;
            public Excel.Window Window;
        }

        private readonly CustomTaskPaneCollection _collection;
        private readonly Dictionary<int, Entry> _entries = new Dictionary<int, Entry>();

        public PaneManager(CustomTaskPaneCollection collection)
        {
            _collection = collection;
        }

        public void EnsurePane(Excel.Window window)
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

            var entry = new Entry { Pane = pane, Control = control, Window = window };
            _entries[hwnd] = entry;

            pane.DockPositionChanged += (s, e) => Refresh(entry);
            control.Resize += (s, e) => Refresh(entry);
            control.LeftClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
            control.RightClicked += (s, e) => entry.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            Refresh(entry);
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

        private static void Refresh(Entry entry)
        {
            try
            {
                string dock = entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft
                    ? "Left"
                    : entry.Pane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight
                        ? "Right" : entry.Pane.DockPosition.ToString();
                entry.Control.SetInfo(entry.Window.Caption, dock);
            }
            catch (Exception)
            {
                // The window or pane may already be closing.
            }
        }
    }
}
