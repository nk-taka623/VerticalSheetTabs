using System;
using System.Runtime.InteropServices;

namespace VerticalSheetTabs
{
    // Excel runs at the real display scaling, but WinForms controls inside a task pane see 96 dpi
    // (Control.DeviceDpi always says 96). The drawing in the pane is scaled by hand, using the DPI
    // of Excel's own main window.
    public static class DpiHelper
    {
        // Excel's main window, set when the add-in starts.
        public static IntPtr ExcelWindow;

        [DllImport("user32.dll")]
        private static extern uint GetDpiForWindow(IntPtr hwnd);

        public static int GetDpi(IntPtr fallbackWindow)
        {
            try
            {
                uint dpi = 0;
                if (ExcelWindow != IntPtr.Zero) dpi = GetDpiForWindow(ExcelWindow);
                if (dpi == 0 && fallbackWindow != IntPtr.Zero) dpi = GetDpiForWindow(fallbackWindow);
                return dpi > 0 ? (int)dpi : 96;
            }
            catch (EntryPointNotFoundException)
            {
                return 96; // before Windows 10 version 1607
            }
        }
    }
}
