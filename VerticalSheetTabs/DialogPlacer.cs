using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Puts the color dialog of "その他の色" where the user clicked (Agents.md 12.234). Excel's own dialog cannot be told where to open, so the
    // first dialog window that is shown on this thread while the placer lives is moved: its top left corner to the point, kept inside the
    // work area of the screen of the point. Windows tells of the window as it is shown (a WinEvent), before the user can work with it.
    internal sealed class DialogPlacer : IDisposable
    {
        private delegate void WinEventProc(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr module, WinEventProc proc, uint processId, uint threadId, uint flags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hook);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder name, int max);

        [DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out Rect rect);

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int cx, int cy, uint flags);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [StructLayout(LayoutKind.Sequential)]
        private struct Rect
        {
            public int Left, Top, Right, Bottom;
        }

        private const uint EventObjectShow = 0x8002;
        private const uint GaRoot = 2;
        private const uint SwpNoSize = 0x0001, SwpNoZOrder = 0x0004, SwpNoActivate = 0x0010;

        private readonly WinEventProc _proc; // (kept: Windows calls it as long as the hook lives)
        private readonly Point _at;
        private IntPtr _hook;
        private bool _done;

        public DialogPlacer(Point at)
        {
            _at = at;
            _proc = OnShow;
            try
            {
                _hook = SetWinEventHook(EventObjectShow, EventObjectShow, IntPtr.Zero, _proc, (uint)System.Diagnostics.Process.GetCurrentProcess().Id, GetCurrentThreadId(), 0);
            }
            catch (Exception)
            {
                _hook = IntPtr.Zero; // (the dialog then opens where it would)
            }
        }

        private void OnShow(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            try
            {
                if (_done || hwnd == IntPtr.Zero || idObject != 0 || GetAncestor(hwnd, GaRoot) != hwnd) return;
                var name = new StringBuilder(64);
                GetClassName(hwnd, name, name.Capacity);
                string cls = name.ToString();
                // Excel's dialogs, and the system's color dialog
                if (cls != "bosa_sdm_XL9" && cls != "#32770") return;
                Rect r;
                if (!GetWindowRect(hwnd, out r)) return;
                int w = r.Right - r.Left, h = r.Bottom - r.Top;
                Rectangle area = Screen.FromPoint(_at).WorkingArea;
                int x = Math.Max(area.Left, Math.Min(_at.X, area.Right - w));
                int y = Math.Max(area.Top, Math.Min(_at.Y, area.Bottom - h));
                SetWindowPos(hwnd, IntPtr.Zero, x, y, 0, 0, SwpNoSize | SwpNoZOrder | SwpNoActivate);
                _done = true;
            }
            catch (Exception)
            {
                // the dialog stays where it opened
            }
        }

        public void Dispose()
        {
            if (_hook == IntPtr.Zero) return;
            try { UnhookWinEvent(_hook); } catch (Exception) { }
            _hook = IntPtr.Zero;
        }
    }
}
