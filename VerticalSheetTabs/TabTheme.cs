using System;
using System.Drawing;
using Microsoft.Win32;

namespace VerticalSheetTabs
{
    // The colors of Excel's tab bar in the current Office theme (measured on a real Excel).
    // The selected tab (white, dark text), the colored tabs and the green
    // underline look the same in every theme; the bar, the separators, the inactive text and the
    // "+" differ.
    public class TabTheme
    {
        public Color BarBack;
        public Color TabBorder;      // outline of a selected tab
        public Color Separator;
        public Color InactiveText;
        public Color PlusColor;
        public Color PlusHoverBack;
        public Color DropMarker;

        // Colorful / white / "use the system setting" with Windows in light mode (measured: colorful).
        public static readonly TabTheme Light = new TabTheme
        {
            BarBack = Color.FromArgb(237, 235, 237),
            TabBorder = Color.FromArgb(224, 224, 224),
            Separator = Color.FromArgb(209, 209, 209),
            InactiveText = Color.FromArgb(36, 36, 36),
            PlusColor = Color.FromArgb(57, 57, 56),
            PlusHoverBack = Color.FromArgb(250, 250, 250),
            DropMarker = Color.Black
        };

        // Dark gray (measured). The drop marker color is assumed (not seen).
        public static readonly TabTheme Dark = new TabTheme
        {
            BarBack = Color.FromArgb(86, 86, 86),
            TabBorder = Color.FromArgb(71, 71, 71),
            Separator = Color.FromArgb(138, 138, 138),
            InactiveText = Color.White,
            PlusColor = Color.FromArgb(212, 212, 212),
            PlusHoverBack = Color.FromArgb(61, 61, 61),
            DropMarker = Color.White
        };

        // Black (measured, except the hover background of the "+" and the drop marker, which are assumed).
        public static readonly TabTheme Black = new TabTheme
        {
            BarBack = Color.FromArgb(10, 10, 10),
            TabBorder = Color.FromArgb(66, 66, 66),
            Separator = Color.FromArgb(102, 102, 102),
            InactiveText = Color.White,
            PlusColor = Color.FromArgb(212, 212, 212),
            PlusHoverBack = Color.FromArgb(36, 36, 36),
            DropMarker = Color.White
        };

        private static TabTheme _cached = Light;
        private static DateTime _checkedAt = DateTime.MinValue;

        // Reads the Office theme from the registry; the answer is cached for a couple of seconds.
        public static TabTheme Current
        {
            get
            {
                if (DateTime.Now - _checkedAt > TimeSpan.FromSeconds(2))
                {
                    _checkedAt = DateTime.Now;
                    _cached = Detect();
                }
                return _cached;
            }
        }

        // "UI Theme" (HKCU\Software\Microsoft\Office\16.0\Common): 0 colorful, 3 dark gray, 4 black,
        // 5 white, 6 use the system setting. When the value is missing Office follows the system.
        private static TabTheme Detect()
        {
            try
            {
                int? value = null;
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Common"))
                {
                    object v = key == null ? null : key.GetValue("UI Theme");
                    if (v is int) value = (int)v;
                }
                if (value == 3) return Dark;
                if (value == 4) return Black;
                if (value == 0 || value == 5) return Light;
                return SystemUsesDarkApps() ? Dark : Light; // "use the system setting"
            }
            catch (Exception)
            {
                return Light;
            }
        }

        private static bool SystemUsesDarkApps()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
            {
                object v = key == null ? null : key.GetValue("AppsUseLightTheme");
                return v is int && (int)v == 0;
            }
        }
    }
}
