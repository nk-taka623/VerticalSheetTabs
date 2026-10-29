using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace VerticalSheetTabs
{
    // The settings of this PC (the display mode, the side, the looks, the animations, the badges), shared by all workbooks:
    // %LocalAppData%\VerticalSheetTabs\settings.ini, one "key=value" per line. The file is written only when the settings dialog
    // is closed with OK and something was changed. A file that is missing, unreadable or unwritable means the defaults;
    // no error is shown.
    public static class PcSettings
    {
        private static string FilePath
        {
            get
            {
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "VerticalSheetTabs", "settings.ini");
            }
        }

        private static Dictionary<string, string> Read()
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            try
            {
                foreach (string line in File.ReadAllLines(FilePath, new UTF8Encoding(false)))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    values[line.Substring(0, eq).Trim()] = line.Substring(eq + 1).Trim();
                }
            }
            catch (Exception)
            {
                // Missing or unreadable: the defaults.
            }
            return values;
        }

        // Puts the PC's values into the settings (what the file lacks or holds wrongly stays as it is).
        public static void ApplyTo(WorkbookSettings s)
        {
            Dictionary<string, string> v = Read();
            string t;
            if (v.TryGetValue("Mode", out t)) s.Mode = t == "A" ? TabMode.HorizontalOnly : TabMode.Both;
            if (v.TryGetValue("Side", out t)) s.Side = t == "Right" ? TabSide.Right : TabSide.Left;
            if (v.TryGetValue("ActiveMarker", out t))
            {
                s.ActiveMarker = t == "None" ? ActiveMarker.None : t == "Triangle" ? ActiveMarker.Triangle : t == "EdgeBars" ? ActiveMarker.EdgeBars : t == "CheckBadge" ? ActiveMarker.CheckBadge : ActiveMarker.Underline;
            }
            if (v.TryGetValue("SelectStyle", out t)) s.SelectStyle = WorkbookSettings.ParseSelectStyle(t);
            if (v.TryGetValue("Outline", out t)) s.OutlineVisible = t != "Off";
            int n;
            if (v.TryGetValue("OutlineLevel", out t) && int.TryParse(t, out n) && n >= 1 && n <= 9) s.OutlineLevel = n;
            WobbleType w;
            if (v.TryGetValue("Wobble", out t) && Enum.TryParse(t, out w) && Enum.IsDefined(typeof(WobbleType), w)) s.Wobble = w;
            if (v.TryGetValue("WobbleAmp", out t) && int.TryParse(t, out n)) s.WobbleAmp = Math.Max(WobbleMotion.AmpMin, Math.Min(WobbleMotion.AmpMax, n));
            if (v.TryGetValue("WobbleDur", out t) && int.TryParse(t, out n)) s.WobbleDur = Math.Max(WobbleMotion.DurMin, Math.Min(WobbleMotion.DurMax, n));
            if (v.TryGetValue("WobbleSound", out t)) s.WobbleSound = t == "On";
            if (v.TryGetValue("WobbleVolume", out t) && int.TryParse(t, out n)) s.WobbleVolume = Math.Max(0, Math.Min(100, n));
            if (v.TryGetValue("WobbleSampleStyle", out t)) s.WobbleSampleStyle = t == "On";
            if (v.TryGetValue("RowAnim", out t)) s.RowAnim = t != "Off";
            if (v.TryGetValue("DoneAnim", out t)) s.DoneAnim = t != "Off";
            if (v.TryGetValue("Badges", out t)) s.Badges = WorkbookSettings.ParseBadges(t);
        }

        // The values as they are written in the file.
        public static Dictionary<string, string> ToValues(WorkbookSettings s)
        {
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                { "Mode", s.Mode == TabMode.HorizontalOnly ? "A" : "C" },
                { "Side", s.Side == TabSide.Right ? "Right" : "Left" },
                { "ActiveMarker", s.ActiveMarker.ToString() },
                { "SelectStyle", WorkbookSettings.SelectStyleToString(s.SelectStyle) },
                { "Outline", s.OutlineVisible ? "On" : "Off" },
                { "OutlineLevel", s.OutlineLevel.ToString() },
                { "Wobble", s.Wobble.ToString() },
                { "WobbleAmp", s.WobbleAmp.ToString() },
                { "WobbleDur", s.WobbleDur.ToString() },
                { "WobbleSound", s.WobbleSound ? "On" : "Off" },
                { "WobbleVolume", s.WobbleVolume.ToString() },
                { "WobbleSampleStyle", s.WobbleSampleStyle ? "On" : "Off" },
                { "RowAnim", s.RowAnim ? "On" : "Off" },
                { "DoneAnim", s.DoneAnim ? "On" : "Off" },
                { "Badges", WorkbookSettings.BadgesToString(s.Badges) }
            };
        }

        // Writes the values the dialog changed (after differs from before) into the file; the other values in the file stay,
        // so that a window that holds older values does not overwrite what another one has saved.
        public static void Update(Dictionary<string, string> before, Dictionary<string, string> after)
        {
            try
            {
                var changed = new List<KeyValuePair<string, string>>();
                foreach (KeyValuePair<string, string> pair in after)
                {
                    string old;
                    if (!before.TryGetValue(pair.Key, out old) || old != pair.Value) changed.Add(pair);
                }
                if (changed.Count == 0) return;

                Dictionary<string, string> values = Read();
                foreach (KeyValuePair<string, string> pair in changed) values[pair.Key] = pair.Value;

                string path = FilePath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                var text = new StringBuilder();
                foreach (KeyValuePair<string, string> pair in values) text.Append(pair.Key).Append('=').Append(pair.Value).Append("\r\n");
                // a name of its own: two Excel processes can write at the same time
                string temp = path + "." + Guid.NewGuid().ToString("N") + ".tmp";
                try
                {
                    File.WriteAllText(temp, text.ToString(), new UTF8Encoding(false));
                    if (File.Exists(path)) File.Replace(temp, path, null);
                    else File.Move(temp, path);
                }
                finally
                {
                    try { if (File.Exists(temp)) File.Delete(temp); }
                    catch (Exception) { }
                }
            }
            catch (Exception)
            {
                // Not writable: the settings stay in this window only.
            }
        }
    }
}
