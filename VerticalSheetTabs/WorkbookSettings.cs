using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VerticalSheetTabs
{
    public enum TabMode
    {
        HorizontalOnly, // A
        VerticalOnly,   // B
        Both            // C
    }

    // How a selected folder or sheet looks (Agents.md 12.99).
    public enum SelectStyle
    {
        Standard,   // Excel's own look: white, with a thin border and (the active sheet) an underline
        DoubleFrame, // a shadow, a double frame and, on the active sheet, a triangle at each end
        FrostedGlass, // a pane of frosted glass whose edges glow from inside (Agents.md 12.105)
        GlossyGlass,  // a pane of glass with a reflection on its upper half
        FrostedGlass2, // frosted glass tinted with the row's color (Agents.md 12.107)
        Inset,         // pressed in: a darker body with a shadow inside its upper and left edges (Agents.md 12.115)
        Stripes,       // diagonal stripes with an outline (Agents.md 12.131)
        Crystal        // diagonal cut faces: a light one and a dark one, the light color fading out into the dark one (Agents.md 12.139)
    }

    // The mark of the active sheet (Agents.md 12.110), drawn with every look of SelectStyle.
    public enum ActiveMarker
    {
        None,
        Underline, // a line under the name
        Triangle,  // a small triangle at each end
        EdgeBars,  // a thick bar at each end (Agents.md 12.116)
        CheckBadge // a round check mark at the left of the name (Agents.md 12.131)
    }

    // The small marks drawn at the right end of a row (Agents.md 12.190); each one is shown only when it is chosen.
    [Flags]
    public enum BadgeKinds
    {
        None = 0,
        Error = 1, // the sheet has formulas that show an error (#REF!, #DIV/0!, ...)
        Memo = 2   // the sheet has notes or comments
    }

    public enum TabSide
    {
        Left,
        Right
    }

    // Per-workbook settings, stored as custom document properties (see Agents.md 2.2):
    //   VTab.Mode  (string)  "A" / "B" / "C"
    //   VTab.Side  (string)  "Left" / "Right"
    //   VTab.Width (number)  pane content width in device-independent pixels (96 dpi)
    public class WorkbookSettings
    {
        public const string ModeName = "VTab.Mode";
        public const string SideName = "VTab.Side";
        public const string WidthName = "VTab.Width";
        public const string OutlineName = "VTab.Outline";           // string "On" / "Off"
        public const string OutlineLevelName = "VTab.OutlineLevel"; // number 1..9
        public const string NameRuleName = "VTab.NameRule";         // string "On" / "Off"
        public const string WobbleName = "VTab.Wobble";       // string "None" / "Shake" / "Wobble" / "Jelly" / "Bounce" / "Pulse" / "Swing" / "Random" (default "None"; Agents.md 12.141)
        public const string WobbleAmpName = "VTab.WobbleAmp"; // number, percent (40 to 200, default 100)
        public const string WobbleDurName = "VTab.WobbleDur"; // number, milliseconds (300 to 1200, default 600)
        public const string WobbleSoundName = "VTab.WobbleSound";   // string "On" / "Off" (default "Off"; Agents.md 12.172)
        public const string WobbleVolumeName = "VTab.WobbleVolume"; // number 0 to 100 (default 60)
        public const string WobbleSampleStyleName = "VTab.WobbleSampleStyle"; // string "On" / "Off" (default "Off"; Agents.md 12.178)
        public const string ActiveMarkerName = "VTab.ActiveMarker"; // string "None" / "Underline" / "Triangle" / "EdgeBars" / "CheckBadge" (default "Underline")
        public const string BadgesName = "VTab.Badges";           // string "None", or "Error", "Memo", "Error,Memo" (default "None"; Agents.md 12.190)
        public const string SelectStyleName = "VTab.SelectStyle";   // string "Standard" / "DoubleFrame" / "FrostedGlass" / "GlossyGlass" / "FrostedGlass2" / "Inset" / "Stripes" / "Crystal"

        // Office.MsoDocProperties values.
        private const int PropTypeNumber = 1;
        private const int PropTypeString = 4;

        // No add-in property in the workbook: it is the first time it is opened with the add-in.
        public bool IsFirstOpen { get; private set; }

        // A workbook opened for the first time shows Excel's own tabs only (Agents.md 12.119); the vertical tabs are
        // turned on in the settings dialog.
        public TabMode Mode { get; set; }
        public TabSide Side { get; set; }
        // 0 means "not stored": the caller uses the default (12 half-width characters).
        public int WidthDip { get; set; }
        // The dashed outline around an open folder: shown or not, and its color level (1..9, 5 = middle).
        public bool OutlineVisible { get; set; }
        public int OutlineLevel { get; set; }
        // The sheet-name rule: a sheet "A | B" is the sheet B in the folder A (default: not used).
        public bool NameRule { get; set; }
        // The look of selected folders and sheets (default: the standard type).
        public SelectStyle SelectStyle { get; set; }
        // The mark of the active sheet (default: the underline).
        public ActiveMarker ActiveMarker { get; set; }
        // The badges at the right end of the rows (default: none).
        public BadgeKinds Badges { get; set; }

        public WobbleType Wobble { get; set; }

        public int WobbleAmp { get; set; }

        public int WobbleDur { get; set; }

        public bool WobbleSound { get; set; }

        public int WobbleVolume { get; set; }

        // the sample of the wobble in the settings dialog follows the chosen look and marker
        public bool WobbleSampleStyle { get; set; }

        public WorkbookSettings()
        {
            IsFirstOpen = true;
            Mode = TabMode.HorizontalOnly;
            Side = TabSide.Left;
            WidthDip = 0;
            OutlineVisible = true;
            OutlineLevel = 5;
            NameRule = false;
            SelectStyle = SelectStyle.Standard;
            ActiveMarker = ActiveMarker.Underline;
            Badges = BadgeKinds.None;
            Wobble = WobbleType.None;
            WobbleAmp = WobbleMotion.AmpDefault;
            WobbleDur = WobbleMotion.DurDefault;
            WobbleSound = false;
            WobbleVolume = WobbleSounds.VolumeDefault;
            WobbleSampleStyle = false;
        }

        public static WorkbookSettings Load(Excel.Workbook wb)
        {
            var s = new WorkbookSettings();
            try
            {
                dynamic props = wb.CustomDocumentProperties;
                dynamic mode = Find(props, ModeName);
                dynamic side = Find(props, SideName);
                dynamic width = Find(props, WidthName);

                s.IsFirstOpen = mode == null && side == null && width == null;

                if (mode != null)
                {
                    switch ((string)mode.Value)
                    {
                        case "A": s.Mode = TabMode.HorizontalOnly; break;
                        case "B": s.Mode = TabMode.VerticalOnly; break;
                        case "C": s.Mode = TabMode.Both; break;
                    }
                }
                if (side != null)
                {
                    s.Side = (string)side.Value == "Right" ? TabSide.Right : TabSide.Left;
                }
                dynamic outline = Find(props, OutlineName);
                dynamic level = Find(props, OutlineLevelName);
                dynamic nameRule = Find(props, NameRuleName);
                if (nameRule != null) s.NameRule = (string)nameRule.Value == "On";
                dynamic selectStyle = Find(props, SelectStyleName);
                if (selectStyle != null) s.SelectStyle = ParseSelectStyle((string)selectStyle.Value);
                dynamic marker = Find(props, ActiveMarkerName);
                if (marker != null)
                {
                    string v = (string)marker.Value;
                    s.ActiveMarker = v == "None" ? ActiveMarker.None : v == "Triangle" ? ActiveMarker.Triangle : v == "EdgeBars" ? ActiveMarker.EdgeBars : v == "CheckBadge" ? ActiveMarker.CheckBadge : ActiveMarker.Underline;
                }
                dynamic badges = Find(props, BadgesName);
                if (badges != null) s.Badges = ParseBadges((string)badges.Value);
                dynamic wobble = Find(props, WobbleName);
                if (wobble != null)
                {
                    WobbleType t;
                    if (Enum.TryParse((string)wobble.Value, out t) && Enum.IsDefined(typeof(WobbleType), t)) s.Wobble = t;
                }
                dynamic wobbleAmp = Find(props, WobbleAmpName);
                if (wobbleAmp != null) s.WobbleAmp = Math.Max(WobbleMotion.AmpMin, Math.Min(WobbleMotion.AmpMax, Convert.ToInt32((object)wobbleAmp.Value)));
                dynamic wobbleDur = Find(props, WobbleDurName);
                if (wobbleDur != null) s.WobbleDur = Math.Max(WobbleMotion.DurMin, Math.Min(WobbleMotion.DurMax, Convert.ToInt32((object)wobbleDur.Value)));
                dynamic wobbleSound = Find(props, WobbleSoundName);
                if (wobbleSound != null) s.WobbleSound = (string)wobbleSound.Value == "On";
                dynamic wobbleVolume = Find(props, WobbleVolumeName);
                if (wobbleVolume != null) s.WobbleVolume = Math.Max(0, Math.Min(100, Convert.ToInt32((object)wobbleVolume.Value)));
                dynamic sampleStyle = Find(props, WobbleSampleStyleName);
                if (sampleStyle != null) s.WobbleSampleStyle = (string)sampleStyle.Value == "On";
                if (outline != null) s.OutlineVisible = (string)outline.Value != "Off";
                if (level != null)
                {
                    int v = Convert.ToInt32((object)level.Value);
                    if (v >= 1 && v <= 9) s.OutlineLevel = v;
                }
                if (width != null)
                {
                    int w = Convert.ToInt32((object)width.Value);
                    if (w > 0) s.WidthDip = w;
                }
            }
            catch (Exception)
            {
                // Unreadable or malformed properties: fall back to the defaults.
            }
            return s;
        }

        public void Save(Excel.Workbook wb)
        {
            try
            {
                dynamic props = wb.CustomDocumentProperties;
                Write(props, ModeName, PropTypeString, ModeToString(Mode));
                Write(props, SideName, PropTypeString, Side == TabSide.Right ? "Right" : "Left");
                if (WidthDip > 0) Write(props, WidthName, PropTypeNumber, WidthDip);
                Write(props, OutlineName, PropTypeString, OutlineVisible ? "On" : "Off");
                Write(props, OutlineLevelName, PropTypeNumber, OutlineLevel);
                Write(props, NameRuleName, PropTypeString, NameRule ? "On" : "Off");
                Write(props, SelectStyleName, PropTypeString, SelectStyleToString(SelectStyle));
                Write(props, ActiveMarkerName, PropTypeString, ActiveMarker == ActiveMarker.None ? "None" : ActiveMarker == ActiveMarker.Underline ? "Underline" : ActiveMarker == ActiveMarker.EdgeBars ? "EdgeBars" : ActiveMarker == ActiveMarker.CheckBadge ? "CheckBadge" : "Triangle");
                Write(props, BadgesName, PropTypeString, BadgesToString(Badges));
                Write(props, WobbleName, PropTypeString, Wobble.ToString());
                Write(props, WobbleAmpName, PropTypeNumber, WobbleAmp);
                Write(props, WobbleDurName, PropTypeNumber, WobbleDur);
                Write(props, WobbleSoundName, PropTypeString, WobbleSound ? "On" : "Off");
                Write(props, WobbleVolumeName, PropTypeNumber, WobbleVolume);
                Write(props, WobbleSampleStyleName, PropTypeString, WobbleSampleStyle ? "On" : "Off");
            }
            catch (Exception)
            {
                // Read-only or protected workbook: skip silently, nothing else depends on it.
            }
        }

        private static BadgeKinds ParseBadges(string value)
        {
            var kinds = BadgeKinds.None;
            if (string.IsNullOrEmpty(value)) return kinds;
            foreach (string part in value.Split(','))
            {
                string p = part.Trim();
                if (p == "Error") kinds |= BadgeKinds.Error;
                else if (p == "Memo") kinds |= BadgeKinds.Memo;
            }
            return kinds;
        }

        private static string BadgesToString(BadgeKinds kinds)
        {
            if (kinds == BadgeKinds.None) return "None";
            return (kinds & BadgeKinds.Error) != 0 && (kinds & BadgeKinds.Memo) != 0 ? "Error,Memo" : (kinds & BadgeKinds.Error) != 0 ? "Error" : "Memo";
        }

        private static SelectStyle ParseSelectStyle(string value)
        {
            switch (value)
            {
                case "DoubleFrame": return SelectStyle.DoubleFrame;
                case "FrostedGlass": return SelectStyle.FrostedGlass;
                case "GlossyGlass": return SelectStyle.GlossyGlass;
                case "FrostedGlass2": return SelectStyle.FrostedGlass2;
                case "Inset": return SelectStyle.Inset;
                case "Stripes": return SelectStyle.Stripes;
                case "Crystal":
                case "GradientFrame": return SelectStyle.Crystal; // (GradientFrame, in v1.1.50 to v1.1.52, became the crystal look)
                default: return SelectStyle.Standard;
            }
        }

        private static string SelectStyleToString(SelectStyle style)
        {
            switch (style)
            {
                case SelectStyle.DoubleFrame: return "DoubleFrame";
                case SelectStyle.FrostedGlass: return "FrostedGlass";
                case SelectStyle.GlossyGlass: return "GlossyGlass";
                case SelectStyle.FrostedGlass2: return "FrostedGlass2";
                case SelectStyle.Inset: return "Inset";
                case SelectStyle.Stripes: return "Stripes";
                case SelectStyle.Crystal: return "Crystal";
                default: return "Standard";
            }
        }

        private static string ModeToString(TabMode m)
        {
            return m == TabMode.HorizontalOnly ? "A" : m == TabMode.VerticalOnly ? "B" : "C";
        }

        private static dynamic Find(dynamic props, string name)
        {
            foreach (dynamic p in props)
            {
                if (string.Equals((string)p.Name, name, StringComparison.OrdinalIgnoreCase)) return p;
            }
            return null;
        }

        private static void Write(dynamic props, string name, int type, object value)
        {
            dynamic existing = Find(props, name);
            if (existing != null)
            {
                existing.Value = value;
            }
            else
            {
                props.Add(name, false, type, value);
            }
        }
    }
}
