using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace VerticalSheetTabs
{
    public enum TabMode
    {
        HorizontalOnly, // A
        Both            // C
    }

    // How a selected folder or sheet looks.
    public enum SelectStyle
    {
        Standard,   // Excel's own look: white, with a thin border and (the active sheet) an underline
        DoubleFrame, // a shadow, a double frame and, on the active sheet, a triangle at each end
        FrostedGlass, // a pane of frosted glass whose edges glow from inside
        GlossyGlass,  // a pane of glass with a reflection on its upper half
        FrostedGlass2, // frosted glass tinted with the row's color
        Inset,         // pressed in: a darker body with a shadow inside its upper and left edges
        Stripes,       // diagonal stripes with an outline
        Crystal        // diagonal cut faces: a light one and a dark one, the light color fading out into the dark one
    }

    // The mark of the active sheet, drawn with every look of SelectStyle.
    public enum ActiveMarker
    {
        None,
        Underline, // a line under the name
        Triangle,  // a small triangle at each end
        EdgeBars,  // a thick bar at each end
        CheckBadge // a round check mark at the left of the name
    }

    // The small marks drawn at the right end of a row; each one is shown only when it is chosen.
    [Flags]
    public enum BadgeKinds
    {
        None = 0,
        Error = 1, // the sheet has formulas that show an error (#REF!, #DIV/0!, ...)
        Memo = 2,  // the sheet has notes or comments
        Protect = 4, // the sheet is protected
        Progress = 8 // the progress of the sheet (not started, in progress, done), set by the user
    }

    public enum TabSide
    {
        Left,
        Right
    }

    // The settings of one window. Two of them belong to the workbook, stored as custom document properties:
    //   VTab.Width    (number)  pane content width in device-independent pixels (96 dpi)
    //   VTab.NameRule (string)  "On" / "Off"
    // The others (the display mode, the side, the looks, the animations, the badges) belong to the PC: see PcSettings. The
    // properties with those names that older versions wrote into the workbook are ignored, and deleted when it is saved.
    public class WorkbookSettings
    {
        public const string ModeName = "VTab.Mode";
        public const string SideName = "VTab.Side";
        public const string WidthName = "VTab.Width";
        public const string OutlineName = "VTab.Outline";           // string "On" / "Off"
        public const string OutlineLevelName = "VTab.OutlineLevel"; // number 1..9
        public const string NameRuleName = "VTab.NameRule";         // string "On" / "Off"
        public const string WobbleName = "VTab.Wobble";       // string "None" / "Shake" / "Wobble" / "Jelly" / "Bounce" / "Pulse" / "Swing" / "Random" (default "None")
        public const string WobbleAmpName = "VTab.WobbleAmp"; // number, percent (40 to 200, default 100)
        public const string WobbleDurName = "VTab.WobbleDur"; // number, milliseconds (300 to 1200, default 600)
        public const string WobbleSoundName = "VTab.WobbleSound";   // string "On" / "Off" (default "Off")
        public const string WobbleVolumeName = "VTab.WobbleVolume"; // number 0 to 100 (default 60)
        public const string DoneAnimName = "VTab.DoneAnim";       // string "On" / "Off" (default "On"; the burst when a sheet is marked done)
        public const string RowAnimName = "VTab.RowAnim";         // string "On" / "Off" (default "On"; the add and delete animation of the rows)
        public const string WobbleSampleStyleName = "VTab.WobbleSampleStyle"; // string "On" / "Off" (default "Off")
        public const string ActiveMarkerName = "VTab.ActiveMarker"; // string "None" / "Underline" / "Triangle" / "EdgeBars" / "CheckBadge" (default "Underline")
        public const string BadgesName = "VTab.Badges";           // string "None", or "Error", "Memo", "Protect" joined by commas (default "None")
        public const string SelectStyleName = "VTab.SelectStyle";   // string "Standard" / "DoubleFrame" / "FrostedGlass" / "GlossyGlass" / "FrostedGlass2" / "Inset" / "Stripes" / "Crystal"

        // Office.MsoDocProperties values.
        private const int PropTypeNumber = 1;
        private const int PropTypeString = 4;

        // The workbook was saved by an older version in the "vertical only" mode (VTab.Mode = "B"): its horizontal tabs are hidden.
        public bool LegacyVerticalOnly { get; private set; }

        // The display mode of this window: the PC's mode, or "horizontal only" for a window whose pane the user closed.
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

        // The rows of added sheets and folders appear and the rows of deleted ones shrink away (default: on).
        public bool RowAnim { get; set; }

        // The progress mark bursts when a sheet is marked as done (default: on).
        public bool DoneAnim { get; set; }

        public WorkbookSettings()
        {
            Mode = TabMode.Both;
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
            RowAnim = true;
            DoneAnim = true;
        }

        public static WorkbookSettings Load(Excel.Workbook wb)
        {
            var s = new WorkbookSettings();
            PcSettings.ApplyTo(s);
            try
            {
                dynamic props = wb.CustomDocumentProperties;
                dynamic mode = Find(props, ModeName);
                s.LegacyVerticalOnly = mode != null && (string)mode.Value == "B";
                dynamic nameRule = Find(props, NameRuleName);
                if (nameRule != null) s.NameRule = (string)nameRule.Value == "On";
                dynamic width = Find(props, WidthName);
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

        // The properties older versions kept in the workbook; the PC holds them now.
        private static readonly string[] ObsoleteNames =
        {
            ModeName, SideName, OutlineName, OutlineLevelName, SelectStyleName, ActiveMarkerName, BadgesName,
            WobbleName, WobbleAmpName, WobbleDurName, WobbleSoundName, WobbleVolumeName, WobbleSampleStyleName, RowAnimName, DoneAnimName
        };

        public void Save(Excel.Workbook wb)
        {
            try
            {
                dynamic props = wb.CustomDocumentProperties;
                if (WidthDip > 0) Write(props, WidthName, PropTypeNumber, WidthDip);
                Write(props, NameRuleName, PropTypeString, NameRule ? "On" : "Off");
                foreach (string name in ObsoleteNames)
                {
                    dynamic old = Find(props, name);
                    if (old != null) old.Delete();
                }
            }
            catch (Exception)
            {
                // Read-only or protected workbook: skip silently, nothing else depends on it.
            }
        }

        internal static BadgeKinds ParseBadges(string value)
        {
            var kinds = BadgeKinds.None;
            if (string.IsNullOrEmpty(value)) return kinds;
            foreach (string part in value.Split(','))
            {
                string p = part.Trim();
                if (p == "Error") kinds |= BadgeKinds.Error;
                else if (p == "Memo") kinds |= BadgeKinds.Memo;
                else if (p == "Protect") kinds |= BadgeKinds.Protect;
                else if (p == "Progress") kinds |= BadgeKinds.Progress;
            }
            return kinds;
        }

        internal static string BadgesToString(BadgeKinds kinds)
        {
            if (kinds == BadgeKinds.None) return "None";
            var parts = new System.Collections.Generic.List<string>();
            if ((kinds & BadgeKinds.Error) != 0) parts.Add("Error");
            if ((kinds & BadgeKinds.Memo) != 0) parts.Add("Memo");
            if ((kinds & BadgeKinds.Protect) != 0) parts.Add("Protect");
            if ((kinds & BadgeKinds.Progress) != 0) parts.Add("Progress");
            return string.Join(",", parts);
        }

        internal static SelectStyle ParseSelectStyle(string value)
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

        internal static string SelectStyleToString(SelectStyle style)
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
