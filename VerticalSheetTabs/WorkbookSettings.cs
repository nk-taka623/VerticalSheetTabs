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

        // Office.MsoDocProperties values.
        private const int PropTypeNumber = 1;
        private const int PropTypeString = 4;

        // No add-in property in the workbook: it is the first time it is opened with the add-in.
        public bool IsFirstOpen { get; private set; }

        // Until Step4 lets the user choose, new workbooks show both tab bars (non-destructive).
        public TabMode Mode { get; set; }
        public TabSide Side { get; set; }
        // 0 means "not stored": the caller uses the default (12 half-width characters).
        public int WidthDip { get; set; }
        // The dashed outline around an open folder: shown or not, and its color level (1..9, 5 = middle).
        public bool OutlineVisible { get; set; }
        public int OutlineLevel { get; set; }
        // The sheet-name rule: a sheet "A | B" is the sheet B in the folder A (default: not used).
        public bool NameRule { get; set; }

        public WorkbookSettings()
        {
            IsFirstOpen = true;
            Mode = TabMode.Both;
            Side = TabSide.Left;
            WidthDip = 0;
            OutlineVisible = true;
            OutlineLevel = 5;
            NameRule = false;
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
            }
            catch (Exception)
            {
                // Read-only or protected workbook: skip silently, nothing else depends on it.
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
