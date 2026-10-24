using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VerticalSheetTabs
{
    // The sheets seen last, newest first (Agents.md 12.186). The sheets are known by their keys (the id in a
    // hidden custom property of the worksheet, see FolderStore.SheetKeys), so a sheet that is renamed stays in
    // the history under its new name and one that is deleted can be taken out. The cursor is the place of the
    // sheet that is shown now: "back" goes to the older ones, "forward" back to the newer ones.
    public class SheetHistory
    {
        public const int Max = 20;

        public List<string> Keys = new List<string>();
        public int Cursor;

        public bool CanBack { get { return Cursor < Keys.Count - 1; } }
        public bool CanForward { get { return Cursor > 0; } }

        // A sheet that is shown now. The sheet the cursor stands on is already there; any other one becomes the
        // newest (what lay "forward" is dropped, as in a browser). Returns true when the history changed.
        public bool Push(string key)
        {
            if (key == null) return false;
            if (Keys.Count > 0 && Cursor < Keys.Count && Keys[Cursor] == key) return false;
            if (Cursor > 0) Keys.RemoveRange(0, Cursor);
            Cursor = 0;
            Keys.RemoveAll(k => k == key);
            Keys.Insert(0, key);
            if (Keys.Count > Max) Keys.RemoveRange(Max, Keys.Count - Max);
            return true;
        }

        // The place of the next usable sheet in the direction (1 = older, -1 = newer); -1 when there is none.
        public int Find(int direction, Func<string, bool> usable)
        {
            for (int i = Cursor + direction; i >= 0 && i < Keys.Count; i += direction)
            {
                if (usable(Keys[i])) return i;
            }
            return -1;
        }

        // Takes out the sheets that no longer exist; the cursor stays on the same sheet (or on the nearest one).
        public bool Prune(ICollection<string> existing)
        {
            if (Keys.All(existing.Contains)) return false;
            string current = Cursor < Keys.Count ? Keys[Cursor] : null;
            Keys.RemoveAll(k => !existing.Contains(k));
            int at = current == null ? -1 : Keys.IndexOf(current);
            Cursor = at >= 0 ? at : Math.Min(Cursor, Math.Max(0, Keys.Count - 1));
            return true;
        }

        public const string Namespace = "urn:VerticalSheetTabs:history:1";

        public string ToXml()
        {
            XNamespace ns = Namespace;
            var root = new XElement(ns + "history");
            foreach (string key in Keys) root.Add(new XElement(ns + "sheet", new XAttribute("id", key)));
            return root.ToString(SaveOptions.DisableFormatting);
        }

        public static SheetHistory Parse(string xml)
        {
            var history = new SheetHistory();
            try
            {
                XNamespace ns = Namespace;
                XElement root = XDocument.Parse(xml).Root;
                if (root == null) return history;
                foreach (XElement e in root.Elements(ns + "sheet"))
                {
                    string id = (string)e.Attribute("id");
                    if (!string.IsNullOrEmpty(id) && !history.Keys.Contains(id)) history.Keys.Add(id);
                }
                if (history.Keys.Count > Max) history.Keys.RemoveRange(Max, history.Keys.Count - Max);
            }
            catch (Exception)
            {
            }
            return history;
        }
    }

    // Reads and writes the history in a custom XML part of its own (a change of the part marks the workbook as
    // changed, so it is written only when the workbook is saved anyway, and only when the text differs).
    public static class HistoryStore
    {
        private static Office.CustomXMLPart Find(Excel.Workbook wb)
        {
            Office.CustomXMLParts found = wb.CustomXMLParts.SelectByNamespace(SheetHistory.Namespace);
            return found != null && found.Count > 0 ? found[1] : null;
        }

        public static SheetHistory Load(Excel.Workbook wb)
        {
            try
            {
                Office.CustomXMLPart part = Find(wb);
                return part == null ? new SheetHistory() : SheetHistory.Parse(part.XML);
            }
            catch (Exception)
            {
                return new SheetHistory();
            }
        }

        public static void Save(Excel.Workbook wb, SheetHistory history)
        {
            try
            {
                Office.CustomXMLPart part = Find(wb);
                if (history.Keys.Count == 0)
                {
                    if (part != null) part.Delete();
                    return;
                }
                string xml = history.ToXml();
                if (part != null)
                {
                    if (SheetHistory.Parse(part.XML).ToXml() == xml) return;
                    part.Delete();
                }
                wb.CustomXMLParts.Add(xml);
            }
            catch (Exception)
            {
            }
        }
    }
}
