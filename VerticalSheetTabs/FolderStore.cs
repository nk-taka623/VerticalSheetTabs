using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace VerticalSheetTabs
{
    // Reads and writes the folder model in the workbook's custom XML part (namespace FolderModel.Namespace).
    // Changing a custom XML part marks the workbook as changed, so Save() does nothing when the
    // XML is the same as the stored one.
    public static class FolderStore
    {
        // The parts of this namespace in the order Excel keeps them (a part that is added comes last).
        private static List<Office.CustomXMLPart> FindAll(Excel.Workbook wb)
        {
            var list = new List<Office.CustomXMLPart>();
            Office.CustomXMLParts found = wb.CustomXMLParts.SelectByNamespace(FolderModel.Namespace);
            if (found == null) return list;
            for (int i = 1; i <= found.Count; i++) list.Add(found[i]);
            return list;
        }

        public static FolderModel Load(Excel.Workbook wb)
        {
            try
            {
                List<Office.CustomXMLPart> parts = FindAll(wb);
                // (the newest one: a part that could not be deleted after a newer one was added is old)
                return parts.Count == 0 ? new FolderModel() : FolderModel.Parse(parts[parts.Count - 1].XML);
            }
            catch (Exception)
            {
                return new FolderModel();
            }
        }

        // Writes the model. A model without folders removes the part, unless the part was unreadable (keepUnreadable): what could not be
        // read is not known to be empty, and stays. A changed model is added first and the old part is deleted afterwards, so that a
        // failure on the way leaves the old folders. Returns true when the workbook was changed.
        public static bool Save(Excel.Workbook wb, FolderModel model, bool keepUnreadable = false)
        {
            try
            {
                List<Office.CustomXMLPart> parts = FindAll(wb);
                if (!model.HasFolders)
                {
                    if (parts.Count == 0 || keepUnreadable) return false;
                    foreach (Office.CustomXMLPart p in parts) p.Delete();
                    return true;
                }
                string xml = model.ToXml();
                int keep = parts.Count; // the parts before this index are old ones to delete
                if (parts.Count > 0)
                {
                    FolderModel current = FolderModel.Parse(parts[parts.Count - 1].XML);
                    if (!current.Unreadable && current.ToXml() == xml)
                    {
                        if (parts.Count == 1) return false;
                        keep = parts.Count - 1; // (the same XML twice: the older ones go)
                        for (int i = 0; i < keep; i++) DeleteQuietly(parts[i]);
                        return true;
                    }
                }
                wb.CustomXMLParts.Add(xml);
                foreach (Office.CustomXMLPart p in parts) DeleteQuietly(p);
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static void DeleteQuietly(Office.CustomXMLPart part)
        {
            try { part.Delete(); }
            catch (Exception) { } // (the newer part wins at the next read)
        }

        public const string IdProperty = "VTab.Id";

        // Keys of all sheets (hidden ones too) in Excel's order. Sheets are identified by a hidden
        // custom property of the worksheet (VTab.Id); CodeName cannot be used, because Excel returns
        // it empty for workbooks without a VBA project. Chart sheets have no custom properties,
        // so their key is their name ("n:" + name). With assign, worksheets without an id get one.
        // Giving an id changes the workbook; with keepSaved a workbook that was unchanged stays unchanged
        // (for the background sync: opening a workbook must not make Excel ask "save changes?").
        public static List<string> SheetKeys(Excel.Workbook wb, bool assign, bool keepSaved = false)
        {
            bool complete;
            return SheetKeys(wb, assign, keepSaved, out complete);
        }

        // complete is false when a sheet could not be read (its key is then only a guess, "u:" + index).
        public static List<string> SheetKeys(Excel.Workbook wb, bool assign, bool keepSaved, out bool complete)
        {
            complete = true;
            var list = new List<string>();
            int index = 0;
            bool wasSaved = false;
            bool assigned = false;
            try { wasSaved = wb.Saved; } catch (Exception) { }
            foreach (object o in wb.Sheets)
            {
                index++;
                string key = "u:" + index;
                try
                {
                    var ws = o as Excel.Worksheet;
                    if (ws != null)
                    {
                        string id = ReadId(ws);
                        if (id == null && assign)
                        {
                            id = Guid.NewGuid().ToString("N");
                            ws.CustomProperties.Add(IdProperty, id);
                            assigned = true;
                        }
                        if (id != null) key = id;
                    }
                    else
                    {
                        var ch = o as Excel.Chart;
                        if (ch != null) key = "n:" + ch.Name;
                    }
                }
                catch (Exception)
                {
                    complete = false;
                }
                list.Add(key);
            }
            if (assigned && keepSaved && wasSaved)
            {
                try { wb.Saved = true; } catch (Exception) { }
            }
            return list;
        }

        // The key of one sheet (a worksheet gets its id when it has none and assign is set; the key of a chart sheet is its name).
        public static string KeyOf(object sheet, bool assign)
        {
            try
            {
                var ws = sheet as Excel.Worksheet;
                if (ws != null)
                {
                    string id = ReadId(ws);
                    if (id == null && assign)
                    {
                        id = Guid.NewGuid().ToString("N");
                        ws.CustomProperties.Add(IdProperty, id);
                    }
                    return id;
                }
                var ch = sheet as Excel.Chart;
                return ch != null ? "n:" + ch.Name : null;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private static string ReadId(Excel.Worksheet ws)
        {
            foreach (Excel.CustomProperty p in ws.CustomProperties)
            {
                if (p.Name == IdProperty) return Convert.ToString(p.Value);
            }
            return null;
        }
    }
}
