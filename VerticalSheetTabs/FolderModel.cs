using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace VerticalSheetTabs
{
    // One entry of the vertical tab list: a folder with its sheets, or a sheet outside any folder.
    public class FolderNode
    {
        // Folder: a stable id (survives renaming) and its settings. A sheet outside a folder has FolderId == null.
        public string FolderId;
        public string Name;
        public TabColorChoice Color = new TabColorChoice { Kind = TabColorKind.None };
        public bool Expanded = true;
        // Folder: all its sheets have the folder's color, also the ones that come into the folder later.
        public bool SyncColor;
        // Folder: the keys of its sheets in Excel's order (see FolderStore.SheetKeys). Loose sheet: SheetCode is its key.
        public List<string> Sheets = new List<string>();
        public string SheetCode;

        public bool IsFolder { get { return FolderId != null; } }
    }

    // The folder structure of one workbook. The nodes are in display order,
    // and that order is also the order of the sheets in Excel: the folders are ranges of sheets.
    // The model knows nothing about Excel; sheets are identified by a key (see FolderStore.SheetKeys; it does not
    // change when a sheet is renamed). It is kept in a custom XML part of the workbook.
    public class FolderModel
    {
        public const string Namespace = "urn:VerticalSheetTabs:folders:1";
        // Folder and sheet names are limited like Excel's sheet names.
        public const int MaxNameLength = 31;

        public List<FolderNode> Nodes = new List<FolderNode>();

        public bool HasFolders
        {
            get { return Nodes.Any(n => n.IsFolder); }
        }

        // All sheet codes in display order (= Excel's order).
        public List<string> Flatten()
        {
            var list = new List<string>();
            foreach (FolderNode n in Nodes)
            {
                if (n.IsFolder) list.AddRange(n.Sheets);
                else list.Add(n.SheetCode);
            }
            return list;
        }

        // The folder a sheet belongs to, or null.
        public FolderNode FolderOf(string sheetCode)
        {
            return Nodes.FirstOrDefault(n => n.IsFolder && n.Sheets.Contains(sheetCode));
        }

        // ---- XML ----

        public string ToXml()
        {
            XNamespace ns = Namespace;
            var root = new XElement(ns + "folders", new XAttribute("version", 1));
            foreach (FolderNode n in Nodes)
            {
                if (!n.IsFolder)
                {
                    root.Add(new XElement(ns + "sheet", new XAttribute("code", n.SheetCode)));
                    continue;
                }
                var f = new XElement(ns + "folder",
                    new XAttribute("id", n.FolderId),
                    new XAttribute("name", n.Name ?? ""),
                    new XAttribute("expanded", n.Expanded ? 1 : 0));
                if (n.SyncColor) f.Add(new XAttribute("syncColor", 1));
                TabColorChoice c = n.Color ?? new TabColorChoice { Kind = TabColorKind.None };
                f.Add(new XAttribute("color", c.Kind == TabColorKind.Rgb ? "rgb" : c.Kind == TabColorKind.Theme ? "theme" : "none"));
                if (c.Kind == TabColorKind.Rgb) f.Add(new XAttribute("bgr", c.Bgr));
                if (c.Kind == TabColorKind.Theme)
                {
                    f.Add(new XAttribute("theme", c.ThemeIndex));
                    f.Add(new XAttribute("tint", c.Tint.ToString("R", CultureInfo.InvariantCulture)));
                }
                foreach (string code in n.Sheets) f.Add(new XElement(ns + "sheet", new XAttribute("code", code)));
                root.Add(f);
            }
            return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), root).ToString(SaveOptions.DisableFormatting);
        }

        // Reads the XML of the custom part. Anything unreadable gives an empty model (no folders).
        public static FolderModel Parse(string xml)
        {
            var model = new FolderModel();
            if (string.IsNullOrEmpty(xml)) return model;
            try
            {
                XNamespace ns = Namespace;
                XElement root = XDocument.Parse(xml).Root;
                if (root == null || root.Name != ns + "folders") return model;

                var seenSheets = new HashSet<string>();
                var seenIds = new HashSet<string>();
                foreach (XElement e in root.Elements())
                {
                    if (e.Name == ns + "sheet")
                    {
                        string code = (string)e.Attribute("code");
                        if (!string.IsNullOrEmpty(code) && seenSheets.Add(code))
                            model.Nodes.Add(new FolderNode { SheetCode = code });
                    }
                    else if (e.Name == ns + "folder")
                    {
                        string id = (string)e.Attribute("id");
                        if (string.IsNullOrEmpty(id) || !seenIds.Add(id)) id = Guid.NewGuid().ToString("N");
                        var node = new FolderNode
                        {
                            FolderId = id,
                            Name = Truncate((string)e.Attribute("name") ?? ""),
                            Expanded = ((string)e.Attribute("expanded")) != "0",
                            SyncColor = ((string)e.Attribute("syncColor")) == "1",
                            Color = ParseColor(e)
                        };
                        foreach (XElement s in e.Elements(ns + "sheet"))
                        {
                            string code = (string)s.Attribute("code");
                            if (!string.IsNullOrEmpty(code) && seenSheets.Add(code)) node.Sheets.Add(code);
                        }
                        model.Nodes.Add(node);
                    }
                }
            }
            catch (Exception)
            {
                return new FolderModel();
            }
            return model;
        }

        private static TabColorChoice ParseColor(XElement e)
        {
            string kind = (string)e.Attribute("color");
            try
            {
                if (kind == "rgb") return new TabColorChoice { Kind = TabColorKind.Rgb, Bgr = (int)e.Attribute("bgr") };
                if (kind == "theme")
                {
                    return new TabColorChoice
                    {
                        Kind = TabColorKind.Theme,
                        ThemeIndex = (int)e.Attribute("theme"),
                        Tint = double.Parse((string)e.Attribute("tint"), CultureInfo.InvariantCulture)
                    };
                }
            }
            catch (Exception)
            {
                // A damaged color: no color.
            }
            return new TabColorChoice { Kind = TabColorKind.None };
        }

        private static string Truncate(string name)
        {
            return name.Length > MaxNameLength ? name.Substring(0, MaxNameLength) : name;
        }

        // ---- keeping the model and Excel's real sheet order together ----

        // Brings the model in line with the sheets that exist now, in Excel's order (keys of
        // all sheets, hidden ones included). Returns true when the model changed.
        //  - sheets that no longer exist are dropped;
        //  - sheets the model does not know (added, copied, moved in the horizontal tabs) are placed
        //    where Excel has them: between two sheets of one folder, they join that folder, anywhere
        //    else they are outside any folder;
        //  - Excel's order is the truth. A folder is one run of consecutive sheets: when the members
        //    of a folder were separated (a sheet was moved away in the horizontal tabs), the longest
        //    run stays the folder (the first one when equally long) and the others leave it;
        //  - folders without sheets are kept, in front of the entry that followed them.
        public bool Reconcile(IList<string> excelOrder)
        {
            string before = ToXml();
            int count = excelOrder.Count;

            // membership of every sheet as the model knows it
            var memberOf = new Dictionary<string, FolderNode>();
            foreach (FolderNode n in Nodes)
            {
                if (!n.IsFolder) continue;
                foreach (string code in n.Sheets) memberOf[code] = n;
            }
            var membership = new FolderNode[count];
            for (int i = 0; i < count; i++)
            {
                FolderNode f;
                membership[i] = memberOf.TryGetValue(excelOrder[i], out f) ? f : null;
            }

            // 1. sheets the model does not know, between two sheets of the same folder, join it
            //    (sheets the model knows as outside any folder stay outside)
            var known = new HashSet<string>(memberOf.Keys);
            foreach (FolderNode n in Nodes.Where(n => !n.IsFolder)) known.Add(n.SheetCode);
            int p = 0;
            while (p < count)
            {
                if (membership[p] != null) { p++; continue; }
                int q = p;
                bool allUnknown = true;
                while (q < count && membership[q] == null) { if (known.Contains(excelOrder[q])) allUnknown = false; q++; }
                if (allUnknown && p > 0 && q < count && membership[p - 1] == membership[q])
                {
                    for (int k = p; k < q; k++) membership[k] = membership[p - 1];
                }
                p = q;
            }

            // 2. a folder is one run: keep the longest, the others become loose
            foreach (FolderNode f in Nodes.Where(n => n.IsFolder))
            {
                var runs = new List<int[]>(); // start, length
                int i = 0;
                while (i < count)
                {
                    if (membership[i] != f) { i++; continue; }
                    int start = i;
                    while (i < count && membership[i] == f) i++;
                    runs.Add(new[] { start, i - start });
                }
                if (runs.Count <= 1) continue;
                int[] keep = runs.OrderByDescending(r => r[1]).ThenBy(r => r[0]).First();
                foreach (int[] r in runs)
                {
                    if (r == keep) continue;
                    for (int k = r[0]; k < r[0] + r[1]; k++) membership[k] = null;
                }
            }

            // 3. rebuild the nodes in Excel's order
            var rebuilt = new List<FolderNode>();
            var placed = new Dictionary<FolderNode, FolderNode>();
            for (int i = 0; i < count; i++)
            {
                FolderNode f = membership[i];
                if (f == null)
                {
                    rebuilt.Add(new FolderNode { SheetCode = excelOrder[i] });
                    continue;
                }
                FolderNode target;
                if (!placed.TryGetValue(f, out target))
                {
                    target = new FolderNode { FolderId = f.FolderId, Name = f.Name, Color = f.Color, Expanded = f.Expanded, SyncColor = f.SyncColor };
                    placed[f] = target;
                    rebuilt.Add(target);
                }
                target.Sheets.Add(excelOrder[i]);
            }

            // 4. empty folders stay where they were: in front of the next entry that still exists
            var oldNodes = Nodes;
            for (int index = 0; index < oldNodes.Count; index++)
            {
                FolderNode f = oldNodes[index];
                if (!f.IsFolder || placed.ContainsKey(f)) continue;
                var empty = new FolderNode { FolderId = f.FolderId, Name = f.Name, Color = f.Color, Expanded = f.Expanded, SyncColor = f.SyncColor };
                int at = rebuilt.Count;
                for (int next = index + 1; next < oldNodes.Count; next++)
                {
                    FolderNode o = oldNodes[next];
                    int found = o.IsFolder
                        ? rebuilt.FindIndex(r => r.IsFolder && r.FolderId == o.FolderId)
                        : rebuilt.FindIndex(r => !r.IsFolder && r.SheetCode == o.SheetCode);
                    if (found >= 0) { at = found; break; }
                }
                rebuilt.Insert(at, empty);
                placed[f] = empty;
            }

            Nodes = rebuilt;
            return ToXml() != before;
        }

        // ---- editing (the caller then puts Excel's sheet order in line with Flatten()) ----

        public FolderNode FindFolder(string folderId)
        {
            return Nodes.FirstOrDefault(n => n.IsFolder && n.FolderId == folderId);
        }

        public static FolderNode NewFolder(string name)
        {
            return new FolderNode { FolderId = Guid.NewGuid().ToString("N"), Name = name };
        }

        // A new folder holding these sheets, placed where the first of them is now.
        public FolderNode CreateFolder(string name, IList<string> sheetKeys)
        {
            List<string> order = Flatten();
            string first = order.FirstOrDefault(k => sheetKeys.Contains(k));
            FolderNode folder = NewFolder(name);
            FolderNode top = first == null ? null : Nodes.FirstOrDefault(n => n.IsFolder ? n.Sheets.Contains(first) : n.SheetCode == first);
            int at = top == null ? Nodes.Count : Nodes.IndexOf(top);
            Nodes.Insert(at, folder);
            if (first != null) Move(sheetKeys, new string[0], folder.FolderId, null, null);
            return folder;
        }

        // Deletes the folder only: its sheets stay, where they are, outside any folder.
        public bool RemoveFolder(string folderId)
        {
            int i = Nodes.FindIndex(n => n.IsFolder && n.FolderId == folderId);
            if (i < 0) return false;
            FolderNode f = Nodes[i];
            Nodes.RemoveAt(i);
            Nodes.InsertRange(i, f.Sheets.Select(k => new FolderNode { SheetCode = k }));
            return true;
        }

        // Takes sheets out of the model (they were deleted in Excel).
        public void RemoveSheets(ICollection<string> sheetKeys)
        {
            foreach (FolderNode n in Nodes.Where(n => n.IsFolder)) n.Sheets.RemoveAll(sheetKeys.Contains);
            Nodes.RemoveAll(n => !n.IsFolder && sheetKeys.Contains(n.SheetCode));
        }

        // The sheets of the given folders, in display order.
        public List<string> SheetsOfFolders(ICollection<string> folderIds)
        {
            var list = new List<string>();
            foreach (FolderNode n in Nodes)
            {
                if (n.IsFolder && folderIds.Contains(n.FolderId)) list.AddRange(n.Sheets);
            }
            return list;
        }

        // Moves sheets and whole folders (with their sheets) to a new place. The moved things keep
        // the order they have now. Sheets of a moved folder travel with it.
        //  - intoFolderId != null: the sheets go into that folder, in front of beforeSheet (null = at the end);
        //    folders cannot go into a folder, so moved folders are ignored then;
        //  - otherwise everything goes to the top level, in front of the loose sheet beforeSheet, or the
        //    folder beforeFolderId, or at the end when both are null. Sheets arriving there are outside any folder.
        // Returns false when nothing was moved.
        public bool Move(IList<string> sheetKeys, IList<string> folderIds, string intoFolderId, string beforeSheet, string beforeFolderId)
        {
            var movedFolders = new HashSet<string>(folderIds);
            if (intoFolderId != null)
            {
                if (movedFolders.Contains(intoFolderId)) return false;
                movedFolders.Clear();
            }
            var want = new HashSet<string>(sheetKeys);

            // what moves, in display order: top-level units (loose sheets, whole folders, sheets out of folders)
            var old = new List<FolderNode>(Nodes);
            var units = new List<FolderNode>();
            var movedSheets = new HashSet<string>();
            foreach (FolderNode n in old)
            {
                if (!n.IsFolder)
                {
                    if (want.Contains(n.SheetCode)) { units.Add(n); movedSheets.Add(n.SheetCode); }
                }
                else if (movedFolders.Contains(n.FolderId)) units.Add(n);
                else
                {
                    foreach (string key in n.Sheets)
                    {
                        if (want.Contains(key)) { units.Add(new FolderNode { SheetCode = key }); movedSheets.Add(key); }
                    }
                }
            }
            if (units.Count == 0) return false;

            // an anchor that moves itself is replaced by the next top-level entry that stays
            Func<FolderNode, bool> moves = n => n.IsFolder ? movedFolders.Contains(n.FolderId) : movedSheets.Contains(n.SheetCode);
            FolderNode anchor = null;
            if (intoFolderId == null)
            {
                if (beforeSheet != null) anchor = old.FirstOrDefault(n => !n.IsFolder && n.SheetCode == beforeSheet);
                else if (beforeFolderId != null) anchor = old.FirstOrDefault(n => n.IsFolder && n.FolderId == beforeFolderId);
                if (anchor != null && moves(anchor))
                {
                    anchor = old.Skip(old.IndexOf(anchor) + 1).FirstOrDefault(n => !moves(n));
                }
            }

            if (intoFolderId != null && beforeSheet != null && movedSheets.Contains(beforeSheet))
            {
                FolderNode inFolder = FindFolder(intoFolderId);
                beforeSheet = inFolder == null ? null
                    : inFolder.Sheets.Skip(inFolder.Sheets.IndexOf(beforeSheet) + 1).FirstOrDefault(k => !movedSheets.Contains(k));
            }

            if (intoFolderId != null && FindFolder(intoFolderId) == null) return false;

            // take everything out
            foreach (FolderNode n in Nodes.Where(n => n.IsFolder)) n.Sheets.RemoveAll(movedSheets.Contains);
            Nodes.RemoveAll(n => moves(n));

            if (intoFolderId != null)
            {
                FolderNode target = FindFolder(intoFolderId);
                if (target == null) { Nodes = old; return false; }
                int at = beforeSheet == null ? -1 : target.Sheets.IndexOf(beforeSheet);
                var keys = units.Select(u => u.SheetCode).ToList();
                if (at < 0) target.Sheets.AddRange(keys);
                else target.Sheets.InsertRange(at, keys);
                return true;
            }

            int index = anchor == null ? Nodes.Count : Nodes.IndexOf(anchor);
            Nodes.InsertRange(index < 0 ? Nodes.Count : index, units);
            return true;
        }
    }
}
