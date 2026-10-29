using System;
using System.Collections.Generic;

namespace VerticalSheetTabs
{
    // The sheet-name rule: a sheet named "A | B" is the sheet "B" in the
    // folder "A" of the vertical tabs. These are the pure parts of it (parsing, grouping, building the
    // folder model from the names); PaneManager applies them to the workbook.
    public static class NameRule
    {
        public const string Separator = " | ";
        // Excel's longest sheet name (31) minus the separator (3): folder name + sheet name.
        public const int MaxCombined = 28;
        public const int MaxSheetName = 31;

        // "A | B" -> folder "A", sheet "B". Both parts must be there.
        public static bool TryParse(string name, out string folder, out string sheet)
        {
            folder = null;
            sheet = name;
            if (name == null) return false;
            int i = name.IndexOf(Separator, StringComparison.Ordinal);
            if (i <= 0 || i + Separator.Length >= name.Length) return false;
            folder = name.Substring(0, i);
            sheet = name.Substring(i + Separator.Length);
            return true;
        }

        public static bool IsFolderSheet(string name)
        {
            string f, s;
            return TryParse(name, out f, out s);
        }

        // The sheet's own name: "B" for "A | B", the name itself otherwise.
        public static string BaseOf(string name)
        {
            string f, s;
            return TryParse(name, out f, out s) ? s : name;
        }

        public static string Compose(string folder, string sheet)
        {
            return folder + Separator + sheet;
        }

        public static bool HasSeparator(string text)
        {
            return text != null && text.IndexOf(Separator, StringComparison.Ordinal) >= 0;
        }

        public static bool SameFolder(string a, string b)
        {
            return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
        }

        // A sheet of a folder that is not next to the others of its folder goes to the right of the
        // nearest sheet to its left whose name starts with the same "A | ". Returns the moves in the
        // order they are made (a sheet at index From goes to index To, counted after it was taken out; To
        // is always smaller than From); names is changed to the resulting order.
        public static List<KeyValuePair<int, int>> GroupingMoves(List<string> names)
        {
            var moves = new List<KeyValuePair<int, int>>();
            for (int i = 0; i < names.Count; i++)
            {
                string folder, s;
                if (!TryParse(names[i], out folder, out s)) continue;
                string leftFolder, leftSheet;
                if (i > 0 && TryParse(names[i - 1], out leftFolder, out leftSheet) && SameFolder(leftFolder, folder)) continue;

                int j = -1;
                for (int k = i - 1; k >= 0; k--)
                {
                    string kf, ks;
                    if (TryParse(names[k], out kf, out ks) && SameFolder(kf, folder)) { j = k; break; }
                }
                if (j < 0) continue;
                string moved = names[i];
                names.RemoveAt(i);
                names.Insert(j + 1, moved);
                moves.Add(new KeyValuePair<int, int>(i, j + 1));
            }
            return moves;
        }

        // The folder model that the names describe. The folders of the old model with the same name keep
        // their identity, color and open/closed state.
        public static FolderModel BuildModel(IList<string> names, IList<string> keys, FolderModel old)
        {
            var model = new FolderModel();
            var oldByName = new Dictionary<string, FolderNode>(StringComparer.OrdinalIgnoreCase);
            if (old != null)
            {
                foreach (FolderNode n in old.Nodes)
                {
                    if (n.IsFolder && n.Name != null && !oldByName.ContainsKey(n.Name)) oldByName[n.Name] = n;
                }
            }
            var used = new HashSet<FolderNode>();

            FolderNode current = null;
            for (int i = 0; i < names.Count && i < keys.Count; i++)
            {
                string folder, s;
                if (TryParse(names[i], out folder, out s))
                {
                    if (current == null || !SameFolder(current.Name, folder))
                    {
                        FolderNode reuse;
                        if (oldByName.TryGetValue(folder, out reuse) && !used.Contains(reuse))
                        {
                            reuse.Sheets.Clear();
                            reuse.Name = folder;
                            current = reuse;
                        }
                        else current = FolderModel.NewFolder(folder);
                        used.Add(current);
                        model.Nodes.Add(current);
                    }
                    current.Sheets.Add(keys[i]);
                }
                else
                {
                    current = null;
                    model.Nodes.Add(new FolderNode { SheetCode = keys[i] });
                }
            }
            return model;
        }
    }
}
