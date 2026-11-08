using System;
using System.Collections.Generic;
using System.Linq;

namespace VerticalSheetTabs.Tests
{
    // A small test runner: each Test(...) runs one check and prints the failures. The exit code is the number of failures.
    internal static class Program
    {
        private static int _run;
        private static int _failed;

        private static void Test(string name, Action body)
        {
            _run++;
            try
            {
                body();
            }
            catch (Exception e)
            {
                _failed++;
                Console.WriteLine("FAIL " + name + ": " + e.Message);
            }
        }

        private static void Equal<T>(T expected, T actual)
        {
            if (!EqualityComparer<T>.Default.Equals(expected, actual)) throw new Exception("expected <" + expected + "> but was <" + actual + ">");
        }

        private static void Seq(IEnumerable<string> expected, IEnumerable<string> actual)
        {
            string e = string.Join(",", expected), a = string.Join(",", actual);
            if (e != a) throw new Exception("expected [" + e + "] but was [" + a + "]");
        }

        private static void True(bool value, string what)
        {
            if (!value) throw new Exception(what);
        }

        private static void Throws(Action body, string what)
        {
            try { body(); }
            catch (Exception) { return; }
            throw new Exception(what + " did not throw");
        }

        private static int Main()
        {
            // ---- NameRule
            Test("TryParse splits at the first separator", () =>
            {
                string f, s;
                True(NameRule.TryParse("A | B | C", out f, out s), "parse");
                Equal("A", f); Equal("B | C", s);
            });
            Test("TryParse rejects an empty side", () =>
            {
                string f, s;
                True(!NameRule.TryParse(" | B", out f, out s), "empty folder");
                True(!NameRule.TryParse("A | ", out f, out s), "empty sheet");
                True(!NameRule.TryParse("A|B", out f, out s), "no spaces");
                True(!NameRule.TryParse(null, out f, out s), "null");
            });
            Test("Compose and BaseOf", () =>
            {
                Equal("A | B", NameRule.Compose("A", "B"));
                Equal("B", NameRule.BaseOf("A | B"));
                Equal("plain", NameRule.BaseOf("plain"));
            });
            Test("GroupingMoves puts a stray sheet next to its folder", () =>
            {
                var names = new List<string> { "A | 1", "x", "A | 2" };
                var moves = NameRule.GroupingMoves(names);
                Seq(new[] { "A | 1", "A | 2", "x" }, names);
                Equal(1, moves.Count);
                Equal(2, moves[0].Key); Equal(1, moves[0].Value);
            });
            Test("GroupingMoves leaves a grouped order alone", () =>
            {
                var names = new List<string> { "A | 1", "A | 2", "x", "B | 1" };
                Equal(0, NameRule.GroupingMoves(names).Count);
            });
            Test("BuildModel makes folders from names and keeps the old identity", () =>
            {
                var old = new FolderModel();
                FolderNode a = FolderModel.NewFolder("A");
                a.Expanded = false;
                old.Nodes.Add(a);
                var model = NameRule.BuildModel(new[] { "A | 1", "A | 2", "x" }, new[] { "k1", "k2", "k3" }, old);
                Equal(2, model.Nodes.Count);
                True(model.Nodes[0].IsFolder, "folder first");
                Equal(a.FolderId, model.Nodes[0].FolderId);
                True(!model.Nodes[0].Expanded, "state kept");
                Seq(new[] { "k1", "k2" }, model.Nodes[0].Sheets);
                Equal("k3", model.Nodes[1].SheetCode);
            });

            // ---- FolderModel
            Test("FolderModel XML round trip", () =>
            {
                var m = new FolderModel();
                FolderNode f = FolderModel.NewFolder("Folder 1");
                f.Sheets.Add("s1"); f.Sheets.Add("s2"); f.Expanded = false; f.SyncColor = true;
                m.Nodes.Add(f);
                m.Nodes.Add(new FolderNode { SheetCode = "s3" });
                FolderModel back = FolderModel.Parse(m.ToXml());
                True(!back.Unreadable, "readable");
                Equal(2, back.Nodes.Count);
                Equal("Folder 1", back.Nodes[0].Name);
                True(!back.Nodes[0].Expanded && back.Nodes[0].SyncColor, "flags");
                Seq(new[] { "s1", "s2" }, back.Nodes[0].Sheets);
                Equal("s3", back.Nodes[1].SheetCode);
                Equal(m.ToXml(), back.ToXml());
            });
            Test("FolderModel.Parse marks garbage as unreadable", () =>
            {
                True(FolderModel.Parse("<not xml").Unreadable, "broken xml");
                True(FolderModel.Parse("<other xmlns=\"urn:x\"/>").Unreadable, "wrong root");
                True(!FolderModel.Parse("").Unreadable, "empty is no folders");
                True(!FolderModel.Parse("").HasFolders, "empty has no folders");
            });
            Test("FolderModel.Parse refuses a DTD", () =>
            {
                string xml = "<!DOCTYPE x [<!ENTITY e \"a\">]><folders xmlns=\"" + FolderModel.Namespace + "\"/>";
                True(FolderModel.Parse(xml).Unreadable, "DTD");
            });
            Test("FolderModel.Parse drops duplicate sheets", () =>
            {
                string xml = "<folders xmlns=\"" + FolderModel.Namespace + "\"><sheet code=\"a\"/><sheet code=\"a\"/><folder id=\"f\" name=\"F\"><sheet code=\"b\"/></folder></folders>";
                Equal(2, FolderModel.Parse(xml).Nodes.Count);
            });
            Test("Reconcile follows the order of the workbook", () =>
            {
                var m = new FolderModel();
                FolderNode f = FolderModel.NewFolder("F");
                f.Sheets.Add("a"); f.Sheets.Add("b");
                m.Nodes.Add(f);
                m.Nodes.Add(new FolderNode { SheetCode = "c" });
                m.Reconcile(new[] { "c", "a", "b" });
                Seq(new[] { "c", "a", "b" }, m.Flatten());
                Seq(new[] { "a", "b" }, m.FolderOf("a").Sheets);
            });
            Test("Reconcile adds new sheets and drops deleted ones", () =>
            {
                var m = new FolderModel();
                FolderNode f = FolderModel.NewFolder("F");
                f.Sheets.Add("a"); f.Sheets.Add("gone");
                m.Nodes.Add(f);
                m.Reconcile(new[] { "a", "new" });
                Seq(new[] { "a", "new" }, m.Flatten());
                Seq(new[] { "a" }, m.FindFolder(f.FolderId).Sheets);
            });
            Test("CreateFolder and RemoveFolder keep the sheets", () =>
            {
                var m = new FolderModel();
                m.Nodes.Add(new FolderNode { SheetCode = "a" });
                m.Nodes.Add(new FolderNode { SheetCode = "b" });
                FolderNode f = m.CreateFolder("F", new[] { "b" });
                Seq(new[] { "b" }, f.Sheets);
                True(m.RemoveFolder(f.FolderId), "removed");
                Seq(new[] { "a", "b" }, m.Flatten().OrderBy(x => x));
            });

            // ---- SafeXml
            Test("SafeXml parses plain XML", () =>
            {
                Equal("r", SafeXml.Parse("<r><c/></r>", 1000).Root.Name.LocalName);
            });
            Test("SafeXml refuses a DTD", () =>
            {
                Throws(() => SafeXml.Parse("<!DOCTYPE r [<!ENTITY e \"a\">]><r>&e;</r>", 1000), "DTD");
            });
            Test("SafeXml refuses a too long document", () =>
            {
                string big = "<r>" + new string('x', 5000) + "</r>";
                Throws(() => SafeXml.Parse(big, 1000), "long document");
            });


            // ---- Wallpaper layout (picture 90 x 120 in an area of 150 x 200)
            Func<WallpaperFit, bool, bool, WallpaperHAlign, WallpaperVAlign, WallpaperLayout> lay = (fit, keep, tile, h, v) =>
                Wallpaper.Layout(90, 120, 150, 200, new WallpaperOptions { Fit = fit, KeepRatio = keep, Tile = tile, HAlign = h, VAlign = v });
            Test("Wallpaper original size is centered", () =>
            {
                WallpaperLayout l = lay(WallpaperFit.Original, true, false, WallpaperHAlign.Center, WallpaperVAlign.Center);
                Equal(90, l.Width); Equal(120, l.Height); Equal(30, l.X); Equal(40, l.Y);
            });
            Test("Wallpaper original size at the right bottom", () =>
            {
                WallpaperLayout l = lay(WallpaperFit.Original, true, false, WallpaperHAlign.Right, WallpaperVAlign.Bottom);
                Equal(60, l.X); Equal(80, l.Y);
            });
            Test("Wallpaper fit width keeps the ratio and places only vertically", () =>
            {
                WallpaperLayout l = lay(WallpaperFit.FitWidth, true, false, WallpaperHAlign.Right, WallpaperVAlign.Top);
                Equal(150, l.Width); Equal(200, l.Height); Equal(0, l.X); Equal(0, l.Y);
                l = Wallpaper.Layout(90, 120, 150, 120, new WallpaperOptions { Fit = WallpaperFit.FitWidth, VAlign = WallpaperVAlign.Bottom });
                Equal(200, l.Height); Equal(-80, l.Y);
            });
            Test("Wallpaper fit height keeps the ratio and places only horizontally", () =>
            {
                WallpaperLayout l = Wallpaper.Layout(90, 120, 100, 200, new WallpaperOptions { Fit = WallpaperFit.FitHeight, HAlign = WallpaperHAlign.Center, VAlign = WallpaperVAlign.Bottom });
                Equal(150, l.Width); Equal(200, l.Height); Equal(-25, l.X); Equal(0, l.Y);
            });
            Test("Wallpaper stretch fills the area", () =>
            {
                WallpaperLayout l = lay(WallpaperFit.FitBoth, false, true, WallpaperHAlign.Right, WallpaperVAlign.Bottom);
                Equal(150, l.Width); Equal(200, l.Height); Equal(0, l.X); Equal(0, l.Y); True(!l.Tile, "no tile when fitting both");
            });
            Test("Wallpaper cover keeps the ratio and can be placed", () =>
            {
                WallpaperLayout l = Wallpaper.Layout(90, 120, 150, 120, new WallpaperOptions { Fit = WallpaperFit.FitBoth, KeepRatio = true, HAlign = WallpaperHAlign.Left, VAlign = WallpaperVAlign.Center });
                Equal(150, l.Width); Equal(200, l.Height); Equal(0, l.X); Equal(-40, l.Y);
            });
            Test("Wallpaper tile only applies to the fits that allow it", () =>
            {
                True(lay(WallpaperFit.Original, true, true, WallpaperHAlign.Left, WallpaperVAlign.Top).Tile, "original tiles");
                True(!lay(WallpaperFit.FitBoth, true, true, WallpaperHAlign.Left, WallpaperVAlign.Top).Tile, "cover does not tile");
            });
            Test("Wallpaper settings that do not apply are disabled", () =>
            {
                True(!WallpaperOptions.HAlignApplies(WallpaperFit.FitWidth, true), "width: no horizontal");
                True(!WallpaperOptions.VAlignApplies(WallpaperFit.FitHeight, true), "height: no vertical");
                True(!WallpaperOptions.HAlignApplies(WallpaperFit.FitBoth, false), "stretch: no position");
                True(WallpaperOptions.HAlignApplies(WallpaperFit.FitBoth, true) && WallpaperOptions.VAlignApplies(WallpaperFit.FitBoth, true), "cover: both");
                True(WallpaperOptions.KeepRatioApplies(WallpaperFit.FitBoth) && !WallpaperOptions.KeepRatioApplies(WallpaperFit.Original), "ratio only for fit both");
            });
            Test("Wallpaper accepts jpg and png files only", () =>
            {
                True(Wallpaper.IsSupportedFile("a.JPG") && Wallpaper.IsSupportedFile("b.jpeg") && Wallpaper.IsSupportedFile("c.png"), "supported");
                True(!Wallpaper.IsSupportedFile("d.gif") && !Wallpaper.IsSupportedFile("") && !Wallpaper.IsSupportedFile(null), "unsupported");
                True(Wallpaper.Load(@"C:\no\such\file.png") == null, "missing file");
            });
            Test("Row transparency steps", () =>
            {
                Equal(0, RowTransparencyScale.Clamp(-5)); Equal(45, RowTransparencyScale.Clamp(47)); Equal(90, RowTransparencyScale.Clamp(200));
            });

            // ---- RowPlacement (area 100 px high, rows 32 px)
            Test("StackOffset is 0 when hanging", () => Equal(0, RowPlacement.StackOffset(false, 100, 1, 32)));
            Test("StackOffset pushes the rows to the bottom", () =>
            {
                Equal(68, RowPlacement.StackOffset(true, 100, 1, 32));
                Equal(4, RowPlacement.StackOffset(true, 100, 3, 32));
                Equal(100, RowPlacement.StackOffset(true, 100, 0, 32));
            });
            Test("StackOffset is 0 when the rows fill or exceed the area", () =>
            {
                Equal(0, RowPlacement.StackOffset(true, 96, 3, 32));
                Equal(0, RowPlacement.StackOffset(true, 100, 4, 32));
                Equal(0, RowPlacement.StackOffset(true, 0, 1, 32));
            });
            Test("RowTop adds the offset and subtracts the scroll", () =>
            {
                Equal(0 + 4 + 64, RowPlacement.RowTop(0, 4, 2, 32, 0));
                Equal(62 + 0 + 64 - 10, RowPlacement.RowTop(62, 0, 2, 32, 10));
            });
            Test("stacked: the last row ends where the area ends", () =>
            {
                int off = RowPlacement.StackOffset(true, 100, 3, 32);
                Equal(100, RowPlacement.RowTop(0, off, 2, 32, 0) + 32);
            });
            Test("IndexFromY inside the rows", () =>
            {
                Equal(0, RowPlacement.IndexFromY(62, 62, 300, 0, 0, 32, 5));
                Equal(1, RowPlacement.IndexFromY(94, 62, 300, 0, 0, 32, 5));
                Equal(2, RowPlacement.IndexFromY(94, 62, 300, 0, 32, 32, 5));
            });
            Test("IndexFromY: the space above the first stacked row is not a row", () =>
            {
                // 3 rows at the bottom of an area 0..100 (offset 4): y 0..3 is empty, y 4 is the first row
                Equal(-1, RowPlacement.IndexFromY(0, 0, 100, 4, 0, 32, 3));
                Equal(-1, RowPlacement.IndexFromY(3, 0, 100, 4, 0, 32, 3));
                Equal(0, RowPlacement.IndexFromY(4, 0, 100, 4, 0, 32, 3));
                Equal(2, RowPlacement.IndexFromY(99, 0, 100, 4, 0, 32, 3));
                Equal(-1, RowPlacement.IndexFromY(-1, 0, 100, 4, 0, 32, 3));
                Equal(-1, RowPlacement.IndexFromY(10, 0, 100, 40, 0, 32, 2)); // a gap of 40 above: y 10 is in it, not in row 0 (a division on -30 would give 0)
            });
            Test("IndexFromY: below the area, below the last row, no rows", () =>
            {
                Equal(-1, RowPlacement.IndexFromY(100, 0, 100, 4, 0, 32, 3));
                Equal(-1, RowPlacement.IndexFromY(150, 62, 300, 0, 0, 32, 2));
                Equal(-1, RowPlacement.IndexFromY(70, 62, 300, 0, 0, 32, 0));
                Equal(-1, RowPlacement.IndexFromY(70, 62, 300, 0, 0, 0, 3));
            });
            Test("MaxScroll", () =>
            {
                Equal(0, RowPlacement.MaxScroll(3, 32, 100));
                Equal(28, RowPlacement.MaxScroll(4, 32, 100));
                Equal(0, RowPlacement.MaxScroll(0, 32, 100));
            });
            Test("FromBottom and ScrollFromBottom go back and forth", () =>
            {
                // 10 rows, area 100: max scroll 220
                Equal(0, RowPlacement.FromBottom(220, 10, 32, 100));
                Equal(220, RowPlacement.FromBottom(0, 10, 32, 100));
                Equal(100, RowPlacement.FromBottom(120, 10, 32, 100));
                Equal(120, RowPlacement.ScrollFromBottom(100, 10, 32, 100));
                // one row more: the same distance from the bottom keeps the view of the bottom rows
                Equal(152, RowPlacement.ScrollFromBottom(100, 11, 32, 100));
                // clamps
                Equal(0, RowPlacement.ScrollFromBottom(9999, 10, 32, 100));
                Equal(220, RowPlacement.ScrollFromBottom(-5, 10, 32, 100));
                Equal(0, RowPlacement.ScrollFromBottom(0, 2, 32, 100));
            });
            Test("OldRowTop", () => Equal(10 + 3 * 32, RowPlacement.OldRowTop(10, 3, 32)));
            Test("WallpaperSpan hanging", () =>
            {
                int t, b;
                RowPlacement.WallpaperSpan(false, false, 62, 300, 62, 3, 32, 0, out t, out b);
                Equal(62 + 96, t); Equal(300, b);
                RowPlacement.WallpaperSpan(false, true, 62, 300, 62, 3, 32, 0, out t, out b);
                Equal(62, t); Equal(300, b);
                RowPlacement.WallpaperSpan(false, false, 62, 300, 62 - 500, 30, 32, 500, out t, out b);
                True(b <= t, "no wallpaper when the rows fill the area");
            });
            Test("WallpaperSpan stacked", () =>
            {
                int t, b;
                RowPlacement.WallpaperSpan(true, false, 0, 200, 104, 3, 32, 0, out t, out b);
                Equal(0, t); Equal(104, b);
                RowPlacement.WallpaperSpan(true, true, 0, 200, 104, 3, 32, 0, out t, out b);
                Equal(0, t); Equal(200, b);
                RowPlacement.WallpaperSpan(true, false, 0, 200, -50, 10, 32, 0, out t, out b);
                True(b <= t, "no wallpaper when the rows fill the area");
            });

            Console.WriteLine(_run + " tests, " + _failed + " failed");
            return _failed;
        }
    }
}
