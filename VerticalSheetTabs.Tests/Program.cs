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
            // ---- Wallpaper cropping
            Test("WallpaperCrop reads and writes the text", () =>
            {
                Equal("0.1,0.2,0.9,0.8", WallpaperCrop.Parse("0.1,0.2,0.9,0.8").Format());
                Equal("0.1235,0,1,1", WallpaperCrop.Parse("0.12346, 0, 1, 1").Format());
                Equal("0,0,1,1", WallpaperCrop.Parse("abc").Format());
                Equal("0,0,1,1", WallpaperCrop.Parse("0,0,1").Format());
                Equal("0,0,1,1", WallpaperCrop.Parse("").Format());
                Equal("0,0,1,1", WallpaperCrop.Parse(null).Format());
                Equal("0,0,1,1", WallpaperCrop.Parse("0,0,1,1,1").Format());
                var old = System.Threading.Thread.CurrentThread.CurrentCulture;
                try
                {
                    System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("de-DE");
                    Equal("0.25,0.5,0.75,1", new WallpaperCrop(0.25, 0.5, 0.75, 1).Format());
                    Equal("0.25,0.5,0.75,1", WallpaperCrop.Parse("0.25,0.5,0.75,1").Format());
                }
                finally { System.Threading.Thread.CurrentThread.CurrentCulture = old; }
            });
            Test("WallpaperCrop treats a value that is not a part as no cropping", () =>
            {
                True(default(WallpaperCrop).IsFull, "default");
                True(WallpaperCrop.Full.IsFull, "full");
                True(WallpaperCrop.Parse("-1,-1,2,2").IsFull, "out of the picture is fitted in");
                True(WallpaperCrop.Parse("0.5,0,0.51,1").IsFull, "too narrow");
                True(WallpaperCrop.Parse("0,0.5,1,0.505").IsFull, "too low");
                True(WallpaperCrop.Parse("0.5,0.5,0.4,0.9").IsFull, "turned inside out");
                True(new WallpaperCrop(double.NaN, 0, 1, 1).IsFull, "NaN");
                True(!WallpaperCrop.Parse("0.1,0,1,1").IsFull, "a real part");
                True(new WallpaperCrop(0.1, 0.1, 0.9, 0.9).SameAs(WallpaperCrop.Parse("0.1,0.1,0.9,0.9")), "same");
                True(!new WallpaperCrop(0.1, 0.1, 0.9, 0.9).SameAs(WallpaperCrop.Full), "different");
            });
            Test("WallpaperCrop pixels are rounded outwards and stay in the picture", () =>
            {
                Equal(new System.Drawing.Rectangle(0, 0, 100, 50), default(WallpaperCrop).ToPixels(100, 50));
                Equal(new System.Drawing.Rectangle(10, 5, 80, 40), new WallpaperCrop(0.1, 0.1, 0.9, 0.9).ToPixels(100, 50));
                Equal(new System.Drawing.Rectangle(33, 0, 34, 50), new WallpaperCrop(0.333, 0, 0.667, 1).ToPixels(100, 50));
                Equal(new System.Drawing.Rectangle(0, 0, 1, 1), new WallpaperCrop(0.1, 0.1, 0.9, 0.9).ToPixels(1, 1));
                System.Drawing.Rectangle r = new WallpaperCrop(0.98, 0.98, 1, 1).ToPixels(7, 7);
                True(r.Right <= 7 && r.Bottom <= 7 && r.Width >= 1 && r.Height >= 1, "inside");
            });
            Test("CropDrag finds what the pointer holds", () =>
            {
                var f = new System.Drawing.RectangleF(100, 100, 200, 100);
                Func<float, float, CropHandle> hit = (x, y) => CropDrag.HitTest(f, new System.Drawing.PointF(x, y), 8, true);
                Equal(CropHandle.NW, hit(100, 100)); Equal(CropHandle.NW, hit(104, 96));
                Equal(CropHandle.NE, hit(300, 100)); Equal(CropHandle.SW, hit(100, 200)); Equal(CropHandle.SE, hit(300, 200));
                Equal(CropHandle.N, hit(200, 100)); Equal(CropHandle.S, hit(200, 204)); Equal(CropHandle.W, hit(96, 150)); Equal(CropHandle.E, hit(300, 150));
                Equal(CropHandle.Move, hit(200, 150));
                Equal(CropHandle.None, hit(50, 50)); Equal(CropHandle.None, hit(200, 250)); Equal(CropHandle.None, hit(320, 150));
                Equal(CropHandle.None, CropDrag.HitTest(f, new System.Drawing.PointF(200, 150), 8, false));
                Equal(CropHandle.NW, CropDrag.HitTest(f, new System.Drawing.PointF(100, 100), 8, false));
                // a small frame: the grab is narrowed, so that the inside can be held
                var small = new System.Drawing.RectangleF(0, 0, 12, 12);
                Equal(CropHandle.Move, CropDrag.HitTest(small, new System.Drawing.PointF(6, 6), 8, true));
                Equal(CropHandle.NW, CropDrag.HitTest(small, new System.Drawing.PointF(1, 1), 8, true));
            });
            Test("CropDrag moves and sizes the frame inside the picture", () =>
            {
                var bounds = new System.Drawing.RectangleF(0, 0, 400, 300);
                var start = new System.Drawing.RectangleF(100, 100, 200, 100);
                Func<CropHandle, float, float, System.Drawing.RectangleF> drag = (h, x, y) => CropDrag.Drag(h, start, new System.Drawing.PointF(0, 0), new System.Drawing.PointF(x, y), bounds, 24, 24);
                Equal(new System.Drawing.RectangleF(50, 60, 250, 140), drag(CropHandle.NW, -50, -40));
                Equal(new System.Drawing.RectangleF(0, 0, 300, 200), drag(CropHandle.NW, -500, -500));
                Equal(new System.Drawing.RectangleF(100, 100, 300, 200), drag(CropHandle.SE, 500, 500));
                Equal(new System.Drawing.RectangleF(100, 100, 220, 100), drag(CropHandle.E, 20, 99));
                Equal(new System.Drawing.RectangleF(100, 80, 200, 120), drag(CropHandle.N, 99, -20));
                Equal(new System.Drawing.RectangleF(100, 100, 200, 120), drag(CropHandle.S, 99, 20));
                Equal(new System.Drawing.RectangleF(120, 100, 180, 100), drag(CropHandle.W, 20, 99));
                // it does not turn inside out: the least size stays
                Equal(new System.Drawing.RectangleF(276, 176, 24, 24), drag(CropHandle.NW, 900, 900));
                Equal(new System.Drawing.RectangleF(100, 100, 24, 24), drag(CropHandle.SE, -900, -900));
                // moving keeps the size and stays in the picture
                Equal(new System.Drawing.RectangleF(150, 130, 200, 100), drag(CropHandle.Move, 50, 30));
                Equal(new System.Drawing.RectangleF(200, 0, 200, 100), drag(CropHandle.Move, 500, -500));
                // a small picture: the least size is at most half of it
                System.Drawing.RectangleF tiny = CropDrag.Drag(CropHandle.SE, new System.Drawing.RectangleF(0, 0, 30, 30), new System.Drawing.PointF(0, 0), new System.Drawing.PointF(-50, -50), new System.Drawing.RectangleF(0, 0, 30, 30), 24, 24);
                Equal(new System.Drawing.RectangleF(0, 0, 15, 15), tiny);
            });
            Test("CropPreviewMath fits the picture and converts the frame", () =>
            {
                System.Drawing.RectangleF r = CropPreviewMath.FitRect(400, 300, 100, 200, 12);
                True(Math.Abs(r.Height - 276) < 0.01 && Math.Abs(r.Width - 138) < 0.01 && Math.Abs(r.X - 131) < 0.01 && Math.Abs(r.Y - 12) < 0.01, "tall picture");
                r = CropPreviewMath.FitRect(400, 300, 10, 10, 12);
                True(Math.Abs(r.Width - 276) < 0.01 && Math.Abs(r.Height - 276) < 0.01, "a small picture is enlarged");
                r = CropPreviewMath.FitRect(400, 300, 800, 100, 12);
                True(Math.Abs(r.Width - 376) < 0.01 && Math.Abs(r.Height - 47) < 0.01 && Math.Abs(r.X - 12) < 0.01, "wide picture");
                True(CropPreviewMath.FitRect(400, 300, 0, 10, 12).IsEmpty && CropPreviewMath.FitRect(20, 20, 10, 10, 12).IsEmpty, "nothing to show");
                var bounds = new System.Drawing.RectangleF(131, 12, 138, 276);
                var crop = new WallpaperCrop(0.25, 0.1, 0.75, 0.6);
                System.Drawing.RectangleF f = CropPreviewMath.ToFrame(crop, bounds);
                True(Math.Abs(f.Left - (131 + 34.5)) < 0.01 && Math.Abs(f.Top - (12 + 27.6)) < 0.01 && Math.Abs(f.Width - 69) < 0.01 && Math.Abs(f.Height - 138) < 0.01, "frame in pixels");
                True(CropPreviewMath.FromFrame(f, bounds).SameAs(crop), "and back");
                True(CropPreviewMath.FromFrame(bounds, bounds).IsFull, "whole picture");
            });
            Test("Wallpaper.Render draws only the part that is used", () =>
            {
                using (var src = new System.Drawing.Bitmap(4, 2))
                {
                    for (int y = 0; y < 2; y++)
                        for (int x = 0; x < 4; x++)
                            src.SetPixel(x, y, x < 2 ? System.Drawing.Color.FromArgb(255, 0, 0) : System.Drawing.Color.FromArgb(0, 0, 255));
                    var o = new WallpaperOptions { Fit = WallpaperFit.FitBoth, KeepRatio = false };
                    using (var whole = Wallpaper.Render(src, 8, 4, o))
                    {
                        True(whole.GetPixel(1, 2).R > 200 && whole.GetPixel(6, 2).B > 200, "no cropping: red on the left, blue on the right");
                    }
                    o.Crop = new WallpaperCrop(0.5, 0, 1, 1);
                    using (var right = Wallpaper.Render(src, 8, 4, o))
                    {
                        bool allBlue = true;
                        string bad = "";
                        for (int y = 0; y < 4; y++)
                            for (int x = 0; x < 8; x++)
                            {
                                System.Drawing.Color c = right.GetPixel(x, y);
                                if (c.R > 8 || c.B < 247) { allBlue = false; bad += " (" + x + "," + y + ")=" + c.R + "," + c.G + "," + c.B + "," + c.A; }
                            }
                        True(allBlue, "the right half only, with no red on the edges:" + bad);
                    }
                    // tiles and the original size use the part too
                    var t = new WallpaperOptions { Fit = WallpaperFit.Original, Tile = true, Crop = new WallpaperCrop(0.5, 0, 1, 1) };
                    using (var tiled = Wallpaper.Render(src, 6, 4, t))
                    {
                        bool allBlue = true;
                        for (int y = 0; y < 4; y++)
                            for (int x = 0; x < 6; x++) if (tiled.GetPixel(x, y).B < 247 || tiled.GetPixel(x, y).R > 8) allBlue = false;
                        True(allBlue, "tiles of the part");
                    }
                    // a whole-picture frame draws what no frame draws
                    var a = new WallpaperOptions { Fit = WallpaperFit.FitBoth, KeepRatio = true };
                    var b = new WallpaperOptions { Fit = WallpaperFit.FitBoth, KeepRatio = true, Crop = WallpaperCrop.Full };
                    using (var ra = Wallpaper.Render(src, 9, 7, a))
                    using (var rb = Wallpaper.Render(src, 9, 7, b))
                    {
                        bool same = true;
                        for (int y = 0; y < 7; y++)
                            for (int x = 0; x < 9; x++) if (ra.GetPixel(x, y) != rb.GetPixel(x, y)) same = false;
                        True(same, "the same pixels");
                    }
                    True(!new WallpaperOptions().SameAs(new WallpaperOptions { Crop = new WallpaperCrop(0.1, 0.1, 0.9, 0.9) }), "options differ by the part");
                }
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
