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

            Console.WriteLine(_run + " tests, " + _failed + " failed");
            return _failed;
        }
    }
}
