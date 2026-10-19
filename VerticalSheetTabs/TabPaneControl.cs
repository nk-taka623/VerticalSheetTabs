using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Content of the task pane: the tab strip, the "+" button and the context menu. The tabs
    // themselves (drawing, selection, drag and drop, renaming) are in TabStrip.
    public class TabPaneControl : UserControl
    {
        private readonly TabStrip _strip = new TabStrip();

        // (selected sheets in tab order, the active one) after the user changed the selection.
        public event Action<IList<string>, string> SelectionRequested;
        // Raised by the "Tab display settings" item at the top of the context menu.
        public event Action SettingsRequested;
        // Sheet commands: the pane only reports the request, PaneManager runs it in Excel.
        public event Action AddRequested;
        public event Action InsertRequested;
        public event Action DeleteRequested;
        public event Action HideRequested;
        public event Action UnhideRequested;
        public event Action SelectAllRequested;
        public event Action UngroupRequested;
        public event Action ViewCodeRequested;
        public event Action ProtectSheetRequested;
        // (selected sheets, the chosen tab color): Excel applies a color to every sheet of a group.
        public event Action<IList<string>, TabColorChoice> TabColorRequested;
        // (old name, new name) after the user finishes in-place editing.
        public event Action<string, string> RenameRequested;
        public event Action MoveCopyRequested;
        // (sheets, the sheet they are dropped in front of or null for the end, copy instead of move)
        public event Action<IList<string>, string, bool> MoveRequested;

        // The "unhide" item is only enabled while the workbook has hidden sheets.
        public bool HasHiddenSheets { get; set; }
        // With the workbook structure protected, Excel does not allow adding, deleting, renaming,
        // moving, copying, hiding or unhiding sheets, so those commands are disabled.
        public bool StructureProtected { get; set; }
        // Whether the active sheet is protected (the menu item then reads "unprotect").
        public bool ActiveSheetProtected { get; set; }
        // The workbook's ten theme colors for the tab color palette; null = the Office theme.
        public Func<Color[]> ThemeColorProvider { get; set; }

        private readonly ContextMenuStrip _menu = new ContextMenuStrip();

        public TabPaneControl()
        {
            _strip.Dock = DockStyle.Fill;
            _strip.SelectionRequested += (names, active) => { if (SelectionRequested != null) SelectionRequested(names, active); };
            _strip.RenameRequested += (o, n) => { if (RenameRequested != null) RenameRequested(o, n); };
            _strip.MoveRequested += (names, before, copy) => { if (MoveRequested != null) MoveRequested(names, before, copy); };

            // The "+" (adds a sheet after the active one, like Excel's own) is drawn by the strip.
            _strip.AddRequested += () => { if (AddRequested != null) AddRequested(); };
            Controls.Add(_strip);

            // The first item is ours; the sheet items follow Excel's own tab menu order.
            var settings = new ToolStripMenuItem("タブ表示の設定");
            settings.Click += (s, e) => { if (SettingsRequested != null) SettingsRequested(); };
            _menu.Items.Add(settings);
            _menu.Items.Add(new ToolStripSeparator());
            var insert = MenuItem("挿入(&I)...", () => InsertRequested);
            var delete = MenuItem("削除(&D)", () => DeleteRequested);
            var rename = new ToolStripMenuItem("名前の変更(&R)");
            rename.Click += (s, e) => _strip.BeginRenameActive();
            var moveCopy = MenuItem("移動またはコピー(&M)...", () => MoveCopyRequested);
            var viewCode = MenuItem("コードの表示(&V)", () => ViewCodeRequested);
            var protect = MenuItem("シートの保護(&P)...", () => ProtectSheetRequested);
            var tabColor = new ToolStripMenuItem("シート見出しの色(&T)");
            var palette = new TabColorPalette();
            var paletteHost = new ToolStripControlHost(palette) { Padding = Padding.Empty, Margin = Padding.Empty, AutoSize = false, Size = palette.Size };
            tabColor.DropDown = new ToolStripDropDown { Padding = Padding.Empty };
            tabColor.DropDown.Items.Add(paletteHost);
            palette.ColorChosen += choice =>
            {
                _menu.Close();
                if (TabColorRequested != null) TabColorRequested(_strip.SelectedInOrder(), choice);
            };
            tabColor.DropDownOpening += (s, e) =>
            {
                palette.SetTheme(ThemeColorProvider != null ? ThemeColorProvider() : null);
                paletteHost.Size = palette.Size;
            };
            var hide = MenuItem("非表示(&H)", () => HideRequested);
            var unhide = MenuItem("再表示(&U)...", () => UnhideRequested);
            var selectAll = MenuItem("すべてのシートを選択(&S)", () => SelectAllRequested);
            var ungroup = MenuItem("シートのグループ解除(&N)", () => UngroupRequested);

            _menu.Items.Add(insert);
            _menu.Items.Add(delete);
            _menu.Items.Add(rename);
            _menu.Items.Add(moveCopy);
            _menu.Items.Add(viewCode);
            _menu.Items.Add(protect);
            _menu.Items.Add(tabColor);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(hide);
            _menu.Items.Add(unhide);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(selectAll);
            _menu.Items.Add(ungroup);
            _menu.Opening += (s, e) =>
            {
                bool free = !StructureProtected;
                insert.Enabled = free;
                delete.Enabled = free;
                rename.Enabled = free && _strip.SelectedCount <= 1;
                moveCopy.Enabled = free;
                hide.Enabled = free;
                unhide.Enabled = free && HasHiddenSheets;
                protect.Text = ActiveSheetProtected ? "シート保護の解除(&P)" : "シートの保護(&P)...";
                ungroup.Visible = _strip.SelectedCount > 1;
            };
            _strip.ContextMenuStrip = _menu;
        }

        // First-open width: 8 half-width characters plus the tab's paddings.
        public int DefaultContentWidth()
        {
            return _strip.DefaultContentWidth();
        }

        private static ToolStripMenuItem MenuItem(string text, Func<Action> handler)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => { Action a = handler(); if (a != null) a(); };
            return item;
        }

        public void SetSheets(IList<string> names, string activeName, ICollection<string> selected, IDictionary<string, int> colors)
        {
            _strip.StructureProtected = StructureProtected;
            _strip.SetSheets(names, activeName, selected, colors);
        }
    }
}
