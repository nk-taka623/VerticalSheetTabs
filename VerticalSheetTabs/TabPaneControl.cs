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
        public event Action<bool> GridFocusRequested;
        public event Action MoveCopyRequested;
        // (sheets, the sheet they are dropped in front of or null for the end, copy instead of move)
        public event Action<IList<string>, string, bool> MoveRequested;
        // A folder row was clicked (folder id): open or close it.
        public event Action<string> FolderToggleRequested;
        public event Action<bool> AllFoldersRequested;
        // An Excel command (idMso) for a key pressed in the pane, e.g. "FileSave" for Ctrl+S.
        public event Action<string> ExcelCommandRequested;
        public event Action<IList<string>, IList<string>, string, string, string, bool> FolderMoveRequested;
        public event Action<string, string> FolderRenameRequested;
        // (selected sheets) "new folder" from the sheet menu.
        public event Action<IList<string>> NewFolderRequested;
        // (selected folders, selected sheets) the Delete key.
        public event Action<IList<string>, IList<string>> DeleteKeyRequested;
        // (folders, also delete their sheets) from the folder menu.
        public event Action<IList<string>, bool> DeleteFoldersRequested;
        // (command, folders): "select", "deselect", "hide", "unhide" on the sheets of the folders.
        public event Action<string, IList<string>> FolderCommandRequested;
        // (folders, color) the color of the folder rows.
        public event Action<IList<string>, TabColorChoice> FolderColorRequested;
        // (folders, color) the tab color of all sheets of the folders.

        // The "unhide" item is only enabled while the workbook has hidden sheets.
        public bool HasHiddenSheets { get; set; }
        // With the workbook structure protected, Excel does not allow adding, deleting, renaming,
        // moving, copying, hiding or unhiding sheets, so those commands are disabled.
        public bool StructureProtected { get; set; }
        // Whether the active sheet is protected (the menu item then reads "unprotect").
        public bool ActiveSheetProtected { get; set; }
        // The workbook's ten theme colors for the tab color palette; null = the Office theme.
        public Func<Color[]> ThemeColorProvider { get; set; }

        // Opens Excel's "Colors" dialog for "More colors"; returns 0xBBGGRR or null.
        public Func<int?> MoreColorsProvider { get; set; }

        private readonly ContextMenuStrip _menu = new ContextMenuStrip();
        private readonly ContextMenuStrip _folderMenu = new ContextMenuStrip();

        public TabPaneControl()
        {
            _strip.Dock = DockStyle.Fill;
            _strip.SelectionRequested += (names, active) => { if (SelectionRequested != null) SelectionRequested(names, active); };
            _strip.GridFocusRequested += hadFocus => { if (GridFocusRequested != null) GridFocusRequested(hadFocus); };
            _strip.RenameRequested += (o, n) => { if (RenameRequested != null) RenameRequested(o, n); };
            _strip.MoveRequested += (names, before, copy) => { if (MoveRequested != null) MoveRequested(names, before, copy); };

            _strip.ExcelCommandRequested += command => { if (ExcelCommandRequested != null) ExcelCommandRequested(command); };
            _strip.AllFoldersRequested += open => { if (AllFoldersRequested != null) AllFoldersRequested(open); };
            _strip.FolderToggleRequested += id => { if (FolderToggleRequested != null) FolderToggleRequested(id); };
            _strip.FolderMoveRequested += (sheets, folders, into, bs, bf, copy) => { if (FolderMoveRequested != null) FolderMoveRequested(sheets, folders, into, bs, bf, copy); };
            _strip.FolderRenameRequested += (id, name) => { if (FolderRenameRequested != null) FolderRenameRequested(id, name); };
            _strip.DeleteKeyRequested += (sheets, folders) => { if (DeleteKeyRequested != null) DeleteKeyRequested(sheets, folders); };

            // The "+" (adds a sheet after the active one, like Excel's own) is drawn by the strip.
            _strip.AddRequested += () => { if (AddRequested != null) AddRequested(); };
            Controls.Add(_strip);

            // The first item is ours; the sheet items follow Excel's own tab menu order.
            var settings = new ToolStripMenuItem("タブ表示の設定");
            settings.Click += (s, e) => { if (SettingsRequested != null) SettingsRequested(); };
            _menu.Items.Add(settings);
            _menu.Items.Add(new ToolStripSeparator());
            var newFolder = new ToolStripMenuItem("新しいフォルダ(&F)");
            newFolder.Click += (s, e) => { if (NewFolderRequested != null) NewFolderRequested(_strip.SelectedInOrder()); };
            _menu.Items.Add(newFolder);
            _menu.Items.Add(new ToolStripSeparator());
            var insert =MenuItem("挿入(&I)...", () => InsertRequested);
            var delete = MenuItem("削除(&D)", () => DeleteRequested);
            var rename = new ToolStripMenuItem("名前の変更(&R)");
            rename.Click += (s, e) => _strip.BeginRenameActive();
            var moveCopy = MenuItem("移動またはコピー(&M)...", () => MoveCopyRequested);
            var viewCode = MenuItem("コードの表示(&V)", () => ViewCodeRequested);
            var protect = MenuItem("シートの保護(&P)...", () => ProtectSheetRequested);
            var tabColor = new ToolStripMenuItem("シート見出しの色(&T)");
            var palette = new TabColorPalette { MoreColorsProvider = RequestMoreColor };
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
                newFolder.Enabled = free;
            };

            // The menu of a folder row (more commands come with step 8-5).
            var folderRename = new ToolStripMenuItem("フォルダ名の変更(&R)");
            folderRename.Click += (s, e) => _strip.BeginRenameFolder(_strip.RightClickedFolder);
            var folderDelete = new ToolStripMenuItem("フォルダの削除(&D)");
            folderDelete.Click += (s, e) => { if (DeleteFoldersRequested != null) DeleteFoldersRequested(_strip.SelectedFolders(), false); };
            var folderDeleteAll = new ToolStripMenuItem("フォルダと中のシートを削除(&X)");
            folderDeleteAll.Click += (s, e) => { if (DeleteFoldersRequested != null) DeleteFoldersRequested(_strip.SelectedFolders(), true); };
            var fDeselect = FolderItem("フォルダ内のシートをすべて非選択(&N)", "deselect");
            var fSelect = FolderItem("フォルダ内のシートをすべて選択(&S)", "select");
            var fHide = FolderItem("フォルダ内のシートをすべて非表示(&H)", "hide");
            var fUnhide = FolderItem("フォルダ内のシートをすべて再表示(&U)", "unhide");
            var fColor = PaletteItem("フォルダの色(&C)", c => { if (FolderColorRequested != null) FolderColorRequested(_strip.SelectedFolders(), c); });
            var fSyncColor = FolderItem("中の全てを常にフォルダと同一色に設定(&F)", "synccolor");
            var fCloseAll = FolderItem("全てのフォルダを閉じる(&L)", "closeall");
            var fOpenAll = FolderItem("全てのフォルダを開く(&O)", "openall");
            var folderSettings = new ToolStripMenuItem("タブ表示の設定");
            folderSettings.Click += (s, e) => { if (SettingsRequested != null) SettingsRequested(); };
            _folderMenu.Items.Add(folderSettings);
            _folderMenu.Items.Add(new ToolStripSeparator());
            _folderMenu.Items.Add(folderRename);
            _folderMenu.Items.Add(new ToolStripSeparator());
            _folderMenu.Items.Add(fDeselect);
            _folderMenu.Items.Add(fSelect);
            _folderMenu.Items.Add(fHide);
            _folderMenu.Items.Add(fUnhide);
            _folderMenu.Items.Add(new ToolStripSeparator());
            _folderMenu.Items.Add(fColor);
            _folderMenu.Items.Add(fSyncColor);
            _folderMenu.Items.Add(new ToolStripSeparator());
            _folderMenu.Items.Add(folderDelete);
            _folderMenu.Items.Add(folderDeleteAll);
            _folderMenu.Items.Add(new ToolStripSeparator());
            _folderMenu.Items.Add(fCloseAll);
            _folderMenu.Items.Add(fOpenAll);
            _folderMenu.Opening += (s, e) =>
            {
                bool free = !StructureProtected;
                folderRename.Enabled = _strip.SelectedFolders().Count <= 1;
                fHide.Enabled = free;
                fUnhide.Enabled = free;
                folderDeleteAll.Enabled = free;
            };
            _strip.SheetMenu = _menu;
            _strip.FolderMenu = _folderMenu;
            _strip.ContextMenuStrip = _menu;
        }

        private ToolStripMenuItem FolderItem(string text, string command)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => { if (FolderCommandRequested != null) FolderCommandRequested(command, _strip.SelectedFolders()); };
            return item;
        }

        // "More colors" of a palette in one of the menus. The menus are closed first: an open menu takes the keys
        // (the Ctrl+Enter that enters a cell being edited) and the focus away from where they belong.
        private int? RequestMoreColor()
        {
            if (MoreColorsProvider == null) return null;
            _menu.Close();
            _folderMenu.Close();
            Application.DoEvents();
            return MoreColorsProvider();
        }

        // A menu item that opens the tab color palette (the same one as for sheets).
        private ToolStripMenuItem PaletteItem(string text, Action<TabColorChoice> chosen)
        {
            var item = new ToolStripMenuItem(text);
            var palette = new TabColorPalette { MoreColorsProvider = RequestMoreColor };
            var host = new ToolStripControlHost(palette) { Padding = Padding.Empty, Margin = Padding.Empty, AutoSize = false, Size = palette.Size };
            item.DropDown = new ToolStripDropDown { Padding = Padding.Empty };
            item.DropDown.Items.Add(host);
            palette.ColorChosen += choice =>
            {
                _folderMenu.Close();
                chosen(choice);
            };
            item.DropDownOpening += (s, e) =>
            {
                palette.SetTheme(ThemeColorProvider != null ? ThemeColorProvider() : null);
                host.Size = palette.Size;
            };
            return item;
        }

        // First-open width: 12 half-width characters plus the tab's paddings.
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

        public void BeginRenameFolder(string folderId)
        {
            _strip.BeginRenameFolder(folderId);
        }

        // Whether Excel is editing a cell now (set by PaneManager).
        public Func<bool> IsExcelEditing { set { _strip.IsExcelEditing = value; } }

        public bool OutlineVisible { get { return _strip.OutlineVisible; } set { _strip.OutlineVisible = value; } }

        public int OutlineLevel { get { return _strip.OutlineLevel; } set { _strip.OutlineLevel = value; } }

        public SelectStyle SelectionStyle { get { return _strip.SelectionStyle; } set { _strip.SelectionStyle = value; } }

        public ActiveMarker ActiveMarker { get { return _strip.ActiveMarker; } set { _strip.ActiveMarker = value; } }

        public void SetWobble(WobbleType type, int amp, int dur) { _strip.SetWobble(type, amp, dur); }

        public void SetSheets(IList<string> names, string activeName, ICollection<string> selected, IDictionary<string, int> colors, IList<StripRow> rows)
        {
            _strip.StructureProtected = StructureProtected;
            _strip.SetSheets(names, activeName, selected, colors, rows);
        }
    }
}
