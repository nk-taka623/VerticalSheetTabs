using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Content of the task pane. A ListBox (owner-drawn, with its own selection handling) shows
    // the sheet names; the graphics that match Excel's tabs replace the drawing in Step6.
    // Selection follows Excel's tabs: click = one sheet, Ctrl+click = add/remove a sheet to the
    // group, Shift+click = range from the active sheet. The active sheet is drawn in bold.
    public class TabPaneControl : UserControl
    {
        private readonly ListBox _list = new ListBox();

        private List<string> _names = new List<string>();
        private HashSet<string> _selected = new HashSet<string>();
        private string _activeName;

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

        private readonly Dictionary<string, int> _colors = new Dictionary<string, int>();

        private readonly ContextMenuStrip _menu = new ContextMenuStrip();
        private readonly Button _add = new Button();
        private TextBox _editor;
        private string _editingName;

        // Drag and drop of tabs. The line shows where the dragged sheets will land.
        private readonly Panel _dropLine = new Panel { Height = 2, BackColor = SystemColors.Highlight, Visible = false };
        private Point _dragOrigin;
        private int _dragIndex = -1;
        private bool _dragging;
        private List<string> _dragNames = new List<string>();
        private int _dropIndex = -1;
        // Activating a sheet moves the focus into Excel, which would cancel a drag that has just
        // started. So a left click only reaches Excel when the button is released.
        private bool _selectionPending;

        public TabPaneControl()
        {
            _list.Dock = DockStyle.Fill;
            _list.IntegralHeight = false;
            _list.BorderStyle = BorderStyle.None;
            _list.SelectionMode = SelectionMode.None;
            _list.DrawMode = DrawMode.OwnerDrawFixed;
            _list.DrawItem += List_DrawItem;
            _list.MouseDown += List_MouseDown;
            _list.MouseDoubleClick += List_MouseDoubleClick;
            _list.MouseMove += List_MouseMove;
            _list.MouseUp += List_MouseUp;
            _list.KeyDown += List_KeyDown;
            _list.AllowDrop = true;
            _list.DragOver += List_DragOver;
            _list.DragLeave += (s, e) => HideDropLine();
            _list.DragDrop += List_DragDrop;
            _list.Controls.Add(_dropLine);

            // The "+" button adds a sheet after the active one, like Excel's own.
            _add.Text = "+";
            _add.Dock = DockStyle.Bottom;
            _add.FlatStyle = FlatStyle.Flat;
            _add.Click += (s, e) => { if (AddRequested != null) AddRequested(); };
            // The Fill control must be added first.
            Controls.Add(_list);
            Controls.Add(_add);

            // The first item is ours; the sheet items follow Excel's own tab menu order.
            var settings = new ToolStripMenuItem("タブ表示の設定");
            settings.Click += (s, e) => { if (SettingsRequested != null) SettingsRequested(); };
            _menu.Items.Add(settings);
            _menu.Items.Add(new ToolStripSeparator());
            var insert = MenuItem("挿入(&I)...", () => InsertRequested);
            var delete = MenuItem("削除(&D)", () => DeleteRequested);
            var rename = new ToolStripMenuItem("名前の変更(&R)");
            rename.Click += (s, e) => BeginRename(_names.IndexOf(_activeName));
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
                if (TabColorRequested != null) TabColorRequested(SelectedInOrder(), choice);
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
                rename.Enabled = free && _selected.Count <= 1;
                moveCopy.Enabled = free;
                hide.Enabled = free;
                unhide.Enabled = free && HasHiddenSheets;
                protect.Text = ActiveSheetProtected ? "シート保護の解除(&P)" : "シートの保護(&P)...";
                ungroup.Visible = _selected.Count > 1;
            };
            _list.ContextMenuStrip = _menu;

            BackColor = SystemColors.Window;
        }

        // First-open width: 16 half-width characters plus the list's side padding.
        public int DefaultContentWidth()
        {
            Size text = TextRenderer.MeasureText(new string('0', 16), _list.Font, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding);
            return text.Width + 2 * (int)Math.Round(4 * DeviceDpi / 96.0);
        }

        private static ToolStripMenuItem MenuItem(string text, Func<Action> handler)
        {
            var item = new ToolStripMenuItem(text);
            item.Click += (s, e) => { Action a = handler(); if (a != null) a(); };
            return item;
        }

        // Updates the list only when something actually changed, so the
        // polling timer does not cause flicker or lose the scroll position.
        public void SetSheets(IList<string> names, string activeName, ICollection<string> selected, IDictionary<string, int> colors)
        {
            _add.Enabled = !StructureProtected;
            bool sameColors = _colors.Count == colors.Count && _colors.All(kv => { int v; return colors.TryGetValue(kv.Key, out v) && v == kv.Value; });
            if (!sameColors)
            {
                _colors.Clear();
                foreach (var kv in colors) _colors[kv.Key] = kv.Value;
                _list.Invalidate();
            }
            bool sameNames = _names.SequenceEqual(names);
            // While a mouse button is down the user is changing the selection; do not fight them.
            bool userBusy = Control.MouseButtons != MouseButtons.None;
            bool sameSelection = userBusy
                || (_activeName == activeName && _selected.SetEquals(selected));
            if (sameNames && sameSelection) return;

            if (!sameNames)
            {
                EndRename(false);
                _names = new List<string>(names);
                _list.BeginUpdate();
                _list.Items.Clear();
                foreach (string n in _names) _list.Items.Add(n);
                _list.EndUpdate();
            }
            if (!userBusy)
            {
                _activeName = activeName;
                _selected = new HashSet<string>(selected);
            }
            _list.Invalidate();
        }

        private void List_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _names.Count) return;
            string name = _names[e.Index];
            bool selected = _selected.Contains(name);
            bool activeInGroup = selected && _selected.Count > 1 && name == _activeName;

            Color back = selected ? SystemColors.Highlight : SystemColors.Window;
            Color fore = selected ? SystemColors.HighlightText : SystemColors.WindowText;
            using (var brush = new SolidBrush(back)) e.Graphics.FillRectangle(brush, e.Bounds);
            // A sheet with a tab color gets a bar of that color along the left edge (graphics: Step6).
            int textLeft = 2;
            int bgr;
            if (_colors.TryGetValue(name, out bgr))
            {
                int bar = Math.Max(4, e.Bounds.Height / 3);
                using (var brush = new SolidBrush(Color.FromArgb(bgr & 0xFF, (bgr >> 8) & 0xFF, (bgr >> 16) & 0xFF)))
                {
                    e.Graphics.FillRectangle(brush, e.Bounds.X, e.Bounds.Y, bar, e.Bounds.Height);
                }
                textLeft = bar + 3;
            }
            Font font = activeInGroup ? new Font(e.Font, FontStyle.Bold) : e.Font;
            try
            {
                var textBounds = new Rectangle(e.Bounds.X + textLeft, e.Bounds.Y, e.Bounds.Width - textLeft - 2, e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, name, font, textBounds, fore,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            finally
            {
                if (activeInGroup) font.Dispose();
            }
        }

        private List<string> SelectedInOrder()
        {
            return _names.Where(n => _selected.Contains(n)).ToList();
        }

        private void RaiseSelection()
        {
            if (_activeName == null || SelectionRequested == null) return;
            SelectionRequested(SelectedInOrder(), _activeName);
        }

        // Left button down: update the selection on screen, remember where a drag could start.
        // Excel is told when the button is released, unless a drag starts first.
        private void List_MouseDown(object sender, MouseEventArgs e)
        {
            int index = _list.IndexFromPoint(e.Location);
            if (index < 0 || index >= _names.Count) return;
            string name = _names[index];

            if (e.Button == MouseButtons.Left)
            {
                _dragIndex = index;
                _dragOrigin = e.Location;

                if ((ModifierKeys & Keys.Control) != 0)
                {
                    if (_selected.Contains(name) && _selected.Count > 1)
                    {
                        _selected.Remove(name);
                        if (_activeName == name) _activeName = SelectedInOrder().First();
                    }
                    else
                    {
                        _selected.Add(name);
                        _activeName = name;
                    }
                }
                else if ((ModifierKeys & Keys.Shift) != 0 && _names.IndexOf(_activeName) >= 0)
                {
                    int a = _names.IndexOf(_activeName);
                    int lo = Math.Min(a, index), hi = Math.Max(a, index);
                    _selected = new HashSet<string>(_names.Skip(lo).Take(hi - lo + 1));
                    // the anchor stays the active sheet
                }
                else if (_selected.Contains(name) && _selected.Count > 1 && _selected.Count < _names.Count)
                {
                    _activeName = name; // inside a group: switch the active sheet, keep the group
                }
                else
                {
                    _selected = new HashSet<string> { name };
                    _activeName = name;
                }
                _selectionPending = true;
                _list.Invalidate();
                return;
            }

            if (e.Button != MouseButtons.Right) return;
            // Like Excel, a right click on a sheet outside the group ungroups and activates it;
            // inside the group it only makes that sheet the active one. Excel needs to know now,
            // because the menu commands act on the active sheet(s).
            if (!_selected.Contains(name)) _selected = new HashSet<string> { name };
            _activeName = name;
            _list.Invalidate();
            RaiseSelection();
        }

        private void List_MouseUp(object sender, MouseEventArgs e)
        {
            _dragIndex = -1;
            if (e.Button != MouseButtons.Left || !_selectionPending) return;
            _selectionPending = false;
            RaiseSelection();
        }

        // Up/Down arrow keys move the active sheet (a single selection).
        private void List_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode != Keys.Up && e.KeyCode != Keys.Down) return;
            int i = _names.IndexOf(_activeName) + (e.KeyCode == Keys.Down ? 1 : -1);
            if (i < 0 || i >= _names.Count) return;
            e.Handled = true;
            _selected = new HashSet<string> { _names[i] };
            _activeName = _names[i];
            _list.Invalidate();
            RaiseSelection();
        }

        private void List_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left || _dragIndex < 0 || _editor != null || StructureProtected) return;
            Size drag = SystemInformation.DragSize;
            if (Math.Abs(e.X - _dragOrigin.X) < drag.Width && Math.Abs(e.Y - _dragOrigin.Y) < drag.Height) return;

            int index = _dragIndex;
            _dragIndex = -1;
            _selectionPending = false; // a drag does not change the selection in Excel by itself
            if (index >= _names.Count) return;

            // Dragging a sheet of the group takes the whole group along.
            string name = _names[index];
            _dragNames = _selected.Contains(name) ? SelectedInOrder() : new List<string> { name };
            _dragging = true;
            try
            {
                _list.DoDragDrop(name, DragDropEffects.Move | DragDropEffects.Copy);
            }
            finally
            {
                _dragging = false;
                HideDropLine();
            }
        }

        // Ctrl held: copy, otherwise move (the same keys as Excel's own tabs).
        private static bool IsCopy(DragEventArgs e)
        {
            return (e.KeyState & 8) != 0;
        }

        private void List_DragOver(object sender, DragEventArgs e)
        {
            if (!_dragging || !e.Data.GetDataPresent(typeof(string)))
            {
                e.Effect = DragDropEffects.None;
                return;
            }
            e.Effect = IsCopy(e) ? DragDropEffects.Copy : DragDropEffects.Move;

            Point p = _list.PointToClient(new Point(e.X, e.Y));
            ScrollNearEdge(p);
            _dropIndex = InsertIndexAt(p);
            ShowDropLine(_dropIndex);
        }

        private void List_DragDrop(object sender, DragEventArgs e)
        {
            int target = _dropIndex;
            HideDropLine();
            if (!_dragging || target < 0 || _dragNames.Count == 0) return;
            bool copy = IsCopy(e);

            // Dropping the sheets where they already are changes nothing (when they sit together).
            List<int> at = _dragNames.Select(n => _names.IndexOf(n)).OrderBy(i => i).ToList();
            if (at[0] < 0) return;
            bool together = at[at.Count - 1] - at[0] + 1 == at.Count;
            if (!copy && together && target >= at[0] && target <= at[at.Count - 1] + 1) return;

            // The sheets cannot be dropped in front of one of themselves: take the next other one.
            while (target < _names.Count && _dragNames.Contains(_names[target])) target++;
            if (MoveRequested != null) MoveRequested(_dragNames, target < _names.Count ? _names[target] : null, copy);
        }

        // The position (0..count) whose boundary is nearest: upper half of a row = before it.
        private int InsertIndexAt(Point p)
        {
            int index = _list.IndexFromPoint(p);
            if (index < 0) return _names.Count;
            Rectangle r = _list.GetItemRectangle(index);
            return p.Y < r.Top + r.Height / 2 ? index : index + 1;
        }

        private void ShowDropLine(int index)
        {
            int y;
            if (_names.Count == 0) { HideDropLine(); return; }
            if (index < _names.Count) y = _list.GetItemRectangle(index).Top;
            else y = _list.GetItemRectangle(_names.Count - 1).Bottom;
            _dropLine.Bounds = new Rectangle(0, Math.Max(0, y - 1), _list.ClientSize.Width, 2);
            _dropLine.Visible = true;
            _dropLine.BringToFront();
        }

        private void HideDropLine()
        {
            _dropLine.Visible = false;
            _dropIndex = -1;
        }

        // DoDragDrop does not scroll the list, so do it near the top and bottom edges.
        private void ScrollNearEdge(Point p)
        {
            int edge = Math.Max(8, _list.ItemHeight);
            if (p.Y < edge && _list.TopIndex > 0) _list.TopIndex--;
            else if (p.Y > _list.ClientSize.Height - edge && _list.TopIndex < _list.Items.Count - 1) _list.TopIndex++;
        }

        private void List_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            int index = _list.IndexFromPoint(e.Location);
            if (index >= 0) BeginRename(index);
        }

        // In-place name editing: Enter or leaving the box commits, Esc cancels. The length limit
        // is Excel's 31 characters; the other rules are checked by PaneManager when it renames.
        private void BeginRename(int index)
        {
            if (index < 0 || index >= _names.Count || StructureProtected) return;
            EndRename(false);
            Rectangle r = _list.GetItemRectangle(index);
            _editingName = _names[index];
            _editor = new TextBox
            {
                Text = _editingName,
                MaxLength = 31,
                BorderStyle = BorderStyle.FixedSingle,
                Bounds = new Rectangle(r.X, r.Y, r.Width, Math.Max(r.Height, SystemInformation.MenuHeight))
            };
            _editor.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; EndRename(true); }
                else if (e.KeyCode == Keys.Escape) { e.SuppressKeyPress = true; EndRename(false); }
            };
            _editor.LostFocus += (s, e) => EndRename(true);
            _list.Controls.Add(_editor);
            _editor.Focus();
            _editor.SelectAll();
        }

        private void EndRename(bool commit)
        {
            TextBox editor = _editor;
            if (editor == null) return;
            _editor = null; // first, so the LostFocus raised by removing the box does nothing
            string oldName = _editingName;
            string newName = editor.Text;
            _list.Controls.Remove(editor);
            editor.Dispose();
            if (commit && newName != oldName && RenameRequested != null) RenameRequested(oldName, newName);
        }
    }
}
