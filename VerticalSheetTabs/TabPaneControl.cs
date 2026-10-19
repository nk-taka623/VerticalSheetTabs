using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Step4 prototype: content of the task pane. A plain ListBox shows the sheet
    // names; the owner-drawn tab graphics replace it in Step6.
    public class TabPaneControl : UserControl
    {
        private readonly ListBox _list = new ListBox();

        private List<string> _names = new List<string>();
        private string _activeName;
        private bool _updating;

        // Raised when the user clicks a sheet in the list.
        public event Action<string> SheetSelected;
        // Raised by the "Tab display settings" item at the top of the context menu.
        public event Action SettingsRequested;
        // Sheet commands: the pane only reports the request, PaneManager runs it in Excel.
        public event Action AddRequested;
        public event Action InsertRequested;
        public event Action DeleteRequested;
        public event Action HideRequested;
        public event Action UnhideRequested;
        // (old name, new name) after the user finishes in-place editing.
        public event Action<string, string> RenameRequested;

        // The "unhide" item is only enabled while the workbook has hidden sheets.
        public bool HasHiddenSheets { get; set; }

        private readonly ContextMenuStrip _menu = new ContextMenuStrip();
        private readonly Button _add = new Button();
        private TextBox _editor;
        private string _editingName;

        public TabPaneControl()
        {
            _list.Dock = DockStyle.Fill;
            _list.IntegralHeight = false;
            _list.BorderStyle = BorderStyle.None;
            _list.SelectedIndexChanged += List_SelectedIndexChanged;
            _list.MouseDown += List_MouseDown;
            _list.MouseDoubleClick += List_MouseDoubleClick;

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
            _menu.Items.Add(MenuItem("挿入(&I)...", () => InsertRequested));
            _menu.Items.Add(MenuItem("削除(&D)", () => DeleteRequested));
            var rename = new ToolStripMenuItem("名前の変更(&R)");
            rename.Click += (s, e) => BeginRename(_list.SelectedIndex);
            _menu.Items.Add(rename);
            _menu.Items.Add(new ToolStripSeparator());
            _menu.Items.Add(MenuItem("非表示(&H)", () => HideRequested));
            var unhide = MenuItem("再表示(&U)...", () => UnhideRequested);
            _menu.Items.Add(unhide);
            _menu.Opening += (s, e) => unhide.Enabled = HasHiddenSheets;
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
        public void SetSheets(IList<string> names, string activeName)
        {
            bool sameNames = _names.SequenceEqual(names);
            if (sameNames && _activeName == activeName) return;

            _updating = true;
            try
            {
                if (!sameNames)
                {
                    EndRename(false);
                    _names = new List<string>(names);
                    _list.BeginUpdate();
                    _list.Items.Clear();
                    foreach (string n in _names) _list.Items.Add(n);
                    _list.EndUpdate();
                }
                _activeName = activeName;
                _list.SelectedIndex = activeName == null ? -1 : _names.IndexOf(activeName);
            }
            finally
            {
                _updating = false;
            }
        }

        // Like Excel's own tabs, a right click first makes the sheet under the pointer the active one.
        private void List_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            int index = _list.IndexFromPoint(e.Location);
            if (index >= 0 && index != _list.SelectedIndex) _list.SelectedIndex = index;
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
            if (index < 0 || index >= _names.Count) return;
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

        private void List_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            string name = _list.SelectedItem as string;
            if (name != null && SheetSelected != null) SheetSelected(name);
        }
    }
}
