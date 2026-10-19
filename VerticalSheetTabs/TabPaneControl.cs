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

        public TabPaneControl()
        {
            _list.Dock = DockStyle.Fill;
            _list.IntegralHeight = false;
            _list.BorderStyle = BorderStyle.None;
            _list.SelectedIndexChanged += List_SelectedIndexChanged;
            Controls.Add(_list);

            BackColor = SystemColors.Window;
        }

        // First-open width: 16 half-width characters plus the list's side padding.
        public int DefaultContentWidth()
        {
            Size text = TextRenderer.MeasureText(new string('0', 16), _list.Font, new Size(int.MaxValue, int.MaxValue),
                TextFormatFlags.NoPadding);
            return text.Width + 2 * (int)Math.Round(4 * DeviceDpi / 96.0);
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

        private void List_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_updating) return;
            string name = _list.SelectedItem as string;
            if (name != null && SheetSelected != null) SheetSelected(name);
        }
    }
}
