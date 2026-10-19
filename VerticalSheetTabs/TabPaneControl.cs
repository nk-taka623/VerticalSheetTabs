using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Step2 prototype: content of the task pane. A plain ListBox shows the sheet
    // names; the owner-drawn tab graphics replace it in Step6. The Dock buttons
    // are test aids from Step1 and go away when the settings dialog arrives (Step4).
    public class TabPaneControl : UserControl
    {
        private readonly ListBox _list = new ListBox();
        private readonly Label _info = new Label();
        private readonly Button _left = new Button();
        private readonly Button _right = new Button();

        private List<string> _names = new List<string>();
        private string _activeName;
        private bool _updating;

        public event EventHandler LeftClicked;
        public event EventHandler RightClicked;
        // Raised when the user clicks a sheet in the list.
        public event Action<string> SheetSelected;

        public TabPaneControl()
        {
            _list.Dock = DockStyle.Fill;
            _list.IntegralHeight = false;
            _list.BorderStyle = BorderStyle.None;
            _list.SelectedIndexChanged += List_SelectedIndexChanged;

            _info.Dock = DockStyle.Top;
            _info.Height = 48;
            _info.Padding = new Padding(4, 2, 4, 2);

            _left.Text = "Dock Left";
            _left.Dock = DockStyle.Top;
            _left.Click += (s, e) => { if (LeftClicked != null) LeftClicked(this, EventArgs.Empty); };

            _right.Text = "Dock Right";
            _right.Dock = DockStyle.Top;
            _right.Click += (s, e) => { if (RightClicked != null) RightClicked(this, EventArgs.Empty); };

            // The Fill control must be added first; Top controls then stack in reverse order.
            Controls.Add(_list);
            Controls.Add(_right);
            Controls.Add(_left);
            Controls.Add(_info);

            BackColor = SystemColors.Window;
        }

        // Test aid: shows the pane geometry while the width handling is verified.
        public void SetInfo(string dock, int paneWidth, int chrome)
        {
            _info.Text = dock + " pane=" + paneWidth + " ctl=" + Width + " chrome=" + chrome + " dpi=" + DeviceDpi;
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
