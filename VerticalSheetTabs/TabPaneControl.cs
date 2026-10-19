using System;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // Step1 prototype: content of the task pane. Shows pane info and has
    // buttons to dock the pane left/right. Replaced by the real tab list in Step2.
    public class TabPaneControl : UserControl
    {
        private readonly Label _info = new Label();
        private readonly Button _left = new Button();
        private readonly Button _right = new Button();

        public event EventHandler LeftClicked;
        public event EventHandler RightClicked;

        public TabPaneControl()
        {
            _info.Dock = DockStyle.Top;
            _info.Height = 80;
            _info.Padding = new Padding(4);

            _left.Text = "Dock Left";
            _left.Dock = DockStyle.Top;
            _left.Click += (s, e) => { if (LeftClicked != null) LeftClicked(this, EventArgs.Empty); };

            _right.Text = "Dock Right";
            _right.Dock = DockStyle.Top;
            _right.Click += (s, e) => { if (RightClicked != null) RightClicked(this, EventArgs.Empty); };

            // Dock=Top controls are laid out bottom-up in z-order.
            Controls.Add(_right);
            Controls.Add(_left);
            Controls.Add(_info);

            BackColor = SystemColors.Window;
        }

        public void SetInfo(string windowCaption, string dock)
        {
            _info.Text = "Window: " + windowCaption + "\r\nDock: " + dock
                + "\r\nWidth: " + Width + " px";
        }
    }
}
