using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // The groups of the "背景" tab of the settings dialog: the wallpaper file, how it is fitted, where it lies, and how much the
    // folder and sheet rows let it show through. Settings that do not apply to the chosen way of fitting are disabled (their values are kept).
    public class BackgroundPanel
    {
        public readonly GroupBox WallpaperGroup;
        public readonly GroupBox FitGroup;
        public readonly GroupBox PositionGroup;
        public readonly GroupBox TransparencyGroup;

        private string _path = "";
        private readonly TextBox _pathBox = new TextBox { ReadOnly = true };
        private readonly Label _status = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 6, 3, 0) };
        private readonly RadioButton[] _fit = new RadioButton[4];
        private readonly RadioButton _stretch = new RadioButton { Text = "比率を変えて引き伸ばす" };
        private readonly RadioButton _keep = new RadioButton { Text = "比率を保って全体を覆う" };
        private readonly Label _ratioLabel = new Label { Text = "縦横に合わせるとき", AutoSize = true };
        private readonly CheckBox _tile = new CheckBox { Text = "並べて表示する (タイル)" };
        private readonly Label _tileNote = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 2, 3, 0) };
        private readonly RadioButton[] _h = new RadioButton[3];
        private readonly RadioButton[] _v = new RadioButton[3];
        private readonly Label _hLabel = new Label { Text = "左右位置", AutoSize = true };
        private readonly Label _vLabel = new Label { Text = "上下位置", AutoSize = true };
        private readonly Label _posNote = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 6, 3, 0) };
        private readonly TrackBar _transparency = new TrackBar();
        private readonly Label _transparencyText = new Label { AutoSize = true };
        private readonly Label _transparencyNote = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 2, 3, 0) };
        private readonly CheckBox _outline = new CheckBox { Text = "文字に縁を付ける" };
        private readonly Label _outlineNote = new Label { AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(3, 2, 3, 0) };

        public readonly GroupBox PreviewGroup;
        private readonly CropPreview _preview = new CropPreview { Dock = DockStyle.Fill, Margin = new Padding(3, 4, 3, 3) };
        private readonly Button _resetCrop = new Button { Text = "リセット", AutoSize = true, MinimumSize = new Size(88, 0), Anchor = AnchorStyles.Right };
        private readonly CheckBox _limitCrop = new CheckBox { Text = "トリミング枠の範囲制限", AutoSize = true, Checked = true, Anchor = AnchorStyles.Right };
        private readonly CheckBox _flip = new CheckBox { Text = "表示画像の左右反転", AutoSize = true, Checked = false, Anchor = AnchorStyles.Right };
        private WallpaperCrop _cropValue = WallpaperCrop.Full;   // the frame that OK would keep
        private WallpaperRotation _rotationValue;                 // how the picture is turned, that OK would keep
        private bool _previewLoaded;                              // the picture is read when the preview is first shown (or a file is chosen)

        public BackgroundPanel(WallpaperOptions options, int rowTransparency, bool textOutline)
        {
            options = options ?? new WallpaperOptions();
            _path = options.Path ?? "";
            _cropValue = options.Crop.Normalize();
            _rotationValue = options.Rotation.Normalize();
            _limitCrop.Checked = options.LimitCrop;
            _flip.Checked = options.Flip;
            Font font = SystemFonts.MessageBoxFont;

            // 壁紙
            _pathBox.Text = _path;
            var browse = new Button { Text = "参照...", AutoSize = true, MinimumSize = new Size(88, 0) };
            var clear = new Button { Text = "解除", AutoSize = true, MinimumSize = new Size(88, 0) };
            browse.Click += (s, e) => Browse();
            clear.Click += (s, e) =>
            {
                _path = "";
                _pathBox.Text = "";
                _limitCrop.Checked = true;   // (first: see Browse)
                _flip.Checked = false;
                _cropValue = WallpaperCrop.Full;
                _rotationValue = WallpaperRotation.None;
                ShowPreview();
                Refresh();
            };
            var table = new TableLayoutPanel { ColumnCount = 3, RowCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _pathBox.Dock = DockStyle.Fill;
            _pathBox.Margin = new Padding(3, 4, 3, 3);
            _pathBox.MinimumSize = new Size(260, 0);
            table.Controls.Add(_pathBox, 0, 0);
            table.SetColumnSpan(_pathBox, 3);
            table.Controls.Add(_status, 0, 1);
            table.Controls.Add(browse, 1, 1);
            table.Controls.Add(clear, 2, 1);
            _status.Anchor = AnchorStyles.Left;
            browse.Margin = clear.Margin = new Padding(3, 1, 3, 0);   // (the group is as high as before the buttons moved below the path box)
            WallpaperGroup = Group("壁紙", font, table, true);

            // 表示方法
            string[] fitNames = { "元のサイズ", "横幅に合わせる", "縦幅に合わせる", "縦横に合わせる" };
            for (int i = 0; i < 4; i++) _fit[i] = new RadioButton { Text = fitNames[i], Checked = (int)options.Fit == i };
            _keep.Checked = options.KeepRatio;
            _stretch.Checked = !options.KeepRatio;
            _tile.Checked = options.Tile;
            var fitPanel = Stack();
            fitPanel.Controls.Add(Row(font, true, _fit));
            var ratioRow = Row(font, false, _stretch, _keep);
            ratioRow.Controls.Add(_ratioLabel);
            ratioRow.Controls.SetChildIndex(_ratioLabel, 0);
            _ratioLabel.Margin = new Padding(3, 6, 18, 3);
            ratioRow.Margin = new Padding(0, 6, 0, 0);
            fitPanel.Controls.Add(ratioRow);
            _tile.AutoSize = true;
            _tile.Margin = new Padding(3, 10, 3, 0);
            fitPanel.Controls.Add(_tile);
            fitPanel.Controls.Add(_tileNote);
            FitGroup = Group("表示方法", font, fitPanel, false);

            // 表示位置
            string[] hNames = { "左", "中央", "右" }, vNames = { "上", "中央", "下" };
            for (int i = 0; i < 3; i++)
            {
                _h[i] = new RadioButton { Text = hNames[i], Checked = (int)options.HAlign == i };
                _v[i] = new RadioButton { Text = vNames[i], Checked = (int)options.VAlign == i };
            }
            var hRow = Row(font, false, _h);
            var vRow = Row(font, false, _v);
            int labelWidth = Math.Max(TextRenderer.MeasureText(_hLabel.Text, font).Width, TextRenderer.MeasureText(_vLabel.Text, font).Width) + 16;
            _hLabel.AutoSize = false;
            _vLabel.AutoSize = false;
            _hLabel.Size = new Size(labelWidth, 22);
            _vLabel.Size = new Size(labelWidth, 22);
            _hLabel.TextAlign = ContentAlignment.MiddleLeft;
            _vLabel.TextAlign = ContentAlignment.MiddleLeft;
            _hLabel.Margin = _vLabel.Margin = new Padding(3, 3, 3, 3);
            hRow.Controls.Add(_hLabel);
            hRow.Controls.SetChildIndex(_hLabel, 0);
            vRow.Controls.Add(_vLabel);
            vRow.Controls.SetChildIndex(_vLabel, 0);
            var posPanel = Stack();
            posPanel.Controls.Add(hRow);
            posPanel.Controls.Add(vRow);
            posPanel.Controls.Add(_posNote);
            PositionGroup = Group("表示位置", font, posPanel, false);

            // 透過率
            _transparency.Minimum = 0;
            _transparency.Maximum = WorkbookSettings.TransparencyMax / WorkbookSettings.TransparencyStep;
            _transparency.TickFrequency = 1;
            _transparency.SmallChange = 1;
            _transparency.LargeChange = 2;
            _transparency.AutoSize = false;
            _transparency.Size = new Size(320, 36);
            _transparency.Value = WorkbookSettings.ClampTransparency(rowTransparency) / WorkbookSettings.TransparencyStep;
            _transparencyText.Text = "100 %"; // (for measuring)
            var sliderRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            var minText = new Label { Text = "0 %", AutoSize = true, Margin = new Padding(3, 10, 0, 0) };
            var maxText = new Label { Text = "90 %", AutoSize = true, Margin = new Padding(0, 10, 12, 0) };
            _transparencyText.Margin = new Padding(0, 10, 0, 0);
            sliderRow.Controls.Add(minText);
            sliderRow.Controls.Add(_transparency);
            sliderRow.Controls.Add(maxText);
            sliderRow.Controls.Add(_transparencyText);
            _outline.Checked = textOutline;
            _outline.AutoSize = true;
            _outline.Margin = new Padding(3, 10, 3, 0);
            var transPanel = Stack();
            transPanel.Controls.Add(sliderRow);
            transPanel.Controls.Add(_transparencyNote);
            transPanel.Controls.Add(_outline);
            transPanel.Controls.Add(_outlineNote);
            TransparencyGroup = Group("フォルダ・シートの透過率", font, transPanel, false);

            // プレビュー画面
            var caption = new Label { Text = "壁紙として表示したい範囲を指定して下さい。", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(3, 0, 3, 0) };
            _resetCrop.Margin = new Padding(3, 0, 3, 0);
            _limitCrop.Margin = new Padding(3, 0, 3, 0);
            _flip.Margin = new Padding(3, 0, 3, 0);
            var previewTable = new TableLayoutPanel { ColumnCount = 4, RowCount = 2, Dock = DockStyle.Fill, Margin = Padding.Empty };
            previewTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            previewTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            previewTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            previewTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            previewTable.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            previewTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            previewTable.Controls.Add(caption, 0, 0);
            previewTable.Controls.Add(_limitCrop, 1, 0);
            previewTable.Controls.Add(_flip, 2, 0);
            previewTable.Controls.Add(_resetCrop, 3, 0);
            previewTable.Controls.Add(_preview, 0, 1);
            previewTable.SetColumnSpan(_preview, 4);
            PreviewGroup = new GroupBox { Text = "プレビュー画面", Font = font, Padding = new Padding(8, 4, 8, 6) };
            PreviewGroup.Controls.Add(previewTable);
            _resetCrop.Click += (s, e) => { _limitCrop.Checked = true; _flip.Checked = false; _preview.ResetCrop(); };
            _flip.CheckedChanged += (s, e) => { _preview.FlipImage = _flip.Checked; };
            _limitCrop.CheckedChanged += (s, e) => { _preview.LimitCrop = _limitCrop.Checked; };
            _preview.CropChanged += (s, e) => { _cropValue = _preview.Crop; _rotationValue = _preview.Rotation; };
            _preview.VisibleChanged += (s, e) => EnsurePreview();
            ShowPreviewState(false, "");

            foreach (RadioButton r in _fit) r.CheckedChanged += (s, e) => Refresh();
            _keep.CheckedChanged += (s, e) => Refresh();
            _tile.CheckedChanged += (s, e) => Refresh();
            foreach (RadioButton r in _h) r.CheckedChanged += (s, e) => Refresh();
            foreach (RadioButton r in _v) r.CheckedChanged += (s, e) => Refresh();
            _transparency.ValueChanged += (s, e) => Refresh();
            Refresh();
        }

        // The body of the tab (the whole row is width pixels wide): on the left 壁紙, 表示方法, 表示位置 and 透過率 one below the other,
        // on the right the preview, as high as the four together.
        public Control BodyRow(int width)
        {
            int leftW = Math.Max(WallpaperGroup.Width, Math.Max(FitGroup.Width, Math.Max(PositionGroup.Width, TransparencyGroup.Width)));
            leftW = Math.Min(leftW, (width - 12) * 6 / 10);
            var left = Stack();
            GroupBox[] groups = { WallpaperGroup, FitGroup, PositionGroup, TransparencyGroup };
            for (int i = 0; i < groups.Length; i++)
            {
                groups[i].Dock = DockStyle.None;
                groups[i].Width = leftW;
                groups[i].Margin = new Padding(3, i == 0 ? 3 : 8, 3, 3);
                left.Controls.Add(groups[i]);
            }
            int height = left.PreferredSize.Height - 6;
            PreviewGroup.Dock = DockStyle.None;
            PreviewGroup.Size = new Size(Math.Max(200, width - leftW - 12), height);
            PreviewGroup.Margin = new Padding(3);
            var row = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            row.Controls.Add(left);
            row.Controls.Add(PreviewGroup);
            return row;
        }

        // Reads the picture for the preview if that has not been done (the first time the preview is shown).
        public void EnsurePreview()
        {
            if (_previewLoaded || !_preview.Visible) return;
            ShowPreview();
        }

        // Shows the picture of the chosen file in the preview (with the frame in _cropValue and the turn in _rotationValue).
        private void ShowPreview()
        {
            _previewLoaded = true;
            Bitmap picture = null;
            string message = "";
            try
            {
                if (string.IsNullOrEmpty(_path)) message = "壁紙の画像を選ぶと、ここに表示します。";
                else if (!File.Exists(_path)) message = "ファイルが見つかりません。";
                else
                {
                    picture = Wallpaper_Load(_path);
                    if (picture == null) message = "この画像は読み込めません。";
                }
            }
            catch (Exception)
            {
                if (picture != null) { picture.Dispose(); picture = null; }
                message = "この画像は読み込めません。";
            }
            SetPicture(picture, message);
        }

        private void SetPicture(Bitmap picture, string message)
        {
            _preview.Message = message;
            _preview.SetImage(picture, _cropValue, _rotationValue, _limitCrop.Checked, _flip.Checked);
            ShowPreviewState(picture != null, message);
        }

        private void ShowPreviewState(bool hasPicture, string message)
        {
            _resetCrop.Enabled = hasPicture;
            _limitCrop.Enabled = hasPicture;
            _flip.Enabled = hasPicture;
            if (!hasPicture) _preview.Message = message;
        }

        public WallpaperOptions Wallpaper
        {
            get
            {
                var o = new WallpaperOptions { Path = _path, KeepRatio = _keep.Checked, Tile = _tile.Checked, Crop = _cropValue, Rotation = _rotationValue, LimitCrop = _limitCrop.Checked, Flip = _flip.Checked };
                for (int i = 0; i < 4; i++) if (_fit[i].Checked) o.Fit = (WallpaperFit)i;
                for (int i = 0; i < 3; i++)
                {
                    if (_h[i].Checked) o.HAlign = (WallpaperHAlign)i;
                    if (_v[i].Checked) o.VAlign = (WallpaperVAlign)i;
                }
                return o;
            }
        }

        public int RowTransparency { get { return _transparency.Value * WorkbookSettings.TransparencyStep; } }

        public bool TextOutline { get { return _outline.Checked; } }

        private WallpaperFit CurrentFit
        {
            get
            {
                for (int i = 0; i < 4; i++) if (_fit[i].Checked) return (WallpaperFit)i;
                return WallpaperFit.FitBoth;
            }
        }

        private void Browse()
        {
            try
            {
                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = "壁紙にする画像を選ぶ";
                    dialog.Filter = "画像ファイル (*.jpg;*.jpeg;*.png)|*.jpg;*.jpeg;*.png";
                    dialog.CheckFileExists = true;
                    try
                    {
                        string dir = string.IsNullOrEmpty(_path) ? null : Path.GetDirectoryName(_path);
                        if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir)) dialog.InitialDirectory = dir;
                    }
                    catch (Exception) { }
                    if (dialog.ShowDialog() != DialogResult.OK) return;
                    Bitmap picture = Wallpaper_Load(dialog.FileName);
                    if (picture == null)
                    {
                        MessageBox.Show("この画像は読み込めません。jpg か png の画像を選んでください。", "タブ表示の設定", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // a file is chosen (the same one again too): the frame, the turn and the limit go back
                    // (the box first: turning it on fits the old frame, which reports a change that must not undo the reset)
                    _limitCrop.Checked = true;
                    _flip.Checked = false;
                    _cropValue = WallpaperCrop.Full;
                    _rotationValue = WallpaperRotation.None;
                    _path = dialog.FileName;
                    _pathBox.Text = _path;
                    _previewLoaded = true;
                    SetPicture(picture, ""); // (the preview keeps the picture)
                    Refresh();
                }
            }
            catch (Exception)
            {
                // The file dialog could not be shown: the wallpaper stays as it was.
            }
        }

        private static Bitmap Wallpaper_Load(string path)
        {
            return VerticalSheetTabs.Wallpaper.Load(path);
        }

        // Enables what applies to the chosen way of fitting, and writes the notes.
        // The rows are stacked from the bottom: the wallpaper's region differs, so does the note about it.
        private bool _stacked;
        public bool Stacked { get { return _stacked; } set { if (_stacked == value) return; _stacked = value; Refresh(); } }

        private void Refresh()
        {
            WallpaperFit fit = CurrentFit;
            bool keep = _keep.Checked;
            bool ratioOk = WallpaperOptions.KeepRatioApplies(fit);
            _ratioLabel.Enabled = ratioOk;
            _stretch.Enabled = ratioOk;
            _keep.Enabled = ratioOk;
            bool tileOk = WallpaperOptions.TileApplies(fit);
            _tile.Enabled = tileOk;
            bool hOk = WallpaperOptions.HAlignApplies(fit, keep), vOk = WallpaperOptions.VAlignApplies(fit, keep);
            _hLabel.Enabled = hOk;
            _vLabel.Enabled = vOk;
            foreach (RadioButton r in _h) r.Enabled = hOk;
            foreach (RadioButton r in _v) r.Enabled = vOk;

            _tileNote.Text = fit == WallpaperFit.FitBoth ? "「縦横に合わせる」では、画像が領域いっぱいになるため、タイルは使えません。"
                : fit == WallpaperFit.FitWidth ? "「横幅に合わせる」では、縦方向に繰り返します。"
                : fit == WallpaperFit.FitHeight ? "「縦幅に合わせる」では、横方向に繰り返します。"
                : "「元のサイズ」では、縦と横に繰り返します。";
            string where = fit == WallpaperFit.FitBoth ? (keep ? "画像のどの部分を見せるかを選びます (はみ出した部分は切れます)。" : "比率を変えて引き伸ばすため、表示位置は選べません。")
                : fit == WallpaperFit.FitWidth ? "上下位置だけ選べます。"
                : fit == WallpaperFit.FitHeight ? "左右位置だけ選べます。"
                : "左右位置と上下位置を選べます。";
            _posNote.Text = (tileOk && _tile.Checked ? "タイルのときは、並べ始める位置になります。 " : "") + where;

            int percent = RowTransparency;
            _transparencyText.Text = percent + " %";
            _transparencyNote.Text = percent == 0
                ? (_stacked ? "透過なし: 壁紙は、縦タブの上端から、最初の行の上までに表示します。" : "透過なし: 壁紙は、最後の行の下から「+」の上までに表示します。")
                : (_stacked ? "透過あり: 壁紙は、縦タブの上端から、「全開」「全閉」の帯の上までに表示し、行の色が " + percent + "% 透けます。" : "透過あり: 壁紙は、先頭の行の上から「+」の上までに表示し、行の色が " + percent + "% 透けます。");
            _outline.Enabled = percent > 0;
            _outlineNote.Text = percent > 0 ? "文字を、読みやすい色の縁で囲みます。" : "透過なしのときは、文字は読みやすいため、縁は付けません (選べません)。";

            if (string.IsNullOrEmpty(_path)) _status.Text = "壁紙は設定されていません。jpg、png の画像を選べます。";
            else if (!File.Exists(_path)) _status.Text = "ファイルが見つかりません (壁紙なしで表示します)。";
            else _status.Text = "jpg、png の画像を選べます。";
        }

        private static FlowLayoutPanel Stack()
        {
            return new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
        }

        // Radio buttons side by side; equalGaps = the same space after each (otherwise every button takes the width of the widest one).
        private static FlowLayoutPanel Row(Font font, bool equalGaps, params RadioButton[] radios)
        {
            var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            int widest = 0;
            foreach (RadioButton r in radios)
            {
                r.AutoSize = true;
                r.Font = font;
                widest = Math.Max(widest, r.PreferredSize.Width);
            }
            int gap = TextRenderer.MeasureText("0000", font).Width;
            foreach (RadioButton r in radios)
            {
                r.Margin = new Padding(3, 3, equalGaps ? gap : widest - r.PreferredSize.Width + gap, 3);
                panel.Controls.Add(r);
            }
            return panel;
        }

        private static GroupBox Group(string title, Font font, Control content, bool dockTop)
        {
            var group = new GroupBox { Text = title, Font = font, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            group.Controls.Add(content);
            Size size;
            if (dockTop)
            {
                content.Location = group.DisplayRectangle.Location;
                size = content.PreferredSize;
                content.Dock = DockStyle.Top;
                group.Size = new Size(size.Width + group.Padding.Horizontal + 8, content.Top + size.Height + group.Padding.Bottom + 2);
                return group;
            }
            content.Location = group.DisplayRectangle.Location;
            size = content.PreferredSize;
            group.Size = new Size(size.Width + group.Padding.Horizontal + 8, content.Top + size.Height + group.Padding.Bottom + 2);
            return group;
        }
    }
}
