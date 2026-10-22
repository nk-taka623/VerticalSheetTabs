using System;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // "Tab display settings" dialog (Agents.md 2.1): display mode and vertical tab side.
    // Opened from the ribbon button and, in Step5, from the tab context menu.
    public class SettingsForm : Form
    {
        private readonly RadioButton _modeHorizontal = new RadioButton();
        private readonly RadioButton _modeVertical = new RadioButton();
        private readonly RadioButton _modeBoth = new RadioButton();
        private readonly RadioButton _sideLeft = new RadioButton();
        private readonly RadioButton _sideRight = new RadioButton();
        private GroupBox _sideGroup;
        private readonly CheckBox _outlineShow = new CheckBox();
        private readonly TrackBar _outlineLevel = new TrackBar();
        private readonly Panel _outlineSwatch = new Panel();
        private readonly Label _outlineLevelText = new Label();
        private GroupBox _outlineGroup;
        private readonly CheckBox _nameRule = new CheckBox();
        private GroupBox _ruleGroup;
        private readonly RadioButton _styleStandard = new RadioButton();
        private readonly RadioButton _styleDouble = new RadioButton();
        private readonly RadioButton _styleFrosted = new RadioButton();
        private readonly RadioButton _styleGlossy = new RadioButton();
        private readonly RadioButton _styleFrosted2 = new RadioButton();
        private readonly RadioButton _styleInset = new RadioButton();
        private GroupBox _styleGroup;
        private readonly RadioButton _markerNone = new RadioButton();
        private readonly RadioButton _markerUnderline = new RadioButton();
        private readonly RadioButton _markerTriangle = new RadioButton();
        private readonly RadioButton _markerBars = new RadioButton();
        private GroupBox _markerGroup;
        // The samples of the looks: they show the active-sheet mark that is chosen above them.
        private readonly System.Collections.Generic.List<TabStrip> _samples = new System.Collections.Generic.List<TabStrip>();

        public SettingsForm(TabMode mode, TabSide side, bool outlineVisible, int outlineLevel, bool nameRule, SelectStyle selectStyle, ActiveMarker activeMarker)
        {
            Text = "タブ表示の設定";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            _modeHorizontal.Text = "横タブのみ表示 (縦タブは表示しない)";
            _modeVertical.Text = "縦タブのみ表示 (横タブは表示しない)";
            _modeBoth.Text = "横タブと縦タブの両方を表示";
            _sideLeft.Text = "左";
            _sideRight.Text = "右";

            var modeGroup = CreateRowGroup("表示モード", _modeHorizontal, _modeVertical, _modeBoth);
            _sideGroup = CreateRowGroup("縦タブの位置", _sideLeft, _sideRight);

            _outlineGroup = CreateOutlineGroup();
            _ruleGroup = CreateRuleGroup();
            _markerNone.Text = "マーカーなし";
            _markerUnderline.Text = "下線マーカー (名前の下に線)";
            _markerTriangle.Text = "三角マーカー (左右の端に三角)";
            _markerBars.Text = "両端バーマーカー (左右の端に帯)";
            _markerGroup = CreateRowGroup("アクティブシートのマーカー表示", _markerUnderline, _markerTriangle, _markerBars, _markerNone);
            _markerNone.Checked = activeMarker == ActiveMarker.None;
            _markerUnderline.Checked = activeMarker == ActiveMarker.Underline;
            _markerTriangle.Checked = activeMarker == ActiveMarker.Triangle;
            _markerBars.Checked = activeMarker == ActiveMarker.EdgeBars;
            _styleGroup = CreateStyleGroup();
            foreach (RadioButton r in new[] { _markerNone, _markerUnderline, _markerTriangle, _markerBars })
            {
                r.CheckedChanged += (s, e) => { foreach (TabStrip sample in _samples) sample.ActiveMarker = ActiveMarker; };
            }

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(88, 0) };
            var cancel = new Button { Text = "キャンセル", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(88, 0) };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 10, 0, 0)
            };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            AcceptButton = ok;
            CancelButton = cancel;

            // One column: both group boxes and the button row get the width of the widest one.
            var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            modeGroup.Dock = DockStyle.Fill;
            _sideGroup.Dock = DockStyle.Fill;
            _sideGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(modeGroup);
            layout.Controls.Add(_sideGroup);
            _outlineGroup.Dock = DockStyle.Fill;
            _outlineGroup.Margin = new Padding(3, 8, 3, 3);
            _markerGroup.Dock = DockStyle.Fill;
            _markerGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_markerGroup);
            _styleGroup.Dock = DockStyle.Fill;
            _styleGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_styleGroup);
            layout.Controls.Add(_outlineGroup);
            _ruleGroup.Dock = DockStyle.Fill;
            _ruleGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_ruleGroup);
            // The bottom row: the add-in's version (the same as the release tag), copyright and license at the left, the buttons at the right.
            var version = new Label
            {
                Text = "バージョン " + AddInVersion.Text + "\r\n" + AddInVersion.Copyright + "\r\n" + AddInVersion.License,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(3, 6, 3, 0)
            };
            var bottom = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, AutoSize = true, Dock = DockStyle.Fill, Margin = Padding.Empty };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottom.Controls.Add(version, 0, 0);
            bottom.Controls.Add(buttons, 1, 0);
            layout.Controls.Add(bottom);
            Controls.Add(layout);

            _modeHorizontal.Checked = mode == TabMode.HorizontalOnly;
            _modeVertical.Checked = mode == TabMode.VerticalOnly;
            _modeBoth.Checked = mode == TabMode.Both;
            _sideLeft.Checked = side == TabSide.Left;
            _sideRight.Checked = side == TabSide.Right;

            _nameRule.Checked = nameRule;
            _styleStandard.Checked = selectStyle == SelectStyle.Standard;
            _styleDouble.Checked = selectStyle == SelectStyle.DoubleFrame;
            _styleFrosted.Checked = selectStyle == SelectStyle.FrostedGlass;
            _styleGlossy.Checked = selectStyle == SelectStyle.GlossyGlass;
            _styleFrosted2.Checked = selectStyle == SelectStyle.FrostedGlass2;
            _styleInset.Checked = selectStyle == SelectStyle.Inset;
            // the buttons sit in different panels (each next to its samples), so the form keeps them exclusive
            RadioButton[] styles = { _styleStandard, _styleDouble, _styleGlossy, _styleFrosted, _styleFrosted2, _styleInset };
            foreach (RadioButton r in styles)
            {
                RadioButton self = r;
                self.CheckedChanged += (s, e) =>
                {
                    if (!self.Checked) return;
                    foreach (RadioButton other in styles) if (other != self) other.Checked = false;
                };
            }
            _outlineShow.Checked = outlineVisible;
            _outlineLevel.Value = Math.Max(1, Math.Min(TabStrip.OutlineLevels, outlineLevel));
            _outlineShow.CheckedChanged += (s, e) => UpdateOutline();
            _outlineLevel.ValueChanged += (s, e) => UpdateOutline();
            UpdateOutline();

            _modeHorizontal.CheckedChanged += (s, e) => UpdateSideEnabled();
            UpdateSideEnabled();
        }

        public TabMode Mode
        {
            get
            {
                return _modeHorizontal.Checked ? TabMode.HorizontalOnly
                    : _modeVertical.Checked ? TabMode.VerticalOnly : TabMode.Both;
            }
        }

        public TabSide Side
        {
            get { return _sideRight.Checked ? TabSide.Right : TabSide.Left; }
        }

        public SelectStyle SelectStyle
        {
            get
            {
                return _styleDouble.Checked ? SelectStyle.DoubleFrame
                    : _styleFrosted.Checked ? SelectStyle.FrostedGlass
                    : _styleGlossy.Checked ? SelectStyle.GlossyGlass
                    : _styleFrosted2.Checked ? SelectStyle.FrostedGlass2
                    : _styleInset.Checked ? SelectStyle.Inset : SelectStyle.Standard;
            }
        }

        public ActiveMarker ActiveMarker
        {
            get
            {
                return _markerNone.Checked ? ActiveMarker.None : _markerUnderline.Checked ? ActiveMarker.Underline
                    : _markerBars.Checked ? ActiveMarker.EdgeBars : ActiveMarker.Triangle;
            }
        }

        public bool OutlineVisible { get { return _outlineShow.Checked; } }

        public bool NameRule { get { return _nameRule.Checked; } }

        public int OutlineLevel { get { return _outlineLevel.Value; } }

        private void UpdateOutline()
        {
            _outlineLevel.Enabled = _outlineShow.Checked;
            _outlineLevelText.Enabled = _outlineShow.Checked;
            _outlineLevelText.Text = _outlineLevel.Value + " / " + TabStrip.OutlineLevels;
            _outlineSwatch.BackColor = _outlineShow.Checked
                ? TabStrip.OutlineColorAt(TabTheme.Current.BarBack, _outlineLevel.Value)
                : SystemColors.Control;
        }

        // The look of selected folders and sheets: the six types in three columns and two rows, each with samples (not selected, a folder and a plain sheet selected, a colored
        // sheet selected) drawn by the tab strip itself.
        private GroupBox CreateStyleGroup()
        {
            var group = new GroupBox { Text = "フォルダ＆シートの選択デザイン", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            _styleStandard.Text = "標準 (Excelの横タブと同じ見た目)";
            _styleDouble.Text = "ダブルフレーム (二重の枠で、目立たせる)";
            _styleGlossy.Text = "グロッシーグラス (光沢のあるガラス)";
            _styleFrosted.Text = "フロストグラス１ (縁が光る磨りガラス)";
            _styleFrosted2.Text = "フロストグラス２ (行の色を帯びた磨りガラス)";
            _styleInset.Text = "インセット (押し込まれたように凹む)";
            var panel = new TableLayoutPanel
            {
                ColumnCount = 3,
                RowCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty
            };
            panel.Controls.Add(CreateStyleBlock(_styleStandard, SelectStyle.Standard), 0, 0);
            panel.Controls.Add(CreateStyleBlock(_styleDouble, SelectStyle.DoubleFrame), 1, 0);
            panel.Controls.Add(CreateStyleBlock(_styleGlossy, SelectStyle.GlossyGlass), 2, 0);
            panel.Controls.Add(CreateStyleBlock(_styleFrosted, SelectStyle.FrostedGlass), 0, 1);
            panel.Controls.Add(CreateStyleBlock(_styleFrosted2, SelectStyle.FrostedGlass2), 1, 1);
            panel.Controls.Add(CreateStyleBlock(_styleInset, SelectStyle.Inset), 2, 1);
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        private FlowLayoutPanel CreateStyleBlock(RadioButton radio, SelectStyle style)
        {
            radio.AutoSize = true;
            var block = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 0, 16, 8)
            };
            block.Controls.Add(radio);
            var samples = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(16, 0, 0, 0)
            };
            string[][] cells =
            {
                new[] { "非選択", "none" },
                new[] { "選択 (色なしのシート)", "plain" },
                new[] { "選択 (色付きのシート)", "colored" }
            };
            foreach (string[] cell in cells)
            {
                var strip = new TabStrip();
                strip.SetSample(style, cell[1], ActiveMarker);
                _samples.Add(strip);
                strip.Size = strip.SampleSize(120);
                strip.Margin = new Padding(0, 0, 0, 0);
                strip.SampleClicked += () => radio.Checked = true;
                var caption = new Label { Text = cell[0], AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 0, 0, 2) };
                var one = new FlowLayoutPanel
                {
                    FlowDirection = FlowDirection.TopDown,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    WrapContents = false,
                    Margin = new Padding(0, 0, 8, 0)
                };
                one.Controls.Add(caption);
                one.Controls.Add(strip);
                samples.Controls.Add(one);
            }
            block.Controls.Add(samples);
            return block;
        }

        // The sheet-name rule: on/off, with a short explanation.
        private GroupBox CreateRuleGroup()
        {
            var group = new GroupBox { Text = "シート名のルール", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            _nameRule.Text = "シート名のルールを使う";
            _nameRule.AutoSize = true;
            var note = new Label
            {
                Text = "横タブのシート名が「A | B」のとき、縦タブでは、フォルダ A の中の\r\nシート B として表示します。縦タブでの、フォルダの作成、移動、\r\n名前の変更も、横タブのシート名に反映します。",
                AutoSize = true,
                Margin = new Padding(3, 6, 3, 0)
            };
            var panel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = Padding.Empty
            };
            panel.Controls.Add(_nameRule);
            panel.Controls.Add(note);
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        // The dashed outline around an open folder: on/off, and a slider for the color level. The swatch
        // shows the line color on the tab bar background.
        private GroupBox CreateOutlineGroup()
        {
            var group = new GroupBox { Text = "フォルダを囲む点線", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            _outlineShow.Text = "点線を表示する";
            _outlineShow.AutoSize = true;
            var caption = new Label { Text = "色 (左: 背景に近い色、右: 背景から最も離れた色)", AutoSize = true, Margin = new Padding(3, 6, 3, 0) };
            _outlineLevel.Minimum = 1;
            _outlineLevel.Maximum = TabStrip.OutlineLevels;
            _outlineLevel.TickFrequency = 1;
            _outlineLevel.SmallChange = 1;
            _outlineLevel.LargeChange = 1;
            _outlineLevel.AutoSize = false;
            _outlineLevel.Size = new Size(260, 36);
            _outlineSwatch.Size = new Size(36, 24);
            _outlineSwatch.BorderStyle = BorderStyle.FixedSingle;
            _outlineSwatch.BackColor = TabTheme.Current.BarBack;
            _outlineLevelText.AutoSize = true;
            _outlineLevelText.Text = TabStrip.OutlineLevels + " / " + TabStrip.OutlineLevels; // for measuring
            var sliderRow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            sliderRow.Controls.Add(_outlineLevel);
            sliderRow.Controls.Add(_outlineSwatch);
            sliderRow.Controls.Add(_outlineLevelText);
            _outlineSwatch.Margin = new Padding(8, 6, 4, 0);
            _outlineLevelText.Margin = new Padding(4, 10, 0, 0);
            var panel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = Padding.Empty
            };
            panel.Controls.Add(_outlineShow);
            panel.Controls.Add(caption);
            panel.Controls.Add(sliderRow);
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        // The side only matters while a vertical tab bar is shown.
        private void UpdateSideEnabled()
        {
            _sideGroup.Enabled = !_modeHorizontal.Checked;
        }

        // A group box with its radio buttons side by side, at equal steps: the widest one plus a gap of about
        // four half-width characters.
        private static GroupBox CreateRowGroup(string title, params RadioButton[] radios)
        {
            var group = new GroupBox
            {
                Text = title,
                Font = SystemFonts.MessageBoxFont,
                AutoSize = false,
                Padding = new Padding(8, 4, 8, 6)
            };
            var panel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = Padding.Empty
            };
            int widest = 0;
            foreach (RadioButton r in radios)
            {
                r.AutoSize = true;
                r.Font = group.Font;
                widest = Math.Max(widest, r.PreferredSize.Width);
            }
            int gap = TextRenderer.MeasureText("0000", group.Font).Width;
            foreach (RadioButton r in radios)
            {
                r.Margin = new Padding(3, 3, widest - r.PreferredSize.Width + gap, 3);
                panel.Controls.Add(r);
            }
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(
                content.Width + group.Padding.Horizontal + 8,
                panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }
    }

    // Lets the dialog be owned by the Excel main window.
    public class ExcelWindowOwner : IWin32Window
    {
        public ExcelWindowOwner(int hwnd)
        {
            Handle = new IntPtr(hwnd);
        }

        public IntPtr Handle { get; private set; }
    }
}
