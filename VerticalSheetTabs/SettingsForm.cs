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
        private readonly RadioButton _styleStripes = new RadioButton();
        private readonly RadioButton _styleCrystal = new RadioButton();
        private GroupBox _styleGroup;
        private readonly RadioButton _markerNone = new RadioButton();
        private readonly RadioButton _markerUnderline = new RadioButton();
        private readonly RadioButton _markerTriangle = new RadioButton();
        private readonly RadioButton _markerBars = new RadioButton();
        private readonly RadioButton _markerBadge = new RadioButton();
        private GroupBox _markerGroup;
        private readonly RadioButton[] _wobbleTypes = new RadioButton[8]; // index = WobbleType (None, 1 to 6, Random)
        private readonly TrackBar _wobbleAmp = new TrackBar();
        private readonly CheckBox _wobbleSound = new CheckBox();
        private readonly TrackBar _wobbleVolume = new TrackBar();
        private readonly Label _wobbleVolumeText = new Label();
        private readonly TrackBar _wobbleDur = new TrackBar();
        private readonly Label _wobbleAmpText = new Label();
        private readonly Label _wobbleDurText = new Label();
        private TabStrip _wobbleSample;
        private System.Windows.Forms.Timer _wobbleTimer;
        private GroupBox _wobbleGroup;
        private readonly Button _wobbleReset = new Button();
        private readonly CheckBox _wobbleUseStyle = new CheckBox(); // the sample shows the chosen look and marker, or the plain one
        // The samples of the looks: they show the active-sheet mark that is chosen above them.
        private readonly System.Collections.Generic.List<TabStrip> _samples = new System.Collections.Generic.List<TabStrip>();

        public SettingsForm(TabMode mode, TabSide side, bool outlineVisible, int outlineLevel, bool nameRule, SelectStyle selectStyle, ActiveMarker activeMarker, WobbleType wobble, int wobbleAmp, int wobbleDur, bool wobbleSound, int wobbleVolume)
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
            _markerBadge.Text = "チェックバッジマーカー (名前の左に印)";
            _markerGroup = CreateRowGroup("アクティブシートのマーカー表示", _markerUnderline, _markerTriangle, _markerBars, _markerBadge, _markerNone);
            _markerNone.Checked = activeMarker == ActiveMarker.None;
            _markerUnderline.Checked = activeMarker == ActiveMarker.Underline;
            _markerTriangle.Checked = activeMarker == ActiveMarker.Triangle;
            _markerBars.Checked = activeMarker == ActiveMarker.EdgeBars;
            _markerBadge.Checked = activeMarker == ActiveMarker.CheckBadge;
            _styleGroup = CreateStyleGroup();
            _wobbleGroup = CreateWobbleGroup(wobble, wobbleAmp, wobbleDur, wobbleSound, wobbleVolume);
            foreach (RadioButton r in new[] { _markerNone, _markerUnderline, _markerTriangle, _markerBars, _markerBadge })
            {
                r.CheckedChanged += (s, e) => { foreach (TabStrip sample in _samples) sample.ActiveMarker = ActiveMarker; RefreshWobbleSample(); };
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
            _outlineGroup.Dock = DockStyle.Fill;
            _outlineGroup.Margin = new Padding(3, 8, 3, 3);
            _markerGroup.Dock = DockStyle.Fill;
            _markerGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_markerGroup);
            _styleGroup.Dock = DockStyle.Fill;
            _styleGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_styleGroup);
            _wobbleGroup.Dock = DockStyle.Fill;
            _wobbleGroup.Margin = new Padding(3, 8, 3, 3);
            layout.Controls.Add(_wobbleGroup);
            // the display mode and the side of the vertical tabs stand side by side in the top row: the side group takes about a third of the width, the mode group the rest
            modeGroup.Dock = DockStyle.None;
            _sideGroup.Dock = DockStyle.None;
            int topHeight = Math.Max(modeGroup.Height, _sideGroup.Height);
            int sideWidth = (_styleGroup.Width - 12) * 34 / 100;
            modeGroup.Size = new Size(_styleGroup.Width - sideWidth - 12, topHeight);
            _sideGroup.Size = new Size(sideWidth, topHeight);
            modeGroup.Margin = new Padding(3);
            _sideGroup.Margin = new Padding(3);
            var top = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            top.Controls.Add(modeGroup);
            top.Controls.Add(_sideGroup);
            layout.Controls.Add(top);
            layout.Controls.SetChildIndex(top, 0);
            // the dotted outline and the sheet-name rule stand side by side, each half the width of the groups above, both as high as the higher one
            _outlineGroup.Dock = DockStyle.None;
            _ruleGroup.Dock = DockStyle.None;
            int half = (_styleGroup.Width - 12) / 2;
            int pairHeight = Math.Max(_outlineGroup.Height, _ruleGroup.Height);
            _outlineGroup.Size = new Size(half, pairHeight);
            _ruleGroup.Size = new Size(half, pairHeight);
            _outlineGroup.Margin = new Padding(3);
            _ruleGroup.Margin = new Padding(3);
            var pair = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(0, 5, 0, 0) };
            pair.Controls.Add(_outlineGroup);
            pair.Controls.Add(_ruleGroup);
            layout.Controls.Add(pair);
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
            _styleStripes.Checked = selectStyle == SelectStyle.Stripes;
            _styleCrystal.Checked = selectStyle == SelectStyle.Crystal;
            // the buttons sit in different panels (each next to its samples), so the form keeps them exclusive
            RadioButton[] styles = { _styleStandard, _styleDouble, _styleGlossy, _styleStripes, _styleFrosted, _styleFrosted2, _styleInset, _styleCrystal };
            foreach (RadioButton r in styles)
            {
                RadioButton self = r;
                self.CheckedChanged += (s, e) =>
                {
                    if (!self.Checked) return;
                    foreach (RadioButton other in styles) if (other != self) other.Checked = false;
                    RefreshWobbleSample();
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
                    : _styleInset.Checked ? SelectStyle.Inset
                    : _styleStripes.Checked ? SelectStyle.Stripes
                    : _styleCrystal.Checked ? SelectStyle.Crystal : SelectStyle.Standard;
            }
        }

        public ActiveMarker ActiveMarker
        {
            get
            {
                return _markerNone.Checked ? ActiveMarker.None : _markerUnderline.Checked ? ActiveMarker.Underline
                    : _markerBars.Checked ? ActiveMarker.EdgeBars : _markerBadge.Checked ? ActiveMarker.CheckBadge : ActiveMarker.Triangle;
            }
        }

        public WobbleType Wobble
        {
            get
            {
                for (int i = 0; i < _wobbleTypes.Length; i++) if (_wobbleTypes[i].Checked) return (WobbleType)i;
                return WobbleType.None;
            }
        }

        public int WobbleAmp { get { return _wobbleAmp.Value; } }

        public bool WobbleSound { get { return _wobbleSound.Checked; } }

        public int WobbleVolume { get { return _wobbleVolume.Value; } }

        public int WobbleDur { get { return _wobbleDur.Value; } }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_wobbleTimer != null) _wobbleTimer.Stop();
            base.OnFormClosed(e);
        }

        // The wobble of the selected sheet (Agents.md 12.141): the motion (none, 1 to 6, random), its size and length, and one live sample that
        // plays the chosen motion again and again (a new random one each time for "7 ランダム"). The sample follows the chosen look and marker.
        private GroupBox CreateWobbleGroup(WobbleType wobble, int amp, int dur, bool sound, int volume)
        {
            string[] names = { "なし", "1 シェイク", "2 ウォブル", "3 ジェリー", "4 バウンス", "5 パルス", "6 スイング", "7 ランダム" };
            for (int i = 0; i < _wobbleTypes.Length; i++)
            {
                _wobbleTypes[i] = new RadioButton { Text = names[i], Checked = (int)wobble == i };
            }
            GroupBox group = CreateRowGroup("開閉したフォルダ＆選択したシートのアニメーション", _wobbleTypes);
            Control radios = group.Controls[0];
            group.Controls.Clear();

            _wobbleAmp.Minimum = WobbleMotion.AmpMin; _wobbleAmp.Maximum = WobbleMotion.AmpMax;
            _wobbleAmp.TickFrequency = 20; _wobbleAmp.SmallChange = 10; _wobbleAmp.LargeChange = 20;
            _wobbleAmp.Value = Math.Max(WobbleMotion.AmpMin, Math.Min(WobbleMotion.AmpMax, amp));
            _wobbleDur.Minimum = WobbleMotion.DurMin; _wobbleDur.Maximum = WobbleMotion.DurMax;
            _wobbleDur.TickFrequency = 100; _wobbleDur.SmallChange = 50; _wobbleDur.LargeChange = 100;
            _wobbleDur.Value = Math.Max(WobbleMotion.DurMin, Math.Min(WobbleMotion.DurMax, dur));
            foreach (TrackBar t in new[] { _wobbleAmp, _wobbleDur }) { t.AutoSize = false; t.Size = new Size(240, 36); t.Margin = new Padding(3, 0, 3, 0); }
            foreach (Label l in new[] { _wobbleAmpText, _wobbleDurText }) { l.AutoSize = false; l.Size = new Size(70, 36); l.TextAlign = ContentAlignment.MiddleLeft; }
            var ampLabel = new Label { Text = "揺れの大きさ", AutoSize = false, Size = new Size(100, 36), TextAlign = ContentAlignment.MiddleLeft };
            var durLabel = new Label { Text = "揺れの長さ", AutoSize = false, Size = new Size(100, 36), TextAlign = ContentAlignment.MiddleLeft };

            _wobbleSample = new TabStrip();
            _wobbleSample.TabStop = false;
            _wobbleSample.SampleClicked += PlayWobbleSample;
            _wobbleSample.Size = _wobbleSample.SingleSampleSize(120);
            var sampleBox = new Panel { Size = _wobbleSample.Size + new Size(24, 8), Margin = new Padding(16, 0, 3, 0) };
            _wobbleSample.Location = new Point(12, 4);
            sampleBox.Controls.Add(_wobbleSample);

            // right of the sample: the motions, below them the two sliders side by side, below those the check box
            var ampBox = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(0, 0, 24, 0) };
            ampBox.Controls.Add(ampLabel); ampBox.Controls.Add(_wobbleAmp); ampBox.Controls.Add(_wobbleAmpText);
            var durBox = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            durBox.Controls.Add(durLabel); durBox.Controls.Add(_wobbleDur); durBox.Controls.Add(_wobbleDurText); durBox.Controls.Add(_wobbleReset);
            var sliders = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            sliders.Controls.Add(ampBox); sliders.Controls.Add(durBox);
            _wobbleUseStyle.Text = "サンプルの見た目を、選択したデザインとマーカー表示に合わせる";
            _wobbleUseStyle.AutoSize = true;
            _wobbleUseStyle.Margin = new Padding(3, 2, 3, 3);
            _wobbleUseStyle.CheckedChanged += (s, e) => { RefreshWobbleSample(); PlayWobbleSample(); };
            var right = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            radios.Margin = Padding.Empty;
            right.Controls.Add(radios);
            right.Controls.Add(sliders);
            _wobbleReset.Text = "揺れのリセット";
            _wobbleReset.AutoSize = true;
            _wobbleReset.Margin = new Padding(12, 4, 3, 0);
            _wobbleReset.Click += (s, e) =>
            {
                _wobbleAmp.Value = WobbleMotion.AmpDefault;
                _wobbleDur.Value = WobbleMotion.DurDefault;
            };
            // the sound effect: a check box and a volume slider, on the line of the other check box
            _wobbleSound.Text = "効果音を鳴らす";
            _wobbleSound.AutoSize = true;
            _wobbleSound.Checked = sound;
            _wobbleSound.Margin = new Padding(40, 2, 3, 3);
            _wobbleVolume.Minimum = 0; _wobbleVolume.Maximum = 100;
            _wobbleVolume.TickFrequency = 10; _wobbleVolume.SmallChange = 5; _wobbleVolume.LargeChange = 10;
            _wobbleVolume.Value = Math.Max(0, Math.Min(100, volume));
            _wobbleVolume.AutoSize = false; _wobbleVolume.Size = new Size(200, 30); _wobbleVolume.Margin = new Padding(3, 0, 3, 0);
            _wobbleVolumeText.AutoSize = false; _wobbleVolumeText.Size = new Size(60, 30); _wobbleVolumeText.TextAlign = ContentAlignment.MiddleLeft;
            var volumeLabel = new Label { Text = "音量", AutoSize = false, Size = new Size(48, 30), TextAlign = ContentAlignment.MiddleRight };
            var lastLine = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            lastLine.Controls.Add(_wobbleUseStyle);
            lastLine.Controls.Add(_wobbleSound);
            lastLine.Controls.Add(volumeLabel);
            lastLine.Controls.Add(_wobbleVolume);
            lastLine.Controls.Add(_wobbleVolumeText);
            right.Controls.Add(lastLine);
            sampleBox.Margin = new Padding(3, 3, 12, 3);
            var stack = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            // the sample stands in the middle of the group, top to bottom
            sampleBox.Height = Math.Max(sampleBox.Height, right.PreferredSize.Height);
            _wobbleSample.Location = new Point(12, (sampleBox.Height - _wobbleSample.Height) / 2);
            stack.Controls.Add(sampleBox);
            stack.Controls.Add(right);
            group.Controls.Add(stack);
            stack.Location = group.DisplayRectangle.Location;
            Size content = stack.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, stack.Top + content.Height + group.Padding.Bottom + 2);

            EventHandler changed = (s, e) =>
            {
                UpdateWobbleTexts();
                RefreshWobbleSample();
                PlayWobbleSample();
            };
            foreach (RadioButton r in _wobbleTypes) r.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) changed(s, e); };
            _wobbleAmp.ValueChanged += changed;
            _wobbleDur.ValueChanged += changed;
            _wobbleSound.CheckedChanged += (s, e) => { UpdateWobbleTexts(); RefreshWobbleSample(); if (_wobbleSound.Checked) PlayWobbleSample(); };
            _wobbleVolume.ValueChanged += (s, e) => { UpdateWobbleTexts(); RefreshWobbleSample(); };
            _wobbleVolume.MouseUp += (s, e) => { if (_wobbleSound.Checked) PlayWobbleSample(); };
            UpdateWobbleTexts();
            _wobbleSample.HandleCreated += (s, e) =>
            {
                RefreshWobbleSample();
                BeginInvoke(new Action(PlayWobbleSample));
            };
            _wobbleTimer = new System.Windows.Forms.Timer { Interval = 2200 };
            _wobbleTimer.Tick += (s, e) => PlayWobbleSample();
            _wobbleTimer.Start();
            return group;
        }

        private void UpdateWobbleTexts()
        {
            _wobbleAmpText.Text = _wobbleAmp.Value + " %";
            _wobbleDurText.Text = (_wobbleDur.Value / 1000.0).ToString("0.0#") + " 秒";
            _wobbleAmp.Enabled = _wobbleDur.Enabled = _wobbleReset.Enabled = _wobbleSound.Enabled = Wobble != WobbleType.None;
            _wobbleVolumeText.Text = _wobbleVolume.Value + " %";
            _wobbleVolume.Enabled = _wobbleVolumeText.Enabled = Wobble != WobbleType.None && _wobbleSound.Checked;
        }

        // The look and marker may be chosen later than the sample is made: the sample is set up again whenever it changes.
        private void RefreshWobbleSample()
        {
            if (_wobbleSample == null || !_wobbleSample.IsHandleCreated) return;
            _wobbleSample.SetSingleSample(_wobbleUseStyle.Checked ? SelectStyle : SelectStyle.Standard, _wobbleUseStyle.Checked ? ActiveMarker : ActiveMarker.Underline, "サンプル");
            _wobbleSample.SetWobble(Wobble, _wobbleAmp.Value, _wobbleDur.Value);
            _wobbleSample.SetWobbleSound(_wobbleSound.Checked, _wobbleVolume.Value);
        }

        // Plays the sample's wobble, with its sound while the check box "効果音を鳴らす" is checked: every round of the timer sounds as well, so the sound can be heard properly.
        private void PlayWobbleSample()
        {
            if (_wobbleSample != null) _wobbleSample.PlayWobble(true);
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

        // The look of selected folders and sheets: the eight types in four columns and two rows. One "not selected" sample (the same for every type) stands at the left, centered over both rows; each type
        // has two samples (a folder and a plain sheet selected, a colored sheet selected), all drawn by the tab strip itself.
        private GroupBox CreateStyleGroup()
        {
            var group = new GroupBox { Text = "フォルダ＆シート選択のデザイン", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            _styleStandard.Text = "標準 (Excelの横タブと同じ見た目)";
            _styleDouble.Text = "ダブルフレーム (二重の枠で、目立たせる)";
            _styleGlossy.Text = "グロッシーグラス (光沢のあるガラス)";
            _styleFrosted.Text = "フロストグラス１ (縁が光る磨りガラス)";
            _styleFrosted2.Text = "フロストグラス２ (行の色を帯びた磨りガラス)";
            _styleInset.Text = "インセット (押し込まれたように凹む)";
            _styleStripes.Text = "ストライプ (ななめの縞で、目立たせる)";
            _styleCrystal.Text = "クリスタル (斜めのカット面で、明るい色が消えていく)";
            var panel = new TableLayoutPanel
            {
                ColumnCount = 5,
                RowCount = 2,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = Padding.Empty
            };
            var none = CreateNoneBlock();
            none.Anchor = AnchorStyles.None;
            panel.Controls.Add(none, 0, 0);
            panel.SetRowSpan(none, 2);
            panel.Controls.Add(CreateStyleBlock(_styleStandard, SelectStyle.Standard), 1, 0);
            panel.Controls.Add(CreateStyleBlock(_styleDouble, SelectStyle.DoubleFrame), 2, 0);
            panel.Controls.Add(CreateStyleBlock(_styleGlossy, SelectStyle.GlossyGlass), 3, 0);
            panel.Controls.Add(CreateStyleBlock(_styleStripes, SelectStyle.Stripes), 4, 0);
            panel.Controls.Add(CreateStyleBlock(_styleFrosted, SelectStyle.FrostedGlass), 1, 1);
            panel.Controls.Add(CreateStyleBlock(_styleFrosted2, SelectStyle.FrostedGlass2), 2, 1);
            panel.Controls.Add(CreateStyleBlock(_styleInset, SelectStyle.Inset), 3, 1);
            panel.Controls.Add(CreateStyleBlock(_styleCrystal, SelectStyle.Crystal), 4, 1);
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        // The "not selected" sample, shown once at the left of the grid and centered vertically over both rows; it looks the same for every type.
        private FlowLayoutPanel CreateNoneBlock()
        {
            var block = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Margin = new Padding(0, 0, 8, 8)
            };
            var strip = new TabStrip();
            strip.SetSample(SelectStyle.Standard, "none", ActiveMarker);
            _samples.Add(strip);
            strip.Size = strip.SampleSize(120);
            strip.Margin = new Padding(0);
            var caption = new Label { Text = "非選択", AutoSize = true, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 0, 0, 2) };
            var one = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(16, 0, 0, 0) };
            one.Controls.Add(caption);
            one.Controls.Add(strip);
            block.Controls.Add(one);
            return block;
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
                Text = "横タブのシート名が「A | B」のとき、縦タブでは、フォルダ A の中のシート B として表示します。\r\n縦タブでの、フォルダの作成、移動、名前の変更も、横タブのシート名に反映します。",
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
