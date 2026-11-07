using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace VerticalSheetTabs
{
    // "Tab display settings" dialog: display mode and vertical tab side.
    // Opened from the ribbon button and, in Step5, from the tab context menu.
    public class SettingsForm : Form
    {
        private readonly RadioButton _modeHorizontal = new RadioButton();
        private readonly RadioButton _modeBoth = new RadioButton();
        private readonly RadioButton _sideLeft = new RadioButton();
        private readonly RadioButton _sideRight = new RadioButton();
        private GroupBox _sideGroup;
        // The first width of the vertical tabs: by characters or by pixels; each way keeps its own value.
        private readonly RadioButton _widthChars = new RadioButton();
        private readonly RadioButton _widthPx = new RadioButton();
        private readonly TrackBar _widthCharsBar = new TrackBar();
        private readonly NumericUpDown _widthCharsBox = new NumericUpDown();
        private readonly TrackBar _widthPxBar = new TrackBar();
        private readonly NumericUpDown _widthPxBox = new NumericUpDown();
        private GroupBox _widthGroup;
        // How the rows are laid out: two radio buttons, each with a preview drawn by the tab strip itself.
        private readonly RadioButton _layoutHanging = new RadioButton();
        private readonly RadioButton _layoutStacked = new RadioButton();
        private GroupBox _layoutGroup;
        private readonly List<TabStrip> _layoutPreviews = new List<TabStrip>();
        private readonly BackgroundPanel _background;
        private GroupBox _jumpGroup; // the position of the comment jump: only used while the note / comment badge is shown
        private readonly RadioButton _jumpCenter = new RadioButton();
        private readonly RadioButton _jumpTopLeft = new RadioButton();
        private readonly RadioButton _jumpTopCenter = new RadioButton();
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
        private readonly CheckBox _badgeError = new CheckBox();
        private readonly CheckBox _badgeMemo = new CheckBox();
        private readonly CheckBox _badgeProtect = new CheckBox();
        private GroupBox _badgeGroup;
        private readonly CheckBox _rowAnim = new CheckBox();
        private readonly CheckBox _doneAnim = new CheckBox();
        private readonly CheckBox _badgeProgress = new CheckBox();
        private GroupBox _animGroup;
        // The groups in four tabs; the sample of the animation only moves while its tab is shown.
        private readonly TabControl _tabs = new TabControl();
        private TabPage _animPage;
        private DateTime _wobblePlayed = DateTime.MinValue;
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

        public SettingsForm(bool tabsHidden, TabMode mode, TabSide side, bool outlineVisible, int outlineLevel, bool nameRule, SelectStyle selectStyle, ActiveMarker activeMarker, WobbleType wobble, int wobbleAmp, int wobbleDur, bool wobbleSound, int wobbleVolume, bool wobbleSampleStyle, BadgeKinds badges, bool rowAnim, bool doneAnim, CommentJump commentJump = CommentJump.TopLeft, WallpaperOptions wallpaper = null, int rowTransparency = 0, bool textOutline = true, InitialWidthMode initialWidthMode = InitialWidthMode.Chars, int initialWidthChars = 12, int initialWidthPx = 0, RowLayout rowLayout = RowLayout.Hanging)
        {
            _initialPx = initialWidthPx;
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
            _modeBoth.Text = "横タブと縦タブの両方を表示";
            _sideLeft.Text = "左";
            _sideRight.Text = "右";

            var modeGroup = CreateRowGroup("表示モード", _modeHorizontal, _modeBoth);
            _sideGroup = CreateRowGroup("縦タブの位置", _sideLeft, _sideRight);
            _jumpTopCenter.Text = "中央上";
            _jumpCenter.Text = "行・列の中央";
            _jumpTopLeft.Text = "左上 (標準)";
            _jumpGroup = CreateRowGroup("コメントの「そのセルへ移動」で、セルを表示する位置", true, _jumpTopLeft, _jumpTopCenter, _jumpCenter);

            _outlineGroup = CreateOutlineGroup();
            _ruleGroup = CreateRuleGroup();
            _background = new BackgroundPanel(wallpaper, rowTransparency, textOutline);
            _rowAnim.Text = "フォルダ・シートの追加・削除を、動きで見せる";
            _rowAnim.Checked = rowAnim;
            _doneAnim.Text = "進捗を「完了」にしたとき、印をはじけさせる";
            _doneAnim.Checked = doneAnim;
            _badgeProgress.Text = "進捗 (未着手・作業中・完了)";
            _badgeProgress.Checked = (badges & BadgeKinds.Progress) != 0;
            _animGroup = CreateAnimGroup();
            _badgeError.Text = "数式エラー";
            _badgeMemo.Text = "メモ・コメント";
            _badgeProtect.Text = "シートの保護";
            _badgeError.Checked = (badges & BadgeKinds.Error) != 0;
            _badgeMemo.Checked = (badges & BadgeKinds.Memo) != 0;
            _badgeProtect.Checked = (badges & BadgeKinds.Protect) != 0;
            _badgeGroup = CreateBadgeGroup("表示するバッジ (閉じたフォルダには、中のシートの数を表示)", new[] { BadgeKinds.Error, BadgeKinds.Memo, BadgeKinds.Protect, BadgeKinds.Progress }, new[] { _badgeError, _badgeMemo, _badgeProtect, _badgeProgress });
            foreach (CheckBox c in new[] { _badgeError, _badgeMemo, _badgeProtect, _badgeProgress })
            {
                c.CheckedChanged += (s, e) => { foreach (TabStrip sample in _samples) sample.ShownBadges = Badges; };
            }
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
            _wobbleGroup = CreateWobbleGroup(wobble, wobbleAmp, wobbleDur, wobbleSound, wobbleVolume, wobbleSampleStyle);
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

            // The groups are sorted into four tabs; the form always opens on the first one, "基本". Every group is as wide
            // as the widest one (the looks of selected rows).
            int width = _styleGroup.Width;
            // 基本: the display mode and the side of the vertical tabs side by side (the side takes about a third of the width), and the sheet-name rule
            modeGroup.Dock = DockStyle.None;
            _sideGroup.Dock = DockStyle.None;
            int topHeight = Math.Max(modeGroup.Height, _sideGroup.Height);
            // 基本 has two columns: the mode, the side and the first width on the left, the layout of the rows (with previews) on the right
            int dpiNow = DpiHelper.GetDpi(IntPtr.Zero);
            int previewW = 190 * dpiNow / 96;
            int rightW = 2 * previewW + 3 * 12 + 24;
            int leftW = width - rightW - 12;
            int sideWidth = (leftW - 12) * 30 / 100;
            modeGroup.Size = new Size(leftW - sideWidth - 6 - 6, topHeight);
            _sideGroup.Size = new Size(sideWidth, topHeight);
            _widthGroup = CreateWidthGroup(leftW - 6, initialWidthMode, initialWidthChars, initialWidthPx);
            _layoutGroup = CreateLayoutGroup(rightW, previewW, rowLayout, selectStyle, activeMarker);
            _jumpGroup.Dock = DockStyle.None;
            _jumpGroup.Size = new Size(width, _jumpGroup.Height);
            _jumpGroup.Margin = new Padding(3);
            modeGroup.Margin = new Padding(3);
            _sideGroup.Margin = new Padding(3);
            var top = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            top.Controls.Add(modeGroup);
            top.Controls.Add(_sideGroup);
            _tabs.Padding = new Point(16, 5); // (room around the titles of the tabs)
            // (the tabs other than "このブック" are kept for the whole PC)
            string pcNote = "このタブの設定は、このPCのすべてのブックに共通です。";
            _tabs.TabPages.Add(CreateBasicPage(width, leftW, pcNote, top, tabsHidden));
            _tabs.TabPages.Add(CreatePage("背景", width, Banner(pcNote, width), _background.WallpaperGroup, _background.PairRow(width), _background.TransparencyGroup));
            _tabs.TabPages.Add(CreatePage("デザイン", width, Banner(pcNote, width), _markerGroup, _styleGroup, _outlineGroup));
            _animPage = CreatePage("アニメーション", width, Banner(pcNote, width), _wobbleGroup, _animGroup);
            _tabs.TabPages.Add(_animPage);
            _tabs.TabPages.Add(CreatePage("バッジ", width, Banner(pcNote, width), _badgeGroup, _jumpGroup));
            _tabs.TabPages.Add(CreatePage("このブック", width, Banner("このタブの設定は、開いているブックだけに保存されます。", width), _ruleGroup));
            // the tabs are as large as the largest page
            var inner = Size.Empty;
            foreach (TabPage page in _tabs.TabPages)
            {
                Size s = page.Controls[0].PreferredSize;
                inner = new Size(Math.Max(inner.Width, s.Width + page.Padding.Horizontal), Math.Max(inner.Height, s.Height + page.Padding.Vertical));
            }
            IntPtr created = _tabs.Handle; // (the size of the tab headers is known once the control exists)
            _tabs.Size = new Size(1000, 1000);
            Rectangle shown = _tabs.DisplayRectangle;
            _tabs.Size = new Size(inner.Width + 1000 - shown.Width, inner.Height + 1000 - shown.Height);
            _tabs.Margin = new Padding(3);
            // the right column of the basic page is as high as the tabs: the previews take what the left column does not use
            if (_basicLayout != null)
            {
                TabPage basicPage = _tabs.TabPages[0];
                _basicLayout.AutoSize = false;
                _basicLayout.Size = new Size(inner.Width - basicPage.Padding.Horizontal, inner.Height - basicPage.Padding.Vertical);
            }
            // the sample of the animation moves (and sounds) while its tab is shown; it plays at once when the tab is chosen
            _tabs.SelectedIndexChanged += (s, e) =>
            {
                if (_tabs.SelectedTab != _animPage) return;
                RefreshWobbleSample();
                BeginInvoke(new Action(PlayWobbleSample));
                if (_wobbleTimer != null) { _wobbleTimer.Stop(); _wobbleTimer.Start(); }
            };

            var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            layout.Controls.Add(_tabs);
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
            _modeBoth.Checked = mode == TabMode.Both;
            _sideLeft.Checked = side == TabSide.Left;
            _sideRight.Checked = side == TabSide.Right;
            _jumpCenter.Checked = commentJump == CommentJump.Center;
            _jumpTopCenter.Checked = commentJump == CommentJump.TopCenter;
            _jumpTopLeft.Checked = commentJump != CommentJump.Center && commentJump != CommentJump.TopCenter;

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
                    foreach (TabStrip pv in _layoutPreviews) pv.SelectionStyle = SelectStyle;
                };
            }
            _outlineShow.Checked = outlineVisible;
            _outlineLevel.Value = Math.Max(1, Math.Min(TabStrip.OutlineLevels, outlineLevel));
            _outlineShow.CheckedChanged += (s, e) => UpdateOutline();
            _outlineLevel.ValueChanged += (s, e) => UpdateOutline();
            UpdateOutline();

            _badgeMemo.CheckedChanged += (s, e) => UpdateJumpEnabled();
            UpdateJumpEnabled();
            _modeHorizontal.CheckedChanged += (s, e) => UpdateSideEnabled();
            UpdateSideEnabled();
        }

        public TabMode Mode
        {
            get
            {
                return _modeHorizontal.Checked ? TabMode.HorizontalOnly : TabMode.Both;
            }
        }

        private int _initialPx;
        private int _minPx = 0;

        public InitialWidthMode InitialWidthMode { get { return _widthPx.Checked ? InitialWidthMode.Px : InitialWidthMode.Chars; } }

        public RowLayout RowLayout { get { return _layoutStacked.Checked ? RowLayout.Stacked : RowLayout.Hanging; } }

        public int InitialWidthChars { get { return (int)_widthCharsBox.Value; } }

        // 0 stays 0 while the value is untouched at its minimum (the minimum depends on the screen's scale).
        public int InitialWidthPx { get { int v = (int)_widthPxBox.Value; return _initialPx == 0 && v == _minPx ? 0 : v; } }

        public TabSide Side
        {
            get { return _sideRight.Checked ? TabSide.Right : TabSide.Left; }
        }

        public CommentJump CommentJump
        {
            get { return _jumpCenter.Checked ? CommentJump.Center : _jumpTopCenter.Checked ? CommentJump.TopCenter : CommentJump.TopLeft; }
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

        public bool RowAnim { get { return _rowAnim.Checked; } }

        public WallpaperOptions Wallpaper { get { return _background.Wallpaper; } }

        public int RowTransparency { get { return _background.RowTransparency; } }

        public bool TextOutline { get { return _background.TextOutline; } }

        public bool DoneAnim { get { return _doneAnim.Checked; } }

        public BadgeKinds Badges
        {
            get { return (_badgeError.Checked ? BadgeKinds.Error : BadgeKinds.None) | (_badgeMemo.Checked ? BadgeKinds.Memo : BadgeKinds.None) | (_badgeProtect.Checked ? BadgeKinds.Protect : BadgeKinds.None) | (_badgeProgress.Checked ? BadgeKinds.Progress : BadgeKinds.None); }
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

        public bool WobbleSampleStyle { get { return _wobbleUseStyle.Checked; } }

        public int WobbleVolume { get { return _wobbleVolume.Value; } }

        public int WobbleDur { get { return _wobbleDur.Value; } }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (_wobbleTimer != null) _wobbleTimer.Stop();
            base.OnFormClosed(e);
        }

        // The wobble of the selected sheet: the motion (none, 1 to 6, random), its size and length, and one live sample that
        // plays the chosen motion again and again (a new random one each time for "7 ランダム"). The sample follows the chosen look and marker.
        private GroupBox CreateWobbleGroup(WobbleType wobble, int amp, int dur, bool sound, int volume, bool sampleStyle)
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
            _wobbleUseStyle.Checked = sampleStyle;
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
        // Only while the tab of the animation is shown (no sound from a tab that cannot be seen), and not twice at once (the tab chosen and the
        // sample made at the same moment).
        private void PlayWobbleSample()
        {
            if (_wobbleSample == null || !_wobbleSample.Visible) return;
            if ((DateTime.Now - _wobblePlayed).TotalMilliseconds < 300) return;
            _wobblePlayed = DateTime.Now;
            _wobbleSample.PlayWobble(true);
        }

        // A line of text at the top of a tab, as wide as the groups.
        private static Label Banner(string text, int width)
        {
            return new Label { Text = text, AutoSize = true, MaximumSize = new Size(width, 0), Margin = new Padding(3, 3, 3, 6) };
        }

        // A tab with the groups one below the other, each as wide as the widest group of the form.
        private static TabPage CreatePage(string title, int width, params Control[] rows)
        {
            var page = new TabPage(title) { Padding = new Padding(6), UseVisualStyleBackColor = true };
            var layout = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Location = new Point(6, 6), Margin = Padding.Empty };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            for (int i = 0; i < rows.Length; i++)
            {
                Control row = rows[i];
                if (row is GroupBox)
                {
                    row.Dock = DockStyle.None;
                    row.Width = width;
                    row.Margin = new Padding(3, i == 0 ? 3 : 8, 3, 3);
                }
                layout.Controls.Add(row);
            }
            page.Controls.Add(layout);
            return page;
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
            strip.ShownBadges = Badges;
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
            strip.ShownBadges = Badges;
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
        // The jump is started from the badge: with the badge off its position cannot be chosen (the value is kept).
        private void UpdateJumpEnabled()
        {
            _jumpGroup.Enabled = _badgeMemo.Checked;
        }

        private void UpdateSideEnabled()
        {
            _sideGroup.Enabled = !_modeHorizontal.Checked;
            if (_widthGroup != null) _widthGroup.Enabled = !_modeHorizontal.Checked;
            if (_layoutGroup != null) _layoutGroup.Enabled = !_modeHorizontal.Checked;
        }

        // 縦タブの初回幅: a radio button, a slider, a number box and a unit on each of two lines (characters, pixels).
        private GroupBox CreateWidthGroup(int width, InitialWidthMode mode, int chars, int px)
        {
            var group = new GroupBox { Text = "縦タブの初回幅", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6) };
            int dpi = DpiHelper.GetDpi(IntPtr.Zero);
            int minPx = TabStrip.ContentWidthForChars(12, dpi);
            _minPx = minPx;
            _widthChars.Text = "文字数";
            _widthPx.Text = "px";
            _widthChars.AutoSize = false; _widthPx.AutoSize = false;
            // (the height of a slider is its own unless AutoSize is off: with visual styles it grew and covered the line below)
            _widthCharsBar.AutoSize = false; _widthPxBar.AutoSize = false;
            _widthChars.Checked = mode != InitialWidthMode.Px;
            _widthPx.Checked = mode == InitialWidthMode.Px;
            _widthCharsBar.Minimum = 12; _widthCharsBar.Maximum = 64; _widthCharsBar.TickFrequency = 4; _widthCharsBar.SmallChange = 1; _widthCharsBar.LargeChange = 4; _widthCharsBar.Value = Math.Max(12, Math.Min(64, chars));
            _widthPxBar.Minimum = minPx; _widthPxBar.Maximum = 2000; _widthPxBar.TickFrequency = 100; _widthPxBar.SmallChange = 1; _widthPxBar.LargeChange = 50; _widthPxBar.Value = Math.Max(minPx, Math.Min(2000, px > 0 ? px : minPx));
            _widthCharsBox.Minimum = 12; _widthCharsBox.Maximum = 64; _widthCharsBox.Value = _widthCharsBar.Value;
            _widthPxBox.Minimum = minPx; _widthPxBox.Maximum = 2000; _widthPxBox.Value = _widthPxBar.Value;
            _widthCharsBar.ValueChanged += (s, e) => _widthCharsBox.Value = _widthCharsBar.Value;
            _widthCharsBox.ValueChanged += (s, e) => _widthCharsBar.Value = (int)_widthCharsBox.Value;
            _widthPxBar.ValueChanged += (s, e) => _widthPxBox.Value = _widthPxBar.Value;
            _widthPxBox.ValueChanged += (s, e) => _widthPxBar.Value = (int)_widthPxBox.Value;
            _widthChars.CheckedChanged += (s, e) => UpdateWidthEnabled();
            UpdateWidthEnabled();
            int rowH = 38;
            int boxW = TextRenderer.MeasureText("20000", group.Font).Width + SystemInformation.VerticalScrollBarWidth + 10;
            int unitW = TextRenderer.MeasureText("文字分", group.Font).Width + 8;
            int radioW = TextRenderer.MeasureText("文字数", group.Font).Width + 30;
            int inner = width - group.Padding.Horizontal - 12;
            int barW = inner - radioW - boxW - unitW - 12;
            Control[][] lines =
            {
                new Control[] { _widthChars, _widthCharsBar, _widthCharsBox, new Label { Text = "文字分" } },
                new Control[] { _widthPx, _widthPxBar, _widthPxBox, new Label { Text = "px" } },
            };
            var caption = new Label { Text = "新しいブックを初めて開いたときの縦タブ幅の設定です。", AutoSize = false, ForeColor = SystemColors.GrayText };
            caption.SetBounds(8, 20, inner, 20);
            group.Controls.Add(caption);
            int y = 20 + 22;
            foreach (Control[] line in lines)
            {
                int x = 8;
                line[0].SetBounds(x, y + 4, radioW, 24); x += radioW;
                line[1].SetBounds(x, y, barW, rowH - 6); x += barW + 4;
                line[2].SetBounds(x, y + 4, boxW, 24); x += boxW + 4;
                line[3].SetBounds(x, y + 8, unitW, 20);
                ((Label)line[3]).AutoSize = false;
                foreach (Control c in line) group.Controls.Add(c);
                y += rowH;
            }
            var hint = new Label
            {
                Text = "・文字数は、半角で 12 ～ 64 文字。\r\n・px は、半角12文字分 (" + minPx + " px) から 2000 px まで。",
                AutoSize = false, ForeColor = SystemColors.GrayText
            };
            hint.SetBounds(8, y + 2, inner, TextRenderer.MeasureText(hint.Text, group.Font, new Size(inner, 0), TextFormatFlags.WordBreak).Height + 2);
            group.Controls.Add(hint);
            group.Size = new Size(width, hint.Bottom + group.Padding.Bottom + 4);
            return group;
        }

        private TableLayoutPanel _basicLayout;

        // 基本: two columns. Left: the note, the mode and the side, the first width. Right: the layout of the rows.
        private TabPage CreateBasicPage(int width, int leftW, string note, Control top, bool tabsHidden)
        {
            var page = new TabPage("基本") { Padding = new Padding(6), UseVisualStyleBackColor = true };
            var left = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = Padding.Empty };
            left.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            Label banner = Banner(note, leftW);
            left.Controls.Add(banner);
            left.Controls.Add(top);
            _widthGroup.Dock = DockStyle.None;
            _widthGroup.Margin = new Padding(3, 8, 3, 3);
            left.Controls.Add(_widthGroup);
            if (tabsHidden) left.Controls.Add(Banner("このブックは、横タブが非表示のため、縦タブは表示されません。", leftW));
            // the group on the right begins at the height of the mode and the side
            _layoutGroup.Margin = new Padding(3, banner.PreferredSize.Height + banner.Margin.Vertical + 8, 3, 3);
            _basicLayout = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Location = new Point(6, 6), Margin = Padding.Empty };
            _basicLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _basicLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _basicLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            left.Dock = DockStyle.Top;
            _basicLayout.Controls.Add(left, 0, 0);
            _basicLayout.Controls.Add(_layoutGroup, 1, 0);
            page.Controls.Add(_basicLayout);
            return page;
        }

        // フォルダ＆シートの表示方法: the two ways, each with a preview that the tab strip draws (a tool row, a band, rows, the "+").
        private GroupBox CreateLayoutGroup(int groupWidth, int previewW, RowLayout layout, SelectStyle style, ActiveMarker marker)
        {
            var group = new GroupBox { Text = "フォルダ＆シートの表示方法", Font = SystemFonts.MessageBoxFont, AutoSize = false, Padding = new Padding(8, 4, 8, 6), Dock = DockStyle.Fill };
            _layoutHanging.Text = "吊り下げ方式 (標準)";
            _layoutStacked.Text = "積み上げ方式";
            _layoutHanging.AutoSize = true; _layoutStacked.AutoSize = true;
            _layoutHanging.Checked = layout != RowLayout.Stacked;
            _layoutStacked.Checked = layout == RowLayout.Stacked;
            var table = new TableLayoutPanel { ColumnCount = 2, RowCount = 2, Dock = DockStyle.Fill, Margin = Padding.Empty };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            table.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            table.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            RadioButton[] radios = { _layoutHanging, _layoutStacked };
            RowLayout[] kinds = { RowLayout.Hanging, RowLayout.Stacked };
            for (int i = 0; i < 2; i++)
            {
                RadioButton radio = radios[i];
                radio.Margin = new Padding(12, 6, 3, 6);
                table.Controls.Add(radio, i, 0);
                var strip = new TabStrip { Dock = DockStyle.Fill, Margin = new Padding(12, 0, 12, 6), MinimumSize = new Size(previewW, 0) };
                strip.SetLayoutPreview(kinds[i], style, marker);
                strip.SampleClicked += () => radio.Checked = true;
                _layoutPreviews.Add(strip);
                _samples.Add(strip);
                table.Controls.Add(strip, i, 1);
            }
            group.Controls.Add(table);
            // (as high as it needs to show the preview with the tool row, the band, four rows and the "+")
            int need = (38 + 24 + 4 * 32 + 32 + 40) * DpiHelper.GetDpi(IntPtr.Zero) / 96;
            group.MinimumSize = new Size(groupWidth, need + 40);
            group.Size = new Size(groupWidth, need + 40);
            return group;
        }

        private void UpdateWidthEnabled()
        {
            _widthCharsBar.Enabled = _widthCharsBox.Enabled = _widthChars.Checked;
            _widthPxBar.Enabled = _widthPxBox.Enabled = _widthPx.Checked;
        }

        // A group box with its radio buttons side by side, at equal steps: the widest one plus a gap of about
        // four half-width characters.
        // The group of the animations: the add and delete animation of the rows (with a note), and the burst of a progress mark.
        private GroupBox CreateAnimGroup()
        {
            var group = new GroupBox
            {
                Text = "追加・削除・完了のアニメーション",
                Font = SystemFonts.MessageBoxFont,
                AutoSize = false,
                Padding = new Padding(8, 4, 8, 6)
            };
            var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            var row1 = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            _rowAnim.AutoSize = true;
            _rowAnim.Font = group.Font;
            _rowAnim.Margin = new Padding(3, 3, 12, 3);
            var note = new Label
            {
                Text = "(追加: ぽんと現れる / 削除: 縮んで消える)",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Font = group.Font,
                Margin = new Padding(0, 6, 3, 3)
            };
            row1.Controls.Add(_rowAnim);
            row1.Controls.Add(note);
            _doneAnim.AutoSize = true;
            _doneAnim.Font = group.Font;
            _doneAnim.Margin = new Padding(3, 3, 3, 3);
            panel.Controls.Add(row1);
            panel.Controls.Add(_doneAnim);
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        // The icon of a badge, drawn as it is on the rows.
        private class BadgeIcon : Control
        {
            private readonly BadgeKinds _kind;

            public BadgeIcon(BadgeKinds kind, int size)
            {
                _kind = kind;
                SetStyle(ControlStyles.SupportsTransparentBackColor | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint, true);
                BackColor = Color.Transparent;
                Size = new Size(size, size);
                TabStop = false;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                TabStrip.DrawBadgeIcon(e.Graphics, _kind, new Rectangle(1, 1, Width - 2, Height - 2));
            }
        }

        // Like a row of radio buttons (CreateRowGroup), but each check box has the icon of its badge at its left; three in the first row, the rest below.
        private static GroupBox CreateBadgeGroup(string title, BadgeKinds[] kinds, CheckBox[] boxes)
        {
            var group = new GroupBox
            {
                Text = title,
                Font = SystemFonts.MessageBoxFont,
                AutoSize = false,
                Padding = new Padding(8, 4, 8, 6)
            };
            var panel = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
            int widest = 0;
            foreach (CheckBox c in boxes)
            {
                c.AutoSize = true;
                c.Font = group.Font;
                widest = Math.Max(widest, c.PreferredSize.Width);
            }
            int icon = (int)(boxes[0].PreferredSize.Height * 1.1);
            int gap = TextRenderer.MeasureText("0000", group.Font).Width;
            const int perRow = 3;
            FlowLayoutPanel line = null;
            for (int i = 0; i < boxes.Length; i++)
            {
                if (i % perRow == 0)
                {
                    line = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = Padding.Empty };
                    panel.Controls.Add(line);
                }
                var item = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(3, 0, widest - boxes[i].PreferredSize.Width + gap, 0) };
                // the icon and the check box stand on the same middle line, a few pixels apart
                int rowH = Math.Max(icon, boxes[i].PreferredSize.Height) + 4;
                var pic = new BadgeIcon(kinds[i], icon) { Margin = new Padding(0, (rowH - icon) / 2, 6, (rowH - icon) - (rowH - icon) / 2) };
                int bh = boxes[i].PreferredSize.Height;
                boxes[i].Margin = new Padding(0, (rowH - bh) / 2, 0, (rowH - bh) - (rowH - bh) / 2);
                item.Controls.Add(pic);
                item.Controls.Add(boxes[i]);
                line.Controls.Add(item);
            }
            group.Controls.Add(panel);
            panel.Location = group.DisplayRectangle.Location;
            Size content = panel.PreferredSize;
            group.Size = new Size(content.Width + group.Padding.Horizontal + 8, panel.Top + content.Height + group.Padding.Bottom + 2);
            return group;
        }

        private static GroupBox CreateRowGroup(string title, params ButtonBase[] radios)
        {
            return CreateRowGroup(title, false, radios);
        }

        // equalGaps: the space between the texts is the same (otherwise every button is as wide as the widest one)
        private static GroupBox CreateRowGroup(string title, bool equalGaps, params ButtonBase[] radios)
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
            foreach (ButtonBase r in radios)
            {
                r.AutoSize = true;
                r.Font = group.Font;
                widest = Math.Max(widest, r.PreferredSize.Width);
            }
            int gap = TextRenderer.MeasureText("0000", group.Font).Width;
            foreach (ButtonBase r in radios)
            {
                r.Margin = new Padding(3, 3, equalGaps ? gap : widest - r.PreferredSize.Width + gap, 3);
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
