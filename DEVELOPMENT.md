# 開発者向けの手引き (VerticalSheetTabs)

利用者向けの説明は [README.md](README.md) (配布用zipにも入ります)、現在の仕様、実装、守ること、リリースの決まりは [Agents.md](Agents.md) を参照してください。

初めてソースコードを読む人は、次の順に読むと、全体から細部へ進めます。

1. [全体像](#全体像): 何をするアドインで、どんな仕組みか。まず、ここだけ読めば足ります
2. [ソースファイルの構成](#ソースファイルの構成): どのファイルが、何を受け持つか
3. [設計図 (UML)](#設計図-uml): 図で、何ができるか > どんな部品か > どんなクラスか > どう動くか、の順に見る
4. [ビルド](#ビルド) 以降: 動かす、テストする、リリースする

## 全体像

Windows版 Excel のシートタブを、縦一覧で表示する VSTO アドインです (.NET Framework 4.8、C# 7.3)。縦タブは、Excel のタスクペインの中に、自前で描いています (GDI+)。

```mermaid
flowchart LR
    Excel[Excel] -- イベント --> T[ThisAddIn]
    T --> P[PaneManager]
    P -- ウィンドウごとに1つ --> C[TabPaneControl]
    C --> S[TabStrip<br/>自前描画]
    P <-- 読み書き --> W[(ブック<br/>VTab.Width・NameRule / カスタムXML / VTab.Id)]
    P <-- 読み書き --> F[(PCの設定ファイル<br/>settings.ini)]
    P -- 0.3〜3秒おきに同期 --> Excel
    S -- 操作 --> P
```

最初に押さえる、設計の決まり:

- **Excel のオブジェクトモデルを呼ぶのは `PaneManager`** です (ブックの読み書きの部品 `FolderStore`、`HistoryStore`、`WorkbookSettings` も、`PaneManager` から呼ばれて、渡されたブックを読み書きします)。`TabStrip` などの画面の部品は Excel を知らず、操作を「イベント」で `PaneManager` に伝え、表示は `PaneManager` から渡されます
- Excel には、シート名の変更、並べ替え、非表示のイベントが無いため、**0.3から3秒おきのタイマーで同期**します (`PaneManager`)
- **開いただけ、同期しただけでは、ブックを「変更あり」にしません**。ユーザーの操作のときだけ変更扱いにします
- 設定は、**PC** (`settings.ini`) と **ブック** (`VTab.Width`、`VTab.NameRule`、フォルダのXML) に分けて保存します
- フォルダの構造、履歴、シート名のルールは、**Excel に依存しないモデル**です (`FolderModel`、`SheetHistory`、`NameRule`。ブックへの読み書きは、別の部品 `FolderStore`、`HistoryStore` が受け持ちます)。そのため、Excel なしでテストできます (`NameRule`、`FolderModel`、`SafeXml` のテストがあります)

これらの理由と、細かい決まりは、[Agents.md](Agents.md) の3章と4章にあります。

## ソースファイルの構成

| ファイル | 役割 |
|---|---|
| `ThisAddIn.cs` | アドインの入口。Excelのイベント (ウィンドウを有効にした、シートを有効にした、ブックの保存前。バッジを使うときだけ、セルの変更、再計算、シートを離れた) を `PaneManager` につなぐ。すべてのイベント処理は、例外を握りつぶすガードで囲む |
| `PaneManager.cs` | ウィンドウごとのタスクペインの管理。0.3から3秒おきの同期 (シート名、表示状態、タブ色、選択)、シートの操作、フォルダの操作、設定の適用 |
| `TabPaneControl.cs` | ペインの中身 (`TabStrip`、「+」ボタン、右クリックメニュー) |
| `TabStrip.cs` | 縦タブの自前描画 (GDI+)。選択、ドラッグ、名前の変更、スクロール、フォルダの描画 |
| `TabTheme.cs`、`DpiHelper.cs` | Officeテーマの配色、画面の拡大率 |
| `SheetHistory.cs` | 見たシートの履歴 (`SheetHistory`) とそのXML、ブックのXMLパートへの読み書き (`HistoryStore`) |
| `Wobble.cs`、`WobbleSounds.cs`、`Sounds\*.wav` | 揺れの動き、効果音 (アドインに埋め込む) |
| `DialogPlacer.cs` | Excelの「色の設定」画面を、クリックした位置に出す |
| `TabColorPalette.cs` | シート見出しの色のパレット (Excelと同じ構成) |
| `SafeXml.cs` | ブックから読むXML (カスタムXMLパート、ファイルのテーマ) の読み込み。DTD禁止、大きさの上限 |
| `FolderModel.cs`、`FolderStore.cs` | フォルダの構造、XMLの読み書き、Excelのシート順との突き合わせ (`Reconcile`)、シートの識別 |
| `NameRule.cs` | シート名のルール (`フォルダ名 \| シート名`) の解析、並びの整理、名前からのフォルダ構成 |
| `WorkbookSettings.cs` | 1つのウィンドウの設定。ブックのカスタムドキュメントプロパティ (`VTab.Width`、`VTab.NameRule`) の読み書きと、古い `VTab.*` の削除 |
| `PcSettings.cs` | PCの設定ファイル (`%LocalAppData%\VerticalSheetTabs\settings.ini`。表示モード、位置、見た目、アニメーション、バッジ) の読み書き |
| `CommentPopup.cs` | メモ・コメントのバッジの表示 (一覧と内容の窓) |
| `SettingsForm.cs`、`RibbonController.cs`、`RibbonIcon.cs` | 設定画面、リボンのボタンと横タブのメニュー項目、アイコン |
| `AddInVersion.cs` | 表示用のバージョン、著作権、ライセンス |
| `VerticalSheetTabs.Tests\` | Excel を使わない部分 (`NameRule`、`FolderModel`、`SafeXml`) のテスト。ソースを参照で取り込む、フレームワーク不要のコンソールプログラム。`tools\run-tests.ps1` で実行 |
| `tools\dist\` | 配布用zipに入れる、証明書を信頼する / 取り消すスクリプトと、手順書の雛形 |
| `tools\make-dist-zip.ps1`、`tools\make-readme-pdf.ps1` | 配布用zipの作成、README.pdf の作成 |
| `.github\workflows\release.yml` | タグの push で、配布用zipを添付した GitHub リリースを作る |

## 設計図 (UML)

図は PlantUML で描き、SVG にして置きます (GitHub は PlantUML を表示しないため)。ソースは `docs/*.puml`、SVG は同じ名前で `docs/` にあります。

| 知りたいこと | 見る図 |
|---|---|
| 利用者が、何をできるか | [ユースケース図](#ユースケース図) |
| どんな部品でできていて、どうつながるか | [コンポーネント図](#コンポーネント図)、[インターフェースのつながり](#コンポーネント図-インターフェースのつながり) |
| どんなクラスがあるか | [クラス図](#クラス図) |
| どんなデータが、どこからどこへ流れ、どこに保存されるか | [データフロー図](#データフロー図) |
| 動き出しから終わりまでの基本の流れ | [シーケンス図: 基本の流れ](#基本の流れ) |
| シート、フォルダ、検索などの操作が、どう処理されるか | [シーケンス図: 縦タブの操作](#縦タブの操作) |
| 見た目、設定、バッジ、色、シート名のルール | [シーケンス図: 見た目・設定・状態の表示](#見た目設定状態の表示) |
| アニメーションの状態と、起動の判断 | [状態遷移図、アクティビティ図](#状態遷移図とアクティビティ図-アニメーション) |

目的から探すときは、次も使えます。

| 直したいこと | 先に見る図 |
|---|---|
| シートの選択、追加、移動、名前の変更 | シーケンス図「シートのクリックによる選択」「シートの操作」 |
| フォルダの動き | シーケンス図「フォルダの操作」「フォルダの出し入れ・移動・コピー」 |
| 縦タブが出ない、消える | シーケンス図「表示モードとペインの出し入れ」「ウィンドウとブックの終了処理」 |
| 保存したときの内容 | シーケンス図「ブックの保存」 |
| バッジ、進捗、メモ・コメント | シーケンス図「バッジの検出と進捗」「メモ・コメントのバッジのクリック」 |
| 揺れ、追加・削除、完了の動き | シーケンス図「アニメーション」、状態遷移図、アクティビティ図 |

### ユースケース図

利用者ができることを、仕様 (`Agents.md` の2章) にもとづいて、まとまりごとに示す。ソースは [docs/usecase.puml](docs/usecase.puml)。「シートを操作する」のすべては、Excelのオブジェクトモデルを通して行うため、Excelへの線は省いている。

![ユースケース図](docs/usecase.svg)

### コンポーネント図

役割ごとのコンポーネントと、それを構成するモジュール (ソースファイル) を示す。矢印は「呼ぶ / 使う」。Excelのオブジェクトモデルを呼ぶのは、`PaneManager` が中心で、ブックを読み書きする `FolderStore`、`HistoryStore` (`SheetHistory.cs`)、`WorkbookSettings` も、`PaneManager` から渡されたブックを読み書きする (`TabStrip` などの画面の部品は、Excelに触れない)。

![コンポーネント図](docs/component.svg)

### コンポーネント図: インターフェースのつながり

中のモジュールは描かず、役割のコンポーネントの間の、インターフェースのつながりだけを示す。丸が「提供」、受け口が「利用」。ソースは [docs/component-interface.puml](docs/component-interface.puml)。

![インターフェースのつながり](docs/component-interface.svg)

| 番号 | インターフェース | 提供 | 利用 | 中身 |
|---|---|---|---|---|
| I1 | Excelイベント | Excel | 入口 | ウィンドウの有効化、シートの有効化、ブックの保存前のイベント。バッジを使うときだけ、セルの変更、再計算、シートを離れたイベント (`ThisAddIn` が購読) |
| I2 | リボン拡張 | 入口 | Excel | `IRibbonExtensibility`。リボンのボタンと、横タブのメニュー項目の定義、押されたときの通知 |
| I3 | オブジェクトモデル | Excel | 制御 | ウィンドウ、ブック、シートの読み書き (COM)。呼ぶのは制御 (`PaneManager`)。ブックの読み書きの部品 (モデルの `FolderStore`、`HistoryStore`、設定の保持の `WorkbookSettings`) は、制御から渡されたブックを読み書きする |
| I4 | ペイン登録 | Excel | 制御 | `CustomTaskPanes`。ウィンドウごとに縦タブのペインを作る |
| I5 | ペイン管理 | 制御 | 入口 | `EnsurePane`、`EnsureAllPanes`、`ShowSettings`、`SaveSettings`、`SyncSheetOf`、`MarkBadgesDirty`、`Shutdown` |
| I6 | バッジ購読 | 入口 | 制御 | `BadgeEventsWanted` (`Action<bool>`)。バッジを使うウィンドウがあるときだけ、シートの変更・再計算のイベントを購読する |
| I7 | 表示の更新 | 画面 | 制御、設定画面 | 行、バッジ、選択、設定 (見た目、揺れ、アニメーション) を渡す。設定画面のサンプルも、同じ描画を使う |
| I8 | 操作の要求 | 制御 | 画面 | 画面が起こすイベント (`*Requested`: 選択、移動、名前の変更、色、フォルダ、進捗、リンク、メモのバッジなど) を、制御が受ける |
| I9 | 値の問い合わせ | 制御 | 画面 | `Func` で画面が聞く。テーマの色 (`ThemeColorProvider`)、「その他の色」 (`MoreColorsProvider`)、履歴 (`HistoryProvider`)、セルの入力中か (`IsExcelEditing`) |
| I10 | 設定画面 | 設定画面 | 制御 | 現在の値を渡して表示し、OK で値を返す |
| I11 | 動きと音 | 動き・音 | 画面、設定画面、設定の保持 | `WobbleMotion.At` (揺れの動き)、`WobbleSounds.Play` (効果音)、`WobbleType` (種類の値) |
| I12 | モデルの操作 | モデル | 制御、設定画面 | フォルダの構造、XMLの読み書き、シート名のルール、履歴 |
| I13 | 設定の読み書き | 設定の保持 | 制御 | `PcSettings` (`settings.ini`)、`WorkbookSettings` (ブックの `VTab.*`) の読み書き |

### クラス図

主なクラスと、その関係を示す。メンバーは、関係の理解に要るものだけを載せている (すべてではない)。

**主なクラスと関係** ([class-structure.puml](docs/class-structure.puml))。`PaneManager` が、ウィンドウごとの管理情報 (`Entry`) を持ち、その中にペインの中身 (`TabPaneControl`、`TabStrip`) がある。画面は、イベントで操作を `PaneManager` に伝える。

![主なクラスと関係](docs/class-structure.svg)

**データ・モデル・設定** ([class-data.puml](docs/class-data.puml))。フォルダの構造 (`FolderModel`)、履歴、設定 (`WorkbookSettings`。実行中は、PCの設定もここに入る)、画面に渡すデータ (`StripRow`、`BadgeInfo` など)。

![データ・モデル・設定](docs/class-data.svg)

### データフロー図

データが、どこから来て、どこで処理され、どこに保存されるかを示す。四角は外部 (利用者、Excel、Officeテーマ)、楕円はプロセス、円筒はデータの保存先です。処理の順序は表さない (順序は、次のシーケンス図)。

**レベル0 (全体)** ([dfd-level0.puml](docs/dfd-level0.puml))。アドイン全体を1つのプロセスとして、外とのデータのやりとりだけを示す。

![データフロー図 レベル0](docs/dfd-level0.svg)

**レベル1 (プロセスの分解)** ([dfd-level1.puml](docs/dfd-level1.puml))。アドインを8つのプロセスに分け、保存先 (D1からD4) との出入りを示す。括弧内は、主に受け持つクラスやイベント。

![データフロー図 レベル1](docs/dfd-level1.svg)

| 番号 | プロセス | 受け持ち |
|---|---|---|
| 1 | シートの情報を同期する | Excelのシートの名前、表示状態、色、選択を、0.3から3秒おきに読む (`PaneManager`) |
| 2 | 縦タブを描き、操作を受ける | 画面の描画と入力 (`TabStrip`、`TabPaneControl`) |
| 3 | 操作をExcelに反映する | 画面の操作の要求を、Excelのシートの変更にする (`PaneManager`) |
| 4 | フォルダを管理する | フォルダの構造と、シート名のルール (`FolderModel`、`NameRule`) |
| 5 | 履歴を管理する | 見たシートの履歴 (`SheetHistory`) |
| 6 | バッジを検出する | 数式エラー、メモ、保護を調べ、進捗の印を読む (`PaneManager`) |
| 7 | 設定を管理する | 設定画面、PCの設定、ブックの設定 (`SettingsForm`、`PcSettings`、`WorkbookSettings`) |
| 8 | 保存時に書き込む | ブックの保存のとき、フォルダ構成、履歴、ブックの設定を書く (`WorkbookBeforeSave`) |

### シーケンス図

動きの流れを、時間の順に示す。ここまでの図 (部品、クラス) を見てから読むと、だれが、だれを呼ぶかが、つかみやすい。

#### 基本の流れ

アドインの一生 (起動、同期、選択、保存、終了) です。まず、この4つを読んでください。

**起動、ペインの作成、同期** ([sequence-startup.puml](docs/sequence-startup.puml))

Excel の起動から、ペインを作り、同期を始めるまで。

![起動、ペインの作成、同期](docs/sequence-startup.svg)

**シートのクリックによる選択** ([sequence-select.puml](docs/sequence-select.puml))

縦タブのシートをクリックしてから、Excel のシートが切り替わるまで。

![シートのクリックによる選択](docs/sequence-select.svg)

**ブックの保存** ([sequence-save.puml](docs/sequence-save.puml))

ブックを保存するときに、設定、フォルダ、履歴を書き込む流れ。

![ブックの保存](docs/sequence-save.svg)

**ウィンドウとブックの終了処理、アドインの停止** ([sequence-shutdown.puml](docs/sequence-shutdown.puml))

ブックやウィンドウを閉じたとき、アドインを止めるときの後始末。

![ウィンドウとブックの終了処理、アドインの停止](docs/sequence-shutdown.svg)

#### 縦タブの操作

利用者が縦タブで行う操作が、どう処理されるかです。

**シートの操作 (+・挿入・名前の変更・移動とコピー・非表示と再表示)** ([sequence-sheetops.puml](docs/sequence-sheetops.puml))

「+」、挿入、名前の変更、移動・コピー、非表示と再表示。

![シートの操作 (+・挿入・名前の変更・移動とコピー・非表示と再表示)](docs/sequence-sheetops.svg)

**フォルダの操作 (開閉・作成・名前の変更・色・削除)** ([sequence-folder.puml](docs/sequence-folder.puml))

フォルダの作成、開閉、名前の変更、色、削除。

![フォルダの操作 (開閉・作成・名前の変更・色・削除)](docs/sequence-folder.svg)

**フォルダの出し入れ・移動・コピー (ドラッグ)** ([sequence-folder-drop.puml](docs/sequence-folder-drop.puml))

ドラッグでの、フォルダへの出し入れ、フォルダの移動とコピー。

![フォルダの出し入れ・移動・コピー (ドラッグ)](docs/sequence-folder-drop.svg)

**シートの行をセルへドラッグしてリンクを作る** ([sequence-link.puml](docs/sequence-link.puml))

シートの行を Excel のセルにドロップして、ハイパーリンクを作る。

![シートの行をセルへドラッグしてリンクを作る](docs/sequence-link.svg)

**検索** ([sequence-search.puml](docs/sequence-search.puml))

検索ボックスでの絞り込みと、移動。

![検索](docs/sequence-search.svg)

**履歴 (記録、戻る・進む、一覧)** ([sequence-history.puml](docs/sequence-history.puml))

見たシートの記録、戻る・進む、一覧のメニュー。

![履歴 (記録、戻る・進む、一覧)](docs/sequence-history.svg)

**コピーアイコン** ([sequence-copyicon.puml](docs/sequence-copyicon.puml))

左ボタンを押したまま右ボタンも押して、名前をコピーする。

![コピーアイコン](docs/sequence-copyicon.svg)

#### 見た目・設定・状態の表示

設定、色、バッジ、シート名のルールなど、見た目と状態に関わる流れです。

**表示モードとペインの出し入れ** ([sequence-mode.puml](docs/sequence-mode.puml))

表示モード (横のみ / 両方) と、ペインの出し入れ、「×」の扱い。

![表示モードとペインの出し入れ](docs/sequence-mode.svg)

**設定画面** ([sequence-settings.puml](docs/sequence-settings.puml))

設定画面を開いてから、OK / キャンセルで反映するまで。

![設定画面](docs/sequence-settings.svg)

**シート見出しの色と「その他の色」** ([sequence-color.puml](docs/sequence-color.puml))

シート見出しの色の指定と、Excel の「色の設定」画面 (「その他の色」)。

![シート見出しの色と「その他の色」](docs/sequence-color.svg)

**バッジの検出と進捗** ([sequence-badge.puml](docs/sequence-badge.puml))

バッジ (エラー、メモ、保護、進捗) の検出と、進捗の印の操作。

![バッジの検出と進捗](docs/sequence-badge.svg)

**メモ・コメントのバッジのクリック** ([sequence-comment.puml](docs/sequence-comment.puml))

メモ・コメントのバッジをクリックして、内容を読む。

![メモ・コメントのバッジのクリック](docs/sequence-comment.svg)

**シート名のルール** ([sequence-namerule.puml](docs/sequence-namerule.puml))

シート名のルール (`A | B`) による、フォルダの構成と、名前の同期。

![シート名のルール](docs/sequence-namerule.svg)

**アニメーション (揺れ・ゴースト・追加削除・完了)** ([sequence-animation.puml](docs/sequence-animation.puml))

揺れ、ゴースト、追加・削除、完了のはじける動き。

![アニメーション (揺れ・ゴースト・追加削除・完了)](docs/sequence-animation.svg)

### 状態遷移図とアクティビティ図 (アニメーション)

アニメーションは、種類ごとに状態を持ち、起動するかどうかを、いくつもの条件で決める。シーケンス図「アニメーション」の補足として読む。

#### 状態遷移図

アニメーションは、種類ごとに独立した状態を持つ (`TabStrip`)。([state-animation.puml](docs/state-animation.puml))

![アニメーションの状態遷移](docs/state-animation.svg)

#### アクティビティ図

行の増減、揺れ、完了のはじける動きを、起動するかどうか判断する流れ。([activity-animation.puml](docs/activity-animation.puml))

![アニメーションを起動するかどうかの判断](docs/activity-animation.svg)

### 図の作り直し

構成や流れを変えたら、`docs/*.puml` を直し、`tools\make-diagrams.ps1` で、SVG を作り直します (Java、Graphviz、`tools\PlantUML\plantuml-bsd-*.jar` が要る。jar はリポジトリに含めない)。`docs/*.puml` のすべてを、SVG にします。

## ビルド

- Visual Studio 2026 (「Office/SharePoint 開発」ワークロード) で、`VerticalSheetTabs.sln` を開く
- ブランチ運用とリリースの決まりは、Agents.md の8章を参照 (feature / fix ブランチ > `develop` > `master`、タグ `vX.Y.Z`)

```powershell
# コマンドラインでのビルド (Debug)
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" VerticalSheetTabs.sln -v:minimal
```

- Debug ビルドは、ビルド時に、現在のユーザーの Excel に、アドインを登録する (VSTO の標準動作)。動作確認は、Excelを通常起動して行う (COM自動化で起動した Excel では、VSTOアドインは読み込まれない)
- C# 7.3 の制約がある (VSTOプロジェクトの既定の言語バージョン)

## リリースの手順

```mermaid
flowchart TD
    A["featureブランチで開発・実機確認"] --> B["developへ --no-ff でマージ"]
    B --> C["バージョンを上げる<br/>.csproj ApplicationVersion<br/>AssemblyInfo.cs 2か所"]
    C --> D["MSBuild -t:Publish<br/>-p:Configuration=Release"]
    D --> E["必要なら Agents.md を更新して、コミット"]
    E --> F["masterへ --no-ff でマージ"]
    F --> G["tools/make-dist-zip.ps1 -Version X.Y.Z"]
    G --> H["配布用zip をコミット"]
    H --> I["developをmasterに早送り<br/>マージ済みブランチを削除"]
    I --> J["最後のコミットにタグ vX.Y.Z"]
```

### インストーラを作る (ClickOnce)

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    VerticalSheetTabs\VerticalSheetTabs.csproj -t:Publish -p:Configuration=Release -v:minimal
```

成果物は、`VerticalSheetTabs\bin\Release\app.publish\` に作られる (`bin\` は `.gitignore` 済み)。

| ファイル | 内容 |
|---|---|
| `setup.exe` | インストーラ (必要なランタイムを導入して、アドインを登録する) |
| `VerticalSheetTabs.vsto` | アドインの配置用マニフェスト |
| `Application Files\VerticalSheetTabs_X_Y_Z_0\` | アドイン本体 (DLL とマニフェスト) |

バージョンを上げるときは、次の3か所 (2ファイル) を更新する。

- `VerticalSheetTabs\VerticalSheetTabs.csproj` の `ApplicationVersion`
- `VerticalSheetTabs\Properties\AssemblyInfo.cs` の `AssemblyVersion` と `AssemblyFileVersion`

(設定画面のバージョン表示は、`AssemblyVersion` から読む。)

Release ビルドには、`PathMap` を設定してあり、DLL に、開発PCのフォルダのパスが入らない。公開時の警告 MSB3696 (VSTOランタイムを `setup.exe` に含めない) は、Windows 11 では問題ない。

### 配布用zipを作る

```powershell
powershell -File tools\make-dist-zip.ps1 -Version 1.4.1
```

- 公開済みのインストーラの DLL の版数が、指定した版数と同じことを確かめてから、`配布用zip\VerticalSheetTabs_vX.Y.Z.zip` を作り、古いzipを削除する
- zipの中身: `setup.exe`、`VerticalSheetTabs.vsto`、`Application Files\`、`VerticalSheetTabs.cer`、`証明書を信頼する.ps1`、`証明書の信頼を取り消す.ps1`、`はじめにお読みください.txt`、`README.md`、`README.pdf` (`README.md` から、`tools\make-readme-pdf.ps1` で作る。Edge のヘッドレス印刷。GIF は先頭の1コマの静止画)、`license\`。`.cer`、`証明書を信頼する.ps1`、`証明書の信頼を取り消す.ps1`、`はじめにお読みください.txt` は、いまの配布用zipから引き継ぐ。署名の秘密鍵 (`.pfx`) は、含めない

### GitHub への push と、リリースの自動作成

- タグ (`vX.Y.Z`) を push すると、`.github/workflows/release.yml` が動き、そのタグの配布用zipを添付した、GitHub のリリースが作られる。タグのコミットに、`配布用zip/VerticalSheetTabs_vX.Y.Z.zip` が含まれている必要がある
- **一度に4つ以上のタグを push すると、GitHub は、どのタグでも、ワークフローを起動しない。** タグは、1回に3つまでにする (最新のタグだけを push するのが、確実)
- 起動しなかったときは、Actions の「Release Artifact」を開き、「Run workflow」で、タグの名前 (例: `v1.1.71`) を入れて、手動で実行する

## 署名の証明書について

マニフェストは、自己署名の証明書 (`VerticalSheetTabs_TemporaryKey.pfx`、発行先 `CN=VerticalSheetTabs`) で署名している。この証明書は、リポジトリには含まれない (`.gitignore` 済み)。

- インストールするPCが、証明書を信頼していないと、インストール時に、「発行元を確認できません」という警告が出て、インストールできないことがある
- 他のPCに配るときは、証明書の公開部分 (`.cer`) を、そのPCの「信頼された発行元」と「信頼されたルート証明機関」に登録する (配布用zipの「証明書を信頼する.ps1」が行う)

```powershell
# 署名に使った証明書の公開部分を書き出す (署名のPCで実行)
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=VerticalSheetTabs" } | Select-Object -First 1
Export-Certificate -Cert $cert -FilePath VerticalSheetTabs.cer
```

- 別のPCで開発するときは、Visual Studio の「署名」タブで新しい証明書を作り、`.csproj` の `ManifestCertificateThumbprint` を更新する (または、`.pfx` を、安全な方法で受け渡す)
- 正式に配布する場合は、コード署名証明書を購入して、`.csproj` の `ManifestCertificateThumbprint` と `ManifestKeyFile` を差し替える

## アンインストール (開発PC)

```powershell
& "C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Uninstall <.vsto の場所> /S
```

Debug ビルドが登録した開発用のアドインは、Excelの「ファイル > オプション > アドイン」の「COMアドイン」から確認できる。

## テスト

```powershell
powershell -File tools\run-tests.ps1
```

- Excel を使わない部分のテスト (`VerticalSheetTabs.Tests`)。失敗した数が終了コードになる。リリースの前に実行する
- Excel や画面に関わる部分は、下の自動確認か、手で確かめる

## 自動確認の注意 (PowerShell + UI Automation)

- Excelを通常起動して、`System.Windows.Automation`、`mouse_event` (絶対座標、`SetProcessDPIAware` 後)、`SendKeys` で操作する
- 自分が起動したExcelだけを終了する。ほかのExcelが動いているときは、中止する
- COM経由でExcelを操作したスクリプトは、`FinalReleaseComObject` とGCのあとでないと、Excelのプロセスが残る
- Windows PowerShell 5.1 に日本語を含むスクリプトは、UTF-8 BOM付きで保存する
- ペイン内の描画、配置のコードで例外を出さない (出すと、ペインにマウス入力が届かなくなる)
