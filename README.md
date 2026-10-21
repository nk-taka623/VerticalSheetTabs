# 縦並びシートタブ (VerticalSheetTabs)

Excelのシートタブを、縦並びで表示する、Excel (Windows版、Microsoft 365) のアドインです。
横タブと同じ操作、同じ見た目で、縦に並べます。

開発の経緯、仕様、検証結果は、[Agents.md](Agents.md) にまとめています。

## 主な機能

- シートタブを、左または右のペインに、縦並びで表示する
- 表示モードを、ブックごとに選べる (横タブのみ / 縦タブのみ / 両方)。初回は、両方を表示し、縦タブは左
- 横タブと同じ機能: 切り替え、追加、挿入、削除、名前の変更、移動、コピー、非表示、再表示、複数選択 (グループ)、タブ色、シートの保護 など
- フォルダ (1段): シートをフォルダに分けて、縦タブに表示する。フォルダの作成 (シートの右クリック > 新しいフォルダ)、名前の変更 (ダブルクリック、F2)、色、開閉、ドラッグでの移動 (フォルダごと、複数選択、Ctrl+ドラッグでコピー)、Delキーでの削除、フォルダごとの一括操作 (フォルダの右クリック) ができる。フォルダの並びは、Excelのシートの並びにも反映される。フォルダの構成は、ブックのカスタムXMLパートに保存する (このアドインのないExcelでは、無視される)
- 設定は、ブックのカスタムドキュメントプロパティ (`VTab.Mode`、`VTab.Side`、`VTab.Width`、`VTab.Outline`、`VTab.OutlineLevel`、`VTab.NameRule`) に保存し、開くときに復元する
- 設定の変更は、リボンの「表示」タブの「縦タブ > タブ表示の設定」、または、タブの右クリックメニューの「タブ表示の設定」

## 動作環境

- Windows、Microsoft 365 の Excel (64bit で確認)
- .NET Framework 4.8、Visual Studio 2010 Tools for Office Runtime (インストーラが導入する)

## 開発

- Visual Studio 2026 (「Office/SharePoint 開発」ワークロード) で、`VerticalSheetTabs.sln` を開く
- ブランチ運用は、Agents.md の7章を参照

```powershell
# コマンドラインでのビルド (Debug)
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" VerticalSheetTabs.sln -v:minimal
```

Debug ビルドは、ビルド時に、現在のユーザーの Excel に、アドインを登録する (VSTO の標準動作)。

## 配布 (ClickOnce)

### インストーラを作る

```powershell
& "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe" `
    VerticalSheetTabs\VerticalSheetTabs.csproj -t:Publish -p:Configuration=Release -v:minimal
```

成果物は、`VerticalSheetTabs\bin\Release\app.publish\` に作られる。

| ファイル | 内容 |
|---|---|
| `setup.exe` | インストーラ (必要なランタイムを導入して、アドインを登録する) |
| `VerticalSheetTabs.vsto` | アドインの配置用マニフェスト |
| `Application Files\VerticalSheetTabs_1_1_19_0\` | アドイン本体 (DLL とマニフェスト) |

配布するときは、`app.publish` フォルダごと、共有フォルダなどに置く。

バージョンを上げるときは、`VerticalSheetTabs.csproj` の `ApplicationVersion` と、`Properties\AssemblyInfo.cs` の `AssemblyVersion` を、更新する。

### インストールする

1. `setup.exe` を実行する (Excel は、閉じておく)
2. 「Microsoft Office Customization Installer」の画面で、「インストール」を選ぶ
3. Excel を起動すると、リボンの「表示」タブに「縦タブ」グループが出て、シートタブのペインが、左に表示される

### 署名の証明書について

マニフェストは、自己署名の証明書 (`VerticalSheetTabs_TemporaryKey.pfx`) で署名している。この証明書は、リポジトリには含まれない (`.gitignore` 済み)。

- インストールするPCが、証明書を信頼していないと、インストール時に、「発行元を確認できません」という警告が出て、インストールできないことがある
- 他のPCに配るときは、証明書の公開部分 (`.cer`) を書き出して、そのPCの「信頼された発行元」と「信頼されたルート証明機関」に、登録する

```powershell
# 署名に使った証明書の公開部分を書き出す (署名のPCで実行)
$cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Subject -eq "CN=VerticalSheetTabs" } | Select-Object -First 1
Export-Certificate -Cert $cert -FilePath VerticalSheetTabs.cer

# 配布先のPCで実行 (現在のユーザーの信頼済みの証明書として登録する)
Import-Certificate -FilePath VerticalSheetTabs.cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher
Import-Certificate -FilePath VerticalSheetTabs.cer -CertStoreLocation Cert:\CurrentUser\Root
```

- 正式に配布する場合は、コード署名証明書を購入して、`.csproj` の `ManifestCertificateThumbprint` と `ManifestKeyFile` を差し替える (Visual Studio の「署名」タブから設定できる)
- 署名用の証明書は、別のPCで、開発する場合も、新しく作るか、`.pfx` を、安全な方法で、受け渡す

### アンインストールする

- Windows の「設定 > アプリ > インストールされているアプリ」から、「VerticalSheetTabs」を、アンインストールする
- コマンドで行う場合は、`"C:\Program Files (x86)\Common Files\Microsoft Shared\VSTO\10.0\VSTOInstaller.exe" /Uninstall <.vsto の場所>`

アンインストールしても、ブックに保存した設定 (`VTab.*`) は、残る。これは、ブックのデータに影響しない (Excel の「ファイル > 情報 > プロパティ > 詳細プロパティ > ユーザー設定」から、削除できる)。

## 既知の制約

- 「縦タブのみ」で保存したブックを、このアドインのない Excel で開くと、横タブが表示されない。「ファイル > オプション > 詳細設定 > このブックで作業するための表示設定 > シート見出しを表示する」で、元に戻せる
- 暗い配色 (ダークグレー、黒) は、配色の値を、実機から測っているが、縦タブの表示は、実機で確認していない
- 100% 以外の拡大率は、125% のみ確認している
- Office のテーマの切り替えは、Excel の再起動後に、縦タブへ反映される
- フォルダのあるブックでは、シート名の並びが変わるたびに、全シートの識別 (隠しカスタムプロパティ `VTab.Id`) を読み直す (300枚で約1秒)
- 詳しくは、Agents.md の 9 章、10 章、12 章を参照
