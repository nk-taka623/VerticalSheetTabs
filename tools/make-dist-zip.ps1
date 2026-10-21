# 配布用zipを作る: 公開済みのインストーラ (VerticalSheetTabs\bin\Release\app.publish) から、
# 配布用zip\VerticalSheetTabs_vX.Y.Z.zip を作り、古いzipを削除する (Agents.md 13章)。
# 使い方: powershell -File tools\make-dist-zip.ps1 -Version 1.1.20
# 証明書、証明書を信頼するスクリプト、手順書は、いまの配布用zipから引き継ぐ (手順書の版数は書き換える)。
param([Parameter(Mandatory = $true)][string]$Version)
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$pub = Join-Path $repo "VerticalSheetTabs\bin\Release\app.publish"
$distDir = Join-Path $repo "配布用zip"
$verU = ($Version + ".0").Replace(".", "_")   # 例: 1_1_19_0 (インストーラのフォルダ名)
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem

# アドインに埋め込まれた版数が、タグにする版数と同じことを確かめる (違うzipは作らない)
$deploy = Join-Path $pub "Application Files\VerticalSheetTabs_$verU\VerticalSheetTabs.dll.deploy"
if (-not (Test-Path $deploy)) { throw "公開済みのインストーラが、この版数ではありません: $deploy" }
$embedded = [Diagnostics.FileVersionInfo]::GetVersionInfo($deploy).FileVersion
if ($embedded -ne "$Version.0") { throw "アドインの版数 ($embedded) が、指定した版数 ($Version) と違います。版数を上げて、公開し直してください。" }

$prevZip = Get-ChildItem $distDir -Filter "VerticalSheetTabs_v*.zip" | Sort-Object LastWriteTime | Select-Object -Last 1
if (-not $prevZip) { throw "引き継ぐ元の配布用zipがありません。" }

# 作業用の場所は、パスが長すぎないように、短くする
$work = Join-Path $env:TEMP "vst_dist"
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
$old = Join-Path $work "prev"
[IO.Compression.ZipFile]::ExtractToDirectory($prevZip.FullName, $old)
$oldDir = Get-ChildItem $old -Directory | Select-Object -First 1
$new = Join-Path $work "new"
$dir = Join-Path $new "VerticalSheetTabs_v$Version"
New-Item -ItemType Directory -Force $dir | Out-Null

Copy-Item (Join-Path $pub "setup.exe"), (Join-Path $pub "VerticalSheetTabs.vsto") $dir
Copy-Item (Join-Path $pub "Application Files") $dir -Recurse
Copy-Item (Join-Path $oldDir.FullName "VerticalSheetTabs.cer") $dir
Copy-Item (Join-Path $oldDir.FullName "証明書を信頼する.ps1") $dir
Copy-Item (Join-Path $repo "README.md") $dir
$guide = [IO.File]::ReadAllText((Join-Path $oldDir.FullName "はじめにお読みください.txt"), [Text.Encoding]::UTF8)
$oldVersion = $oldDir.Name -replace "^VerticalSheetTabs_v", ""
[IO.File]::WriteAllText((Join-Path $dir "はじめにお読みください.txt"), $guide.Replace("v$oldVersion", "v$Version"), (New-Object Text.UTF8Encoding($true)))

$zip = Join-Path $distDir "VerticalSheetTabs_v$Version.zip"
if (Test-Path $zip) { Remove-Item $zip }
$fs = [IO.File]::Create($zip)
$a = New-Object IO.Compression.ZipArchive($fs, [IO.Compression.ZipArchiveMode]::Create, $false, [Text.Encoding]::UTF8)
Get-ChildItem $new -Recurse -File | ForEach-Object {
    $rel = $_.FullName.Substring($new.Length + 1).Replace('\', '/')
    [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($a, $_.FullName, $rel, [IO.Compression.CompressionLevel]::Optimal)
}
$a.Dispose(); $fs.Dispose()
if ($prevZip.FullName -ne $zip) { Remove-Item $prevZip.FullName }
Remove-Item $work -Recurse -Force

$z = [IO.Compression.ZipFile]::OpenRead($zip)
"アドインの版数: $embedded"
"zip: $zip ($($z.Entries.Count) ファイル, $((Get-Item $zip).Length) バイト)"
$z.Dispose()
