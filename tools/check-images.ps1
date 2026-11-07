# README.md が使う画像ごとに、画像を最後に更新したコミットより後に、見た目に関わるソースが変わっていないかを一覧する。
# リリースの前に実行し、「変わっている」画像は、取り直すか、影響がないことを確かめる。
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root
# 画像の種類ごとに、見た目を決めるソースを対応づける (ここに無い画像は、縦タブ本体 (TabStrip など))
function Sources($name) {
    if ($name -like 'settings_*') { return 'VerticalSheetTabs/SettingsForm.cs', 'VerticalSheetTabs/SettingsBackground.cs' }
    if ($name -like 'comment_popup*') { return 'VerticalSheetTabs/CommentPopup.cs' }
    if ($name -like 'ribbon*') { return 'VerticalSheetTabs/RibbonController.cs', 'VerticalSheetTabs/RibbonIcon.cs' }
    if ($name -like 'menu_*' -or $name -like 'progress_menu*') { return 'VerticalSheetTabs/TabPaneControl.cs', 'VerticalSheetTabs/TabColorPalette.cs' }
    return 'VerticalSheetTabs/TabStrip.cs', 'VerticalSheetTabs/TabTheme.cs', 'VerticalSheetTabs/Wallpaper.cs'
}
$text = [IO.File]::ReadAllText((Join-Path $root 'README.md'), [Text.Encoding]::UTF8)
$names = [regex]::Matches($text, 'docs/images/([A-Za-z0-9_.-]+)') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique
foreach ($n in $names) {
    $path = "docs/images/$n"
    $c = git log -1 --format=%h -- $path
    $d = git log -1 --format=%ad --date=short -- $path
    $changed = @(git diff --name-only "$c..HEAD" -- (Sources $n))
    $mark = if ($changed.Count -gt 0) { '変わっている' } else { '変更なし' }
    '{0,-22} {1} {2}  {3}  {4}' -f $n, $d, $c, $mark, ($changed -join ', ')
}
