# 縦並びシートタブの証明書 (VerticalSheetTabs.cer) の信頼を取り消します。現在のユーザーの
# 「信頼された発行元」と「信頼されたルート証明機関」から削除します。管理者権限は不要です。
# 「信頼されたルート証明機関」からの削除では、Windows の確認画面が出ます。
# アドインを削除したあとに実行してください (実行しても、アドインのファイルは消えません)。
$cer = Join-Path $PSScriptRoot "VerticalSheetTabs.cer"
$thumb = (New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cer)).Thumbprint
foreach ($store in "TrustedPublisher", "Root") {
    $found = Get-ChildItem "Cert:\CurrentUser\$store" | Where-Object { $_.Thumbprint -eq $thumb }
    if ($found) { $found | Remove-Item; Write-Host "削除しました: $store" } else { Write-Host "登録されていません: $store" }
}
