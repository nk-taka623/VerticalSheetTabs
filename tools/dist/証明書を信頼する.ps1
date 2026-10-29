# 縦並びシートタブの証明書 (VerticalSheetTabs.cer。公開部分のみ) を、現在のユーザーの
# 「信頼された発行元」と「信頼されたルート証明機関」に登録します。管理者権限は不要です。
# 「信頼されたルート証明機関」への登録では、Windows の確認画面が出ます。
#
# 注意: 「信頼されたルート証明機関」に入れた証明書で署名されたものは、このユーザーの Windows が信頼します。
#       この証明書は、縦並びシートタブの署名専用です。不要になったら、「証明書の信頼を取り消す.ps1」で取り消せます。
# 確認なしで登録するとき: -Force
param([switch]$Force)
$cer = Join-Path $PSScriptRoot "VerticalSheetTabs.cer"
$cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($cer)
Write-Host "登録する証明書:"
Write-Host "  発行先    : $($cert.Subject)"
Write-Host "  拇印      : $($cert.Thumbprint)"
Write-Host "  有効期限  : $($cert.NotAfter.ToString('yyyy-MM-dd'))"
if (-not $Force) {
    $answer = Read-Host "この証明書を、現在のユーザーの信頼する証明書として登録しますか? (Y/N)"
    if ($answer -notmatch '^[Yy]') { Write-Host "登録しませんでした。"; exit 1 }
}
Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\TrustedPublisher | Out-Null
Import-Certificate -FilePath $cer -CertStoreLocation Cert:\CurrentUser\Root | Out-Null
Write-Host "証明書を登録しました。続けて setup.exe を実行してください。"
