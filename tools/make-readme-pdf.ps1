# README.md から PDF を作る (Edge のヘッドレス印刷)。配布用zip に入れる README.pdf の元になる。
# 使い方: powershell -File tools\make-readme-pdf.ps1 -Output C:\temp\README.pdf
# marked、mermaid、github-markdown-css は、tools\readme-pdf\vendor に同梱している (ネットワーク不要)。
# 画像は、README.md と同じ相対パス (docs\images) から読む。GIF は、先頭の1コマが、静止画として載る。
param([Parameter(Mandatory = $true)][string]$Output)
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$tool = Join-Path $PSScriptRoot "readme-pdf"
$edge = @("${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe", "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe") | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $edge) { throw "Microsoft Edge が見つかりません。" }

function ToUri([string]$p) { return ([Uri]$p).AbsoluteUri }
$md = [IO.File]::ReadAllText((Join-Path $repo "README.md"), [Text.Encoding]::UTF8)
# '</' を分けて、<script> の中の文字列が、早く閉じないようにする
$json = ($md | ConvertTo-Json -Compress).Replace("</", "<\/")
$tpl = [IO.File]::ReadAllText((Join-Path $tool "template.html"), [Text.Encoding]::UTF8)
$base = (ToUri ($repo + "\"))
$vendor = (ToUri (Join-Path $tool "vendor")).TrimEnd('/')
$html = $tpl.Replace("__BASE__", $base).Replace("__VENDOR__", $vendor).Replace("__MARKDOWN__", $json)

$work = Join-Path $env:TEMP "vst_readme_pdf"
if (Test-Path $work) { Remove-Item $work -Recurse -Force }
New-Item -ItemType Directory $work | Out-Null
$page = Join-Path $work "readme.html"
[IO.File]::WriteAllText($page, $html, (New-Object Text.UTF8Encoding($false)))

$outFull = [IO.Path]::GetFullPath($Output)
if (Test-Path $outFull) { Remove-Item $outFull }
$null = New-Item -ItemType Directory -Force (Split-Path -Parent $outFull)
$args = @("--headless=new", "--disable-gpu", "--no-pdf-header-footer", "--allow-file-access-from-files",
    "--user-data-dir=$work\profile", "--virtual-time-budget=30000", "--print-to-pdf=$outFull", (ToUri $page))
$p = Start-Process $edge -ArgumentList $args -PassThru -Wait -WindowStyle Hidden
if (-not (Test-Path $outFull)) { throw "PDF を作れませんでした (Edge の終了コード $($p.ExitCode))。" }
Remove-Item $work -Recurse -Force -ErrorAction SilentlyContinue
"PDF: $outFull ($((Get-Item $outFull).Length) バイト)"
