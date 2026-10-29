# Builds and runs the tests of the parts that do not need Excel (VerticalSheetTabs.Tests). Exit code = number of failed tests.
# Usage: powershell -File tools\run-tests.ps1
$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent $PSScriptRoot
$msbuild = "C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe"
& $msbuild (Join-Path $repo "VerticalSheetTabs.Tests\VerticalSheetTabs.Tests.csproj") -restore -v:minimal -nologo
if ($LASTEXITCODE -ne 0) { exit 1 }
& (Join-Path $repo "VerticalSheetTabs.Tests\bin\Debug\net48\VerticalSheetTabs.Tests.exe")
exit $LASTEXITCODE
