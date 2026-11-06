# Makes the .svg of every docs\*.puml (PlantUML).
# Needs: Java, Graphviz (dot), and tools\PlantUML\plantuml-*.jar.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File tools\make-component-diagram.ps1
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
$jar = Get-ChildItem (Join-Path $PSScriptRoot "PlantUML") -Filter "plantuml*.jar" | Sort-Object Name -Descending | Select-Object -First 1
if (-not $jar) { throw "plantuml jar not found in tools\PlantUML" }

# java and dot: from PATH, or from the usual install folders (a shell opened before the install does not have them in PATH)
function Find-Exe($name, $folders) {
    $c = Get-Command $name -ErrorAction SilentlyContinue
    if ($c) { return $c.Source }
    foreach ($f in $folders) {
        $hit = Get-ChildItem $f -Recurse -Filter $name -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($hit) { return $hit.FullName }
    }
    throw "$name not found"
}
$java = Find-Exe "java.exe" @("C:\Program Files\Microsoft", "C:\Program Files\Eclipse Adoptium", "C:\Program Files\Java")
$dot = Find-Exe "dot.exe" @("C:\Program Files\Graphviz", "C:\Program Files (x86)\Graphviz")
$env:GRAPHVIZ_DOT = $dot

# every docs\*.puml (the component diagram, the sequence diagrams) becomes a .svg beside it
$sources = Get-ChildItem (Join-Path $root "docs") -Filter "*.puml" | ForEach-Object { $_.FullName }
& $java "-Dfile.encoding=UTF-8" -jar $jar.FullName -charset UTF-8 -tsvg $sources
if ($LASTEXITCODE -ne 0) { throw "plantuml failed ($LASTEXITCODE)" }
foreach ($src in $sources) {
    $out = [IO.Path]::ChangeExtension($src, ".svg")
    if (-not (Test-Path $out)) { throw "$out was not made" }
    "made: $out ($((Get-Item $out).Length) bytes)"
}
