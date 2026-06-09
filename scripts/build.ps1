$ErrorActionPreference = "Stop"

$root = Resolve-Path (Join-Path $PSScriptRoot "..")
$build = Join-Path $root "build"

New-Item -ItemType Directory -Force -Path $build | Out-Null

Push-Location $root
try {
    latexmk -xelatex -interaction=nonstopmode -halt-on-error "-outdir=build" source/main.tex
}
finally {
    Pop-Location
}
