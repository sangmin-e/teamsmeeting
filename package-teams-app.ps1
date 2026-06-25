$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$teamsDir = Join-Path $root "teams-app"
$manifest = Join-Path $teamsDir "manifest.json"
$color = Join-Path $teamsDir "color.png"
$outline = Join-Path $teamsDir "outline.png"
$outZip = Join-Path $root "teams-app-package.zip"

if (!(Test-Path $manifest)) {
  throw "manifest.json not found in teams-app folder."
}

if (!(Test-Path $color) -or !(Test-Path $outline)) {
  throw "color.png and outline.png are required in teams-app folder."
}

if (Test-Path $outZip) {
  Remove-Item $outZip -Force
}

Compress-Archive -Path $manifest, $color, $outline -DestinationPath $outZip
Write-Host "Created package: $outZip"
