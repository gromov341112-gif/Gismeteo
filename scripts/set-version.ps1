param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d+\.\d+(?:\.\d+)?$')]
  [string]$Version,

  [ValidateSet('stable', 'dev')]
  [string]$Channel = 'stable'
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$scriptPath = Join-Path $root 'gismeteo-excel.user.js'
$readmePath = Join-Path $root 'README.md'
$indexPath = Join-Path $root 'index.html'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

if ($Channel -eq 'dev') {
  $scriptName = 'Gismeteo Precipitation Dev'
  $namespace = 'gismeteo-excel-dev'
  $scriptVersion = "$Version-dev"
  $branch = 'dev'
} else {
  $scriptName = 'Gismeteo Precipitation'
  $namespace = 'gismeteo-excel'
  $scriptVersion = $Version
  $branch = 'main'
}

$scriptText = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)
$scriptText = $scriptText -replace '(?m)^// @name\s+.*$', "// @name         $scriptName"
$scriptText = $scriptText -replace '(?m)^// @namespace\s+.*$', "// @namespace    $namespace"
$scriptText = $scriptText -replace '(?m)^// @version\s+.*$', "// @version      $scriptVersion"
$scriptText = $scriptText -replace '(?m)^// @downloadURL\s+.*$', "// @downloadURL  https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/$branch/gismeteo-excel.user.js"
$scriptText = $scriptText -replace '(?m)^// @updateURL\s+.*$', "// @updateURL    https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/$branch/gismeteo-excel.user.js"
[System.IO.File]::WriteAllText($scriptPath, $scriptText, $utf8NoBom)

if ($Channel -eq 'stable') {
  $readmeText = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
  $readmeText = $readmeText -replace 'Current userscript version: `[^`]+`', ('Current userscript version: `{0}`' -f $Version)
  $readmeText = $readmeText -replace 'for example `[^`]+-dev`', ('for example `{0}-dev`' -f $Version)
  $readmeText = [regex]::Replace(
    $readmeText,
    '(set-version\.ps1 )\d+\.\d+(?:\.\d+)?( -Channel (?:dev|stable))',
    { param($match) $match.Groups[1].Value + $Version + $match.Groups[2].Value }
  )
  [System.IO.File]::WriteAllText($readmePath, $readmeText, $utf8NoBom)

  $indexText = [System.IO.File]::ReadAllText($indexPath, [System.Text.Encoding]::UTF8)
  $indexText = $indexText -replace 'Version [0-9]+(?:\.[0-9]+)*(?:-dev)? · Author HARIBB', "Version $Version · Author HARIBB"
  [System.IO.File]::WriteAllText($indexPath, $indexText, $utf8NoBom)
}

Write-Host "Updated $Channel version: $scriptVersion"
