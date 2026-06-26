param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^\d+\.\d+(?:\.\d+)?$')]
  [string]$Version
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$stablePath = Join-Path $root 'gismeteo-excel.user.js'
$devPath = Join-Path $root 'dev\gismeteo-excel-dev.user.js'
$readmePath = Join-Path $root 'README.md'
$indexPath = Join-Path $root 'index.html'
$utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Set-Metadata {
  param(
    [string]$Text,
    [string]$Name,
    [string]$Namespace,
    [string]$VersionValue,
    [bool]$AutoUpdate
  )

  $Text = $Text -replace '(?m)^// @name\s+.*$', "// @name         $Name"
  $Text = $Text -replace '(?m)^// @namespace\s+.*$', "// @namespace    $Namespace"
  $Text = $Text -replace '(?m)^// @version\s+.*$', "// @version      $VersionValue"
  $Text = $Text -replace '(?m)^// @author\s+.*$', '// @author       Creator: HARIBB'
  $Text = $Text -replace '(?m)^// @downloadURL\s+.*\r?\n?', ''
  $Text = $Text -replace '(?m)^// @updateURL\s+.*\r?\n?', ''

  if ($AutoUpdate) {
    $insert = "// @downloadURL  https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js`n// @updateURL    https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js"
    $Text = $Text -replace '(?m)^(// @supportURL\s+.*)$', "`$1`n$insert"
  }

  return $Text
}

$source = [System.IO.File]::ReadAllText($devPath, [System.Text.Encoding]::UTF8)

$stable = Set-Metadata `
  -Text $source `
  -Name 'Gismeteo Precipitation' `
  -Namespace 'gismeteo-excel' `
  -VersionValue $Version `
  -AutoUpdate $true

$dev = Set-Metadata `
  -Text $source `
  -Name 'Gismeteo Precipitation Dev' `
  -Namespace 'gismeteo-excel-dev' `
  -VersionValue "$Version-dev" `
  -AutoUpdate $false

[System.IO.File]::WriteAllText($stablePath, $stable, $utf8NoBom)
[System.IO.File]::WriteAllText($devPath, $dev, $utf8NoBom)

foreach ($path in @($stablePath, $devPath)) {
  $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
  if (-not $text.EndsWith("`n")) {
    [System.IO.File]::WriteAllText($path, $text + "`n", $utf8NoBom)
  }
}

$readme = [System.IO.File]::ReadAllText($readmePath, [System.Text.Encoding]::UTF8)
$readme = $readme -replace 'Current userscript version: `[^`]+`', ('Current userscript version: `{0}`' -f $Version)
$readme = [regex]::Replace(
  $readme,
  '(publish-dev\.ps1 )\d+\.\d+(?:\.\d+)?',
  { param($match) $match.Groups[1].Value + $Version }
)
[System.IO.File]::WriteAllText($readmePath, $readme, $utf8NoBom)

$index = [System.IO.File]::ReadAllText($indexPath, [System.Text.Encoding]::UTF8)
$index = $index -replace 'Version [0-9]+(?:\.[0-9]+)*(?:-dev)? · Creator: HARIBB', "Version $Version · Creator: HARIBB"
[System.IO.File]::WriteAllText($indexPath, $index, $utf8NoBom)

Write-Host "Published dev script as stable version $Version"
