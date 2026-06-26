# Gismeteo Precipitation

**Gismeteo Precipitation** is a Tampermonkey userscript by **Creator: HARIBB**.
It collects 10-day precipitation forecasts from Gismeteo and exports them to a
clean Excel report.

Current userscript version: `1.3`

## What It Does

The script is built for quickly preparing a weather precipitation report by a
list of cities.

It:

- finds the exact Gismeteo city forecast page and avoids airport forecast pages;
- opens the 10-day forecast for every city from the input list;
- extracts daily precipitation values in millimeters;
- extracts weather text exactly from Gismeteo tooltips, without replacing it with guessed labels;
- builds a styled Excel workbook with city tables, precipitation charts, and clickable forecast links;
- creates a `Список` sheet with days where precipitation is above 5 mm;
- opens the `Список` sheet with the current day selected in the date filter while keeping all dates available.

## Install

Open the GitHub Pages installer:

[Gismeteo Precipitation](https://gromov341112-gif.github.io/Gismeteo/)

Installation steps:

1. Install **Tampermonkey**.
2. Copy this address, paste it into the Chrome/Yandex address bar, and press Enter:

`chrome://extensions/?id=dhdgffkkebhmkfjojejmpbldmpobfkfo`

3. Enable custom userscripts.
4. Return to the installer page and click **Установить Gismeteo Precipitation**.
5. Confirm script installation in **Tampermonkey**.

Direct stable userscript URL:

`https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js`

## Auto Update

Stable userscript auto-update is enabled only in `gismeteo-excel.user.js`:

```js
// @downloadURL  https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
// @updateURL    https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
```

The test script in `dev/` must not contain `@downloadURL` or `@updateURL`.

## Development

Development is done only in:

`dev/gismeteo-excel-dev.user.js`

This file is for manual testing and has no auto-update metadata.

After testing, publish the dev script into the stable userscript with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-dev.ps1 1.3
```

The command:

- copies tested code from `dev/gismeteo-excel-dev.user.js` to `gismeteo-excel.user.js`;
- sets stable metadata and auto-update URLs in `gismeteo-excel.user.js`;
- keeps dev metadata without auto-update URLs;
- updates the version in the stable script, dev script, README, and installer page.

## Workbook Output

The generated Excel workbook contains:

- `Осадки` - city-by-city precipitation tables and daily charts;
- `Список` - cities/days with precipitation above 5 mm;
- clickable forecast links;
- report styling with clean borders, fills, and chart blocks;
- Excel date filters with dates stored as real dates in `dd.mm.yyyy` format.

## Files

- `gismeteo-excel.user.js` - stable Tampermonkey userscript for users.
- `dev/gismeteo-excel-dev.user.js` - test userscript without auto-update.
- `index.html` - GitHub Pages installer page.
- `assets/icon.svg` - project and userscript icon.
- `.agents/AGENTS.md` - project rules for future agent sessions.
- `scripts/publish-dev.ps1` - publishes the tested dev script as a stable release.
