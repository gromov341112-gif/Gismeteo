# Gismeteo Precipitation

**Gismeteo Precipitation** is a Tampermonkey userscript by **HARIBB** for exporting
Gismeteo 10-day precipitation forecasts to a styled Excel report.

Current userscript version: `1.2`

## Install

Open the GitHub Pages installer:

[Gismeteo Precipitation](https://gromov341112-gif.github.io/Gismeteo/)

The page contains two installation buttons:

1. **Установить Tampermonkey** - opens the Tampermonkey page in Chrome Web Store.
2. **Установить Gismeteo Precipitation** - opens the direct userscript URL.

Installation steps:

1. Install Tampermonkey from Chrome Web Store.
2. Copy this address, paste it into Chrome's address bar, and press Enter:

`chrome://extensions/?id=dhdgffkkebhmkfjojejmpbldmpobfkfo`

3. Enable custom userscripts.
4. Return to the installer page and click **Установить Gismeteo Precipitation**.
5. Confirm script installation in Tampermonkey.

The installer page opens the Tampermonkey installation tab and automatically
stops its loading after a short delay so that Tampermonkey can show the script
installation confirmation.

Direct userscript URL:

`https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js`

## Auto Update

Tampermonkey updates are enabled through userscript metadata:

```js
// @downloadURL  https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
// @updateURL    https://raw.githubusercontent.com/gromov341112-gif/Gismeteo/main/gismeteo-excel.user.js
```

After a new version is pushed to `main`, installed userscripts can receive updates
through Tampermonkey's standard update mechanism.

## What It Does

- Finds the strict city forecast page on Gismeteo and avoids airport forecast pages.
- Opens the 10-day forecast for every city from the input list.
- Extracts daily precipitation values.
- Extracts weather text from Gismeteo `data-tooltip` values without replacing it with guessed text.
- Builds an Excel workbook with:
  - `Осадки` sheet for city-by-city precipitation tables;
  - large precipitation charts for each city;
  - `Список` sheet for days with precipitation above 5 mm;
  - a saved current-day filter on `Список`;
  - clickable forecast links.

## Date Rules

- Dates on `Осадки` are built sequentially from the first forecast date.
- `Список` always uses the same date set as `Осадки`.
- Dates are stored as Excel dates in `dd.mm.yyyy` format.
- The current-day filter is written into the XLSX file so the report opens with the current day selected while other dates remain available in the filter menu.

## Files

- `gismeteo-excel.user.js` - Tampermonkey userscript.
- `index.html` - GitHub Pages installer page.
- `assets/icon.svg` - project and userscript icon.
