# Gismeteo Precipitation Agent Notes

## Current Project Rules

- Stable userscript for users: `gismeteo-excel.user.js`.
- Test userscript for development: `dev/gismeteo-excel-dev.user.js`.
- All new code changes must be made in `dev/gismeteo-excel-dev.user.js` first.
- Do not add `@downloadURL` or `@updateURL` to the dev userscript. The `dev/` copy is for manual testing only and must not auto-update.
- The stable userscript may include Tampermonkey auto-update metadata.
- Visible author label is `Creator:`. In userscript metadata keep the required `@author` key, but use value `Creator: HARIBB`.

## Publishing Command

After testing the dev script, publish it into the stable userscript with:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\publish-dev.ps1 1.4
```

This command:

- copies tested code from `dev/gismeteo-excel-dev.user.js` into `gismeteo-excel.user.js`;
- sets stable metadata and auto-update URLs in `gismeteo-excel.user.js`;
- keeps dev metadata without auto-update URLs;
- updates the version in stable script, dev script, README, and installer page.

Run `node --check gismeteo-excel.user.js` and `node --check dev/gismeteo-excel-dev.user.js` after publishing.
