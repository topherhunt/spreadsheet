# Spreadsheet

A fast, fully offline spreadsheet app for macOS that opens and edits CSV, TSV, XLSX and XLSM files and feels like Google Sheets.

![Spreadsheet showing a workbook in dark mode](docs/screenshot.jpg)

## What it does

- **Opens and saves in the file's own format.** A CSV saves as a CSV, with untouched cells written back byte for byte (leading zeros, quoting and number text survive). If your edits add something a CSV can't hold (formatting, formulas, extra sheets), saving offers to write an XLSX next to it, or keep the CSV with computed values.
- **Never destroys workbook content.** Charts, pivot tables, conditional formatting, data validation, images, comments, macros and anything else the app can't show are kept on save, and a banner lists what's preserved but not active. Sheets you don't edit are written back unchanged.
- **Formulas:** 415 of Excel's 519 functions, via HyperFormula plus extras the app adds (AVERAGEIFS, LOOKUP, RANK, MODE, FORECAST, TEXTBEFORE/AFTER, XMATCH, REGEX* and others), with Excel's TEXT formatting and 1900 date system. References follow inserted/deleted rows and columns, sheet renames, copy/paste, fill and cut/paste. A formula the app can't compute (an unknown function, a 3-D reference like `Sheet1:Sheet3!A1`, or a legacy formula that relied on implicit intersection) keeps the value saved in the file, and a banner says so.
- **Sheets-style filtering and sorting**, plus an **Auto table** button that bolds the header row, freezes it and adds a filter.
- **Formatting:** number formats, fonts, colors, fills, borders, alignment, wrap, merges, and rich text with different fonts, colors and links inside one cell.
- Freeze panes, hide/unhide, resize, find and replace, multiple sheets, zoom, and light and dark themes that follow the system.

## Install

1. Download the `.dmg` from the [latest release](https://github.com/topherhunt/spreadsheet/releases/latest). It's built for Apple Silicon Macs.
2. Open the `.dmg` and drag **Spreadsheet** into **Applications**.
3. The app isn't notarized by Apple, so the first launch is blocked. Open it once and dismiss the warning. Then go to **System Settings > Privacy & Security**, scroll to the message about Spreadsheet, click **Open Anyway** and confirm. After that it opens normally.
   - Or, in Terminal: `xattr -dr com.apple.quarantine /Applications/Spreadsheet.app`
4. To make it the default app for a file type, select a `.csv` or `.xlsx` file in Finder and choose **File > Get Info**. Under **Open with**, pick Spreadsheet, then click **Change All…**.

## Develop

Requires Node 26 (see `.tool-versions`).

```sh
npm install
npm start            # Vite dev server + esbuild watch + Electron, restarting on main-process changes
npm test             # builds, then unit tests (Vitest) and E2E tests (Playwright driving Electron headless)
npm run typecheck
```

Open a file in dev with `npm start -- path/to/file.xlsx`.

## Release

`bash scripts/release.sh` cuts a release from a clean, pushed `main`: it checks that the full `npm test` passed within the hour, bumps the version, builds and signs the arm64 `.dmg` (ad-hoc), tests the packaged app, and publishes a GitHub release after you confirm. The **Build** and **Windows Tests** workflows run by hand from the Actions tab to check the macOS/Windows builds and the E2E suite on Windows.

The app icon is drawn in `build/icon.svg`; after editing it, run `npx electron scripts/render-icon.cjs` to regenerate `build/icon.png`.

## License

GPL-3.0-or-later (see `LICENSE`), as required by HyperFormula's GPLv3 license. Test fixtures in `tests/fixtures/poi` are from Apache POI (Apache 2.0).
