# ExcelStarch (web)

An Excel add-in that creates house-styled charts using branded colour palette and  font.
It is a partial Office JS port of my [ExcelStarch VBA macro](https://github.com/RenRMT/ExcelStarch). Some features of the macro cannot be replicated in Office JS and are not included. But this one can be used in corporate environments with more restrictive IT policies.

The add-in is hosted on GitHub Pages: <https://renrmt.github.io/excelstarch-web/>.

## Install

You don't need admin rights or Node.js.

**Excel desktop (Windows):** run the registration script from a clone of this repo, then restart Excel.

```powershell
powershell -ExecutionPolicy Bypass -File scripts\register-addin.ps1
```

The script downloads the manifest to `%LOCALAPPDATA%\ExcelStarch` and registers it for the
current user. To remove the add-in, run the script again with `-Unregister`.

**Excel on the web:** go to Home → Add-ins → More Add-ins → My Add-ins → **Upload My Add-in**,
then upload the manifest from <https://renrmt.github.io/excelstarch-web/manifest.xml>.

## Usage

The add-in adds an **ExcelStarch** group to the Home tab with two buttons.

**Chart Builder**

1. Select a data range and choose a chart type: Bar, Column, Line, Area, Scatter, Pie,
   Treemap or Box & Whisker. If you select an existing chart instead, it is restyled in place.
2. The chart is placed on a white 600 × 600 canvas, with title, subtitle and source/notes
   text boxes around it.
3. Use the **Chart text** fields to edit those text boxes. Leave a field blank to restore the
   placeholder, or enter `-` to clear it.

**Color Picker**

Select a chart, then:

- apply the data palette (Contrasting or Rainbow order)
- apply a single-colour gradient (up to 6 series) or a diverging gradient (up to 13 series)
- invert the colours
- fill a single series

Your last choices are saved in the workbook.

## Development

You need Node.js (LTS) and npm.

```sh
npm install
npm start        # dev build on https://localhost:3000, sideloaded into Excel desktop
npm stop         # stop debugging and unregister the dev add-in
npm test         # Jest unit tests (config/ and logic/ only)
npm run build    # production build into dist/
```

- **Layout:** `src/config/` (palette, fonts, geometry) and `src/logic/` are pure and covered by
  unit tests. `src/excel/` and `src/chrome/` call Office.js and are tested by sideloading; see
  [docs/testing-manual.md](docs/testing-manual.md).
- **Colours:** colours and ramps must match [reference/roos-palette.json](reference/roos-palette.json).
  A parity test enforces this.
- **Releases:** pull requests run tests and the build in CI. Merging to `main` deploys to
  GitHub Pages. Installed add-ins pick up the new version automatically.
