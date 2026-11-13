# ExcelStarch (web)

An Excel add-in that creates house-styled charts using a branded colour palette and font.
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

**Requirements:** Excel on the web, or a desktop Excel that supports the ExcelApi 1.14
requirement set (current Microsoft 365 builds do; for other versions check Microsoft's
requirement-set table). Inverting fill colours needs ExcelApi 1.16.

## Usage

The add-in adds an **ExcelStarch** group to the Home tab with two buttons.

**Chart Builder**

1. Select a data range and choose a chart type:
   - **Column** and **Bar**: clustered, stacked and 100% stacked
   - **Line & Area**: line, stacked area and 100% area
   - **Pie**: pie and donut
   - **Scatter**: scatter and bubble
   - **Complex**: treemap and box & whisker

   If you select an existing chart instead, it is changed to that type and restyled. To restyle a
   chart without changing its type, select it and click **Apply style to selected chart**; this also
   works on chart types the builder does not create, as far as each type allows.
2. The chart is placed on a white 600 × 600 canvas, with title, subtitle and source/notes
   text boxes around it.
3. Use the **Chart text** fields to edit those text boxes. Leave a field blank to restore the
   placeholder, or enter `-` to clear it.
4. Use **Chart tools** to cycle gridlines, axis labels, axis titles, the legend and data labels
   on the selected chart. Data labels apply to all series or to the series you pick. **Add
   annotation** puts a text box at the plot centre or beside a point you pick.
5. Use **Export** to save the selected chart, with its title and source, as one PNG or JPEG
   image: download it, copy it, or insert it as a picture next to the chart. The chart and its
   text boxes are separate objects in Excel, so copying the chart on its own leaves the text
   behind; export is the way to take a finished chart elsewhere. SVG and PDF are not available
   to Excel add-ins.

**Color Picker**

Select a chart, then:

- apply the data palette (Contrasting or Rainbow order)
- apply a single-colour gradient (up to 6 series) or a diverging gradient (up to 13 series)
- invert the colours
- fill a single series, or a single point of a series

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
- **Colours and fonts:** all palette, ramp and font values live in `src/config/`.
- **Logo:** off by default. To add one bottom-right on every chart, set `logoEnabled` and paste a
  base64 PNG into `src/config/logo.ts`.
- **Releases:** pull requests run tests and the build in CI. Merging to `main` deploys to
  GitHub Pages. Installed add-ins pick up the new version automatically.
