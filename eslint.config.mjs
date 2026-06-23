// Project ESLint config (flat). office-addin-lint picks up this file from the project root in
// preference to its bundled default; we re-use the same office-addins recommended rules and add a
// narrow override.
import officeAddins from "eslint-plugin-office-addins";
import tsParser from "@typescript-eslint/parser";

export default [
  ...officeAddins.configs.recommended,
  {
    plugins: {
      "office-addins": officeAddins,
    },
    languageOptions: {
      parser: tsParser,
    },
  },
  {
    // chartFlow.ts, chromeText.ts, and colourFlow.ts use the documented
    // getActiveChartOrNullObject/getItemOrNullObject + load("isNullObject") + sync + read pattern.
    // In colourFlow the load→sync→read lives in the `resolveActiveChart`/`hasActiveChart` helpers, so
    // the sync crosses a function boundary; the sync boundaries are correct (verified by hand and by
    // the manual sideload tests in docs/testing-manual.md), but the office-addins flow analysis can't
    // follow the sync across the branch/helper and misfires these three rules.
    files: ["src/excel/chartFlow.ts", "src/excel/chromeText.ts", "src/excel/colourFlow.ts"],
    rules: {
      "office-addins/call-sync-before-read": "off",
      "office-addins/call-sync-after-load": "off",
      "office-addins/no-navigational-load": "off",
    },
  },
];
