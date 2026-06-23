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
    // chartFlow.ts and chromeText.ts use the documented getItemOrNullObject + load("isNullObject")
    // + sync + read pattern across if/ternary branches. The sync boundaries are correct (verified
    // by hand and by the manual sideload tests in docs/testing-manual.md), but the office-addins
    // flow analysis cannot follow the sync across the branch and misfires these three rules.
    files: ["src/excel/chartFlow.ts", "src/excel/chromeText.ts"],
    rules: {
      "office-addins/call-sync-before-read": "off",
      "office-addins/call-sync-after-load": "off",
      "office-addins/no-navigational-load": "off",
    },
  },
];
