/* Jest configuration for the pure layer (config/ + logic/).
 * The interop layers (excel/, chrome/, export/) need a live host and are NOT unit-tested;
 * see docs/conversion-plan.md → Verification for their manual sideload steps.
 *
 * ts-jest transpiles TypeScript to CommonJS for Node — the project tsconfig targets ES
 * modules for the webpack browser build, so we override `module` here only.
 */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  roots: ["<rootDir>/test", "<rootDir>/src"],
  testMatch: ["**/*.test.ts"],
  transform: {
    "^.+\\.ts$": [
      "ts-jest",
      {
        tsconfig: {
          module: "commonjs",
        },
      },
    ],
  },
};
