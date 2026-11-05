import js from "@eslint/js";
import tseslint from "typescript-eslint";
import reactHooks from "eslint-plugin-react-hooks";

// What this file deliberately does NOT enable, and why:
//
// The React Compiler rule set that `eslint-plugin-react-hooks` ships with v7
// answers "can the compiler rewrite this component?" — refs read during
// render, setState inside an effect, values mutated after render. This grid is
// a canvas whose whole design is refs and imperative paint passes, so those
// answer "no" on nearly every component by construction (26 + 19 + 7 findings
// at the time this was written). They describe a code style this app does not
// use, not defects. `rules-of-hooks` and `exhaustive-deps` are the two that do
// catch real bugs and are kept.
const compilerOnly = [
  "react-hooks/static-components",
  "react-hooks/use-memo",
  "react-hooks/preserve-manual-memoization",
  "react-hooks/immutability",
  "react-hooks/globals",
  "react-hooks/refs",
  "react-hooks/set-state-in-effect",
  "react-hooks/error-boundaries",
  "react-hooks/purity",
  "react-hooks/set-state-in-render",
  "react-hooks/config",
  "react-hooks/gating",
];

export default tseslint.config(
  {
    ignores: [
      "dist",
      "release",
      "node_modules",
      "src-tauri/target",
      "dev-harness",
      "eslint.config.js",
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    plugins: { "react-hooks": reactHooks },
    languageOptions: {
      parserOptions: {
        projectService: true,
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      "react-hooks/rules-of-hooks": "error",
      // The one rule that has caught real breakage here (a stale effect that
      // never re-ran). It stays an error rather than a warning: nothing in this
      // repo is allowed to re-introduce one.
      "react-hooks/exhaustive-deps": "error",

      // Type-aware checks worth keeping: a promise dropped on the floor is a
      // silent failure, and an `any` leaking in is a hole in the strict mode
      // the rest of the project relies on.
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-unsafe-assignment": "error",
      "@typescript-eslint/no-unsafe-member-access": "error",
      "@typescript-eslint/no-unsafe-call": "error",
      "@typescript-eslint/no-unsafe-return": "error",

      // Fast feedback on dead code; no style opinions (prettier owns those).
      "no-useless-assignment": "error",
      "prefer-const": "error",
    },
  },
  ...compilerOnly.map((rule) => ({ rules: { [rule]: "off" } })),
);
