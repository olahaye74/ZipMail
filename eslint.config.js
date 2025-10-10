// eslint.config.js — version ESM (compatible ESLint v9 + Prettier + Office Add-in)
import js from "@eslint/js";
import globals from "globals";
import prettierPlugin from "eslint-plugin-prettier";

export default [
  js.configs.recommended,
  {
    files: ["src/**/*.js"],
    languageOptions: {
      ecmaVersion: "latest",
      sourceType: "module",
      globals: {
        ...globals.browser,
        Office: "readonly",
        OfficeRuntime: "readonly",
        console: "readonly",
        atob: "readonly",
        FileReader: "readonly",
        zip: "readonly", // pour JSZip global
      },
    },
    plugins: {
      prettier: prettierPlugin,
    },
    rules: {
      "no-console": "off",
      "prettier/prettier": "warn",
      "no-redeclare": "off", // évite les faux positifs avec les globals Office
      "@typescript-eslint/no-unused-vars": "off",
    },
  },
];
