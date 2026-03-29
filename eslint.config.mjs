import tseslint from "typescript-eslint";
import eslintConfigPrettier from "eslint-config-prettier";

export default tseslint.config(
  tseslint.configs.recommended,
  {
    files: ["**/*.ts"],
    languageOptions: {
      parserOptions: {
        project: ["./OpenXmlSdkTs/tsconfig.json", "./OpenXmlSdkTs.Test/tsconfig.json"],
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      curly: "error",
    },
  },
  eslintConfigPrettier,
  {
    ignores: ["**/dist/**", "**/node_modules/**"],
  },
);
