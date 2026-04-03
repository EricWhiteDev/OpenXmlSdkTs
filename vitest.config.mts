import { defineConfig } from "vitest/config";
import * as path from "path";

export default defineConfig({
  resolve: {
    alias: {
      openxmlsdkts: path.resolve(__dirname, "src/index.ts"),
    },
    conditions: ["import", "module", "default"],
  },
  test: {
    include: ["tests/**/*.test.ts"],
    server: {
      deps: {
        inline: ["ltxmlts", "sax"],
      },
    },
  },
});
