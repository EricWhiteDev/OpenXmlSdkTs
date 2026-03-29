import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    conditions: ["import", "module", "default"],
  },
  test: {
    server: {
      deps: {
        inline: ["ltxmlts", "sax"],
      },
    },
  },
});
