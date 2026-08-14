import { defineConfig } from "vite";

export default defineConfig({
  base: "/mongolia-weekly-price-dashboard/",
  server: {
    port: 5173,
    open: true,
  },
  build: {
    outDir: "dist",
    sourcemap: true,
  },
});
