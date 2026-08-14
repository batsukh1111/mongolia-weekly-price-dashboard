import { defineConfig } from "vite";
import { resolve } from "node:path";

const entry = resolve(__dirname, "index.src.html");

export default defineConfig({
  base: "/mongolia-weekly-price-dashboard/",
  server: {
    port: 5173,
    open: true,
  },
  build: {
    outDir: "dist",
    sourcemap: false,
    rollupOptions: {
      input: entry,
    },
  },
  plugins: [
    {
      name: "use-src-index-in-dev",
      configureServer(server) {
        server.middlewares.use((req, _res, next) => {
          const url = req.url || "";
          if (
            url === "/mongolia-weekly-price-dashboard/" ||
            url === "/mongolia-weekly-price-dashboard/index.html" ||
            url === "/" ||
            url === "/index.html"
          ) {
            req.url = "/mongolia-weekly-price-dashboard/index.src.html";
          }
          next();
        });
      },
    },
  ],
});
