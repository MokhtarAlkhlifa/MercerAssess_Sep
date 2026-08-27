import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  build: {
    // This app bundles recharts + xlsx + a large single-file UI, so it
    // naturally crosses Vite's default 500kb warning threshold — that's a
    // heads-up, not a build failure. Splitting the heaviest libraries into
    // their own chunks below (so the browser can cache them separately from
    // app code that changes more often) is the real fix; raising the
    // threshold just stops the warning from firing on the remainder.
    chunkSizeWarningLimit: 1000,
    rollupOptions: {
      output: {
        manualChunks: {
          "vendor-react": ["react", "react-dom"],
          "vendor-charts": ["recharts"],
          "vendor-xlsx": ["xlsx"],
          "vendor-icons": ["lucide-react"],
        },
      },
    },
  },
});
