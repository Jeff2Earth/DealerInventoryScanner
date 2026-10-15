import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Unique ID for every build. Baked into the app and written to version.json
// so open copies of the app can tell when a newer deploy exists.
const BUILD_ID = String(Date.now());

const versionFile = () => ({
  name: "emit-version-json",
  generateBundle() {
    this.emitFile({
      type: "asset",
      fileName: "version.json",
      source: JSON.stringify({ build: BUILD_ID }),
    });
  },
});

// Base must match your GitHub repo name so assets load correctly at
// https://<username>.github.io/<repo-name>/
export default defineConfig({
  plugins: [react(), versionFile()],
  base: "/DealerInventoryScanner/",
  define: { __BUILD_ID__: JSON.stringify(BUILD_ID) },
});
