import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Base must match your GitHub repo name so assets load correctly at
// https://<username>.github.io/<repo-name>/
// Change this if you rename the repo.
export default defineConfig({
  plugins: [react()],
  base: "/DealerInventoryScanner/",
});
