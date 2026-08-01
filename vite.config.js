import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

// Relative base so the built app works whether it's served from
// https://<user>.github.io/  or  https://<user>.github.io/<repo-name>/
export default defineConfig({
  plugins: [react()],
  base: "./",
});
