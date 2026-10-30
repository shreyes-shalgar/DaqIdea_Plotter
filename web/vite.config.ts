import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  base: "/DaqIdea_Plotter/",
  plugins: [react()],
  server: {
    fs: {
      allow: [".."],
    },
  },
});