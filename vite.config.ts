import { VitePWA } from "vite-plugin-pwa";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import tailwindcss from "@tailwindcss/vite";

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    VitePWA({
      registerType: "autoUpdate",
      injectRegister: false,
      includeAssets: ['/timesheet.xlsx'],

      pwaAssets: {
        disabled: false,
        config: true,
      },

      manifest: {
        name: "WTN Punch",
        short_name: "punch",
        description: "Punch Your WTN Shifts",
        theme_color: "#030712",
      },

      workbox: {
        globPatterns: ["**/*.{js,css,html,svg,png,ico,tsx}"],
        cleanupOutdatedCaches: true,
        clientsClaim: true,
      },

      devOptions: {
        enabled: false,
        navigateFallback: "index.html",
        suppressWarnings: true,
        type: "module",
      },
    }),
  ],
});
