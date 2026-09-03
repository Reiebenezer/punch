import { reactRouter } from "@react-router/dev/vite";
import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "vite";
import { VitePWA } from "vite-plugin-pwa";

export default defineConfig({
  plugins: [
    tailwindcss(),
    reactRouter(),
    VitePWA({
      injectRegister: 'inline',
      registerType: "autoUpdate", // Automatically updates service worker
      manifest: {
        name: "WTN Punch",
        short_name: "Punch",
        description: "Punch your WTN shifts",
        theme_color: "#030712",
        background_color: "#030712",
        display: "standalone",
        icons: [{ src: "/favicon.ico" }],
      },
      devOptions: {
        enabled: true, // Enables PWA features in development mode
      },
    }),
  ],
  resolve: {
    tsconfigPaths: true,
  },
});
